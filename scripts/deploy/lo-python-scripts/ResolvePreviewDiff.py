"""
ResolvePreviewDiff.py
---------------------

Handles accept / reject operations for inline preview diffs produced by
PreviewActionPlan.py. Decisions are communicated as JSON payloads containing
``plan_id``, an optional list of ``operation_ids``, and a ``decision``
("accept" or "reject").
"""

import json
import os
import traceback
from datetime import datetime as dt
from typing import Dict, Any, List, Optional, Tuple

from com.sun.star.awt.FontWeight import (
    NORMAL as FONT_WEIGHT_NORMAL,
    BOLD as FONT_WEIGHT_BOLD,
)
from com.sun.star.awt.FontSlant import ITALIC as FONT_SLANT_ITALIC
from com.sun.star.awt.FontUnderline import SINGLE as FONT_UNDERLINE_SINGLE
from com.sun.star.style.ParagraphAdjust import (
    LEFT as PAR_ADJUST_LEFT,
    RIGHT as PAR_ADJUST_RIGHT,
    CENTER as PAR_ADJUST_CENTER,
    BLOCK as PAR_ADJUST_BLOCK,
)
from com.sun.star.text.ControlCharacter import LINE_BREAK, PARAGRAPH_BREAK
from com.sun.star.util import DateTime as UnoDateTime

try:  # LibreOffice runtime provides this module; add fallback for testing.
    from com.sun.star.uno import Exception as UNOException
except ImportError:  # pragma: no cover - fallback when UNO not available
    UNOException = Exception  # type: ignore

PT_TO_HMM = 2540.0 / 72.0
LOG_FILE = "/tmp/sdoc_preview.log"
PREVIEW_STATE_KEY = "sdoc_preview_state"
PREVIEW_BOOKMARK_PREFIXES = ("ov_preview_", "sdoc_preview_")
PREVIEW_GREEN_RGB = 0x008A00
DEBUG_LOG = os.environ.get("SDOC_PREVIEW_DEBUG", "").lower() in {"1", "true", "yes", "debug"}
SUMMARY_LOG_PREFIXES = (
    "ResolvePreviewDiff: raw payload",
    "ResolvePreviewDiff: payload parsed",
    "ResolvePreviewDiff: decision=",
    "ResolvePreviewDiff: host_managed_save",
    "ResolvePreviewDiff: successful=",
    "ResolvePreviewDiff: preview bookmark purge",
    "ResolvePreviewDiff: preview green formatting scrub",
    "ResolvePreviewDiff: controller restored",
    "ResolvePreviewDiff: Document consistency",
    "ResolvePreviewDiff: Triggering WOPI save",
    "ResolvePreviewDiff: WOPI save dispatch",
    "ResolvePreviewDiff: fallback",
)


def _should_log(message: str) -> bool:
    if DEBUG_LOG:
        return True
    lowered = message.lower()
    if "failed" in lowered or "error" in lowered or "warning" in lowered:
        return True
    return message.startswith(SUMMARY_LOG_PREFIXES)


def _log(message: str) -> None:
    if not _should_log(message):
        return
    timestamped = f"{dt.now().isoformat(timespec='milliseconds')} {message}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as handle:
            handle.write(timestamped + "\n")
    except Exception:
        pass
    try:
        print(timestamped, flush=True)
    except UnicodeEncodeError:
        safe_message = timestamped.encode("ascii", "backslashreplace").decode("ascii")
        print(safe_message, flush=True)


def _get_text_columns(owner):
    """Return a mutable XTextColumns object for a page style or section."""
    last_exc = None

    try:
        columns = owner.TextColumns
        if columns is not None:
            return columns
    except Exception as exc:
        last_exc = exc

    try:
        context = XSCRIPTCONTEXT.getComponentContext()  # type: ignore  # NOQA
        service_manager = getattr(context, "ServiceManager", None)
        if service_manager is None:
            service_manager = context.getServiceManager()
        columns = service_manager.createInstanceWithContext(
            "com.sun.star.text.TextColumns", context
        )
        if columns is not None:
            return columns
    except Exception as exc:
        last_exc = exc

    raise RuntimeError(f"Could not obtain TextColumns object: {last_exc}")


def _set_text_columns_property(columns, name: str, value: Any) -> bool:
    try:
        setattr(columns, name, value)
        return True
    except Exception:
        return False


def _configure_text_columns(
    columns,
    column_count: int,
    spacing_mm: Optional[float] = None,
    separator_line: Optional[bool] = None,
):
    """Mutate and return an existing XTextColumns object."""
    count = max(1, int(column_count or 1))
    spacing = float(spacing_mm if spacing_mm is not None else 5.0)
    distance = int(spacing * 100)

    set_column_count = getattr(columns, "setColumnCount", None)
    if callable(set_column_count):
        set_column_count(count)
    else:
        columns.ColumnCount = count

    _set_text_columns_property(columns, "AutomaticDistance", distance if count > 1 else 0)

    if separator_line is not None:
        for prop_name in ("SeparatorLineIsOn", "SeparatorLineIsVisible"):
            if _set_text_columns_property(columns, prop_name, bool(separator_line)):
                break

    return columns


def resolvePreviewDiff(payload_json: str) -> str:
    """
    Accepts a JSON payload describing the user's decision for one or more pending
    preview actions and applies or rolls back the preview content in the document.
    """
    def _result(ok: bool, **fields: Any) -> str:
        base: Dict[str, Any] = {"ok": ok, "event": "word_preview_resolve"}
        base.update(fields)
        try:
            return json.dumps(base)
        except Exception:
            # Must always return a string to Collabora.
            return "{\"ok\": false, \"event\": \"word_preview_resolve\", \"error\": \"failed to serialize result\"}"

    try:
        _log("ResolvePreviewDiff: raw payload received.")
        payload = json.loads(payload_json)
        _log(f"ResolvePreviewDiff: payload parsed. keys={list(payload.keys())}")
    except Exception as exc:
        _log(
            f"ResolvePreviewDiff: failed to parse payload: {exc}\n{traceback.format_exc()}"
        )
        return _result(False, error=f"failed to parse payload: {exc}")

    plan_id = payload.get("plan_id")
    decision = payload.get("decision")
    if decision not in ("accept", "reject"):
        _log(f"ResolvePreviewDiff: invalid decision value: {decision}")
        return _result(False, plan_id=plan_id, decision=decision, error="invalid decision")

    document = XSCRIPTCONTEXT.getDocument()  # type: ignore  # NOQA
    if document is None:
        _log("ResolvePreviewDiff: no active document.")
        return _result(False, plan_id=plan_id, decision=decision, error="no active document")

    state = _load_preview_state(document)
    plan_state = state.get(plan_id, {})
    if not plan_state:
        _log(f"ResolvePreviewDiff: no preview state for plan {plan_id}.")
        return _result(False, plan_id=plan_id, decision=decision, error="no preview state for plan")

    operation_ids = payload.get("operation_ids") or list(plan_state.keys())
    _log(
        f"ResolvePreviewDiff: decision={decision} plan_id={plan_id} operations={len(operation_ids)}"
    )
    host_managed_save = bool(payload.get("host_managed_save", False))
    if host_managed_save:
        _log("ResolvePreviewDiff: host_managed_save=true (skipping script-managed save)")

    doc = document
    controller = doc.getCurrentController()

    # Silent-resolve handling:
    # Preview colors are formatting only (green/gold), not tracked changes.
    # Accept/reject resolves content with tracking paused, then restores the
    # prior recording state. Redline visibility remains hidden.
    was_recording = getattr(doc, "RecordChanges", False)
    original_display_type = getattr(doc, "RedlineDisplayType", 0)

    try:
        # Force tracking OFF during cleanup/normalization phase
        doc.RecordChanges = False
        _log(f"ResolvePreviewDiff: paused RecordChanges (was {was_recording})")
    except Exception as e:
        _log(f"ResolvePreviewDiff: failed to pause RecordChanges: {e}")

    try:
        # Force "Show Changes" OFF during resolution to prevent interference
        # 0 = NONE (hidden)
        doc.RedlineDisplayType = 0
        _log(f"ResolvePreviewDiff: set RedlineDisplayType to 0 (was {original_display_type})")
    except Exception as e:
        _log(f"ResolvePreviewDiff: failed to set RedlineDisplayType: {e}")

    try:
        doc.lockControllers()
        _log("ResolvePreviewDiff: controllers locked.")
    except Exception as exc:
        _log(
            f"ResolvePreviewDiff: failed to lock controllers: {exc}\n{traceback.format_exc()}"
        )
        return _result(False, plan_id=plan_id, decision=decision, error=f"failed to lock controllers: {exc}")

    successful_ops = []
    failed_ops = []

    try:
        for op_id in operation_ids:
            op_state = plan_state.get(op_id)
            if not op_state:
                cleaned = _fallback_resolve_without_state(doc, plan_id, op_id, decision)
                _log(
                    "ResolvePreviewDiff: operation "
                    f"{op_id} not found in state; fallback cleanup matched={cleaned}."
                )
                failed_ops.append(op_id)
                plan_state[op_id] = {
                    "status": "failed",
                    "decision": decision,
                    "error": "operation not found in state",
                    "fallback_cleanup_matches": cleaned,
                    "failed_at": str(dt.now()),
                }
                continue

            # Skip if already resolved
            if op_state.get("status") == "resolved":
                _log(
                    f"ResolvePreviewDiff: operation {op_id} already resolved; skipping."
                )
                continue

            try:
                if decision == "accept":
                    _log(f"ResolvePreviewDiff: accepting {op_id}")
                    _accept_preview(doc, plan_id, op_id, op_state)
                else:
                    _log(f"ResolvePreviewDiff: rejecting {op_id}")
                    _reject_preview(doc, plan_id, op_id, op_state)
                # Mark as resolved instead of deleting, so PreviewActionPlan knows to skip it on re-runs
                plan_state[op_id] = {
                    "status": "resolved",
                    "decision": decision,
                    "resolved_at": str(dt.now()),
                }
                successful_ops.append(op_id)
                _log(
                    f"ResolvePreviewDiff: operation {op_id} resolved and marked in state."
                )
            except Exception as op_exc:
                cleaned = _fallback_resolve_without_state(doc, plan_id, op_id, decision)
                failed_ops.append(op_id)
                # Mark as failed but keep in state for debugging
                plan_state[op_id] = {
                    "status": "failed",
                    "decision": decision,
                    "error": str(op_exc),
                    "fallback_cleanup_matches": cleaned,
                    "failed_at": str(dt.now()),
                }
                _log(
                    f"ResolvePreviewDiff: failed to resolve {op_id}: {op_exc}\n{traceback.format_exc()}"
                )

        # Count only unresolved operations (not resolved or failed)
        unresolved_ops = [
            k
            for k, v in plan_state.items()
            if v.get("status") not in ("resolved", "failed")
        ]

        if plan_state:
            state[plan_id] = plan_state
            _log(
                f"ResolvePreviewDiff: successful={len(successful_ops)}, failed={len(failed_ops)}, remaining unresolved={len(unresolved_ops)}"
            )
        else:
            state.pop(plan_id, None)
        _save_preview_state(doc, state)

        # Hard guarantee: preview bookmarks must not survive resolve.
        purged_count, purge_failed = _dispose_stale_preview_bookmarks(
            doc,
            plan_id=plan_id,
            operation_ids=[
                op_id for op_id in operation_ids if isinstance(op_id, str) and op_id
            ],
        )
        _log(
            "ResolvePreviewDiff: preview bookmark purge complete "
            f"removed={purged_count} failed={purge_failed}"
        )

        preview_green_scrubbed = 0
        if decision == "accept":
            preview_green_scrubbed = _scrub_preview_green_formatting(doc)
            _log(
                "ResolvePreviewDiff: preview green formatting scrub complete "
                f"changed={preview_green_scrubbed}"
            )

        # If there were failures, log warning but continue
        if failed_ops:
            _log(
                f"ResolvePreviewDiff: WARNING - {len(failed_ops)} operations failed: {failed_ops}"
            )

    finally:
        # Restore RecordChanges state robustly
        try:
            doc.RecordChanges = was_recording
            _log(f"ResolvePreviewDiff: Restored RecordChanges to {was_recording}.")
        except Exception as restore_exc:
            _log(f"ResolvePreviewDiff: Failed to restore RecordChanges: {restore_exc}")

        # Keep redlines hidden after resolve. Visibility should only be changed
        # via explicit user toggle (SetShowChanges).
        try:
            doc.RedlineDisplayType = 0
            _log("ResolvePreviewDiff: Enforced RedlineDisplayType=0 after resolve.")
        except Exception as restore_display_exc:
            _log(
                f"ResolvePreviewDiff: Failed to enforce RedlineDisplayType=0: {restore_display_exc}"
            )

        doc.unlockControllers()
        if controller is not None:
            try:
                controller.restoreViewData(controller.getViewData())
                _log("ResolvePreviewDiff: controller restored.")
            except Exception:
                pass

    if not host_managed_save:
        # Check document consistency before saving
        if not _check_document_consistency(doc, failed_ops):
            _log("ResolvePreviewDiff: Document consistency check failed; save may fail.")

        # Trigger WOPI-aware save using UNO dispatch (not just doc.store()).
        # This sends PutFile to the WOPI server, unlike doc.store() which only saves locally.
        # Always attempt a save so preview resolution never requires manual save to sync.
        if failed_ops and not successful_ops:
            _log(
                "ResolvePreviewDiff: no successful operations, but forcing save for sync."
            )

        try:
            frame = doc.getCurrentController().getFrame()
            dispatcher = XSCRIPTCONTEXT.getComponentContext().ServiceManager.createInstanceWithContext(  # type: ignore  # NOQA
                "com.sun.star.frame.DispatchHelper",
                XSCRIPTCONTEXT.getComponentContext(),  # type: ignore  # NOQA
            )
            _log("ResolvePreviewDiff: Triggering WOPI save via .uno:Save dispatch.")
            dispatcher.executeDispatch(frame, ".uno:Save", "", 0, ())
            _log("ResolvePreviewDiff: WOPI save dispatch executed.")
        except Exception as save_error:
            _log(
                f"ResolvePreviewDiff: WOPI save dispatch failed, falling back to doc.store(): {save_error}"
            )
            try:
                doc.store()
                _log("ResolvePreviewDiff: fallback doc.store() completed.")
            except Exception as fallback_error:
                _log(f"ResolvePreviewDiff: fallback save also failed: {fallback_error}")

    operation_results = [
        {
            "operation_id": op_id,
            "status": "ok",
            "decision": decision,
        }
        for op_id in successful_ops
    ]
    operation_results.extend(
        {
            "operation_id": op_id,
            "status": "error",
            "decision": decision,
        }
        for op_id in failed_ops
    )

    return _result(
        True,
        plan_id=plan_id,
        decision=decision,
        host_managed_save=host_managed_save,
        results=operation_results,
        all_done=len(failed_ops) == 0,
        successful_operation_ids=successful_ops,
        failed_operation_ids=failed_ops,
        successful_ops=len(successful_ops),
        failed_ops=len(failed_ops),
        preview_green_scrubbed=preview_green_scrubbed if decision == "accept" else 0,
    )


def _check_document_consistency(doc, failed_ops: List[str]) -> bool:
    """
    Check if document is in a consistent state after operations.
    Returns True if consistent, False otherwise.
    """
    try:
        # Basic checks: document is accessible and not in error state
        if doc is None:
            _log("_check_document_consistency: document is None")
            return False

        # Check if document can access text
        try:
            text = doc.getText()
            if text is None:
                _log("_check_document_consistency: document text is None")
                return False
        except Exception as e:
            _log(f"_check_document_consistency: failed to access document text: {e}")
            return False

        # If there were failed operations, document might be inconsistent
        if failed_ops:
            _log(
                f"_check_document_consistency: {len(failed_ops)} operations failed - document may be inconsistent"
            )
            # Don't return False here, as partial success might be acceptable
            # But log the warning

        return True
    except Exception as e:
        _log(f"_check_document_consistency: error during consistency check: {e}")
        return False


def _user_props_has_property(props, name: str) -> bool:
    getter = getattr(props, "getPropertySetInfo", None)
    if getter is not None:
        try:
            info = getter()
            if info and info.hasPropertyByName(name):
                return True
        except Exception:
            pass
    try:
        props.getPropertyValue(name)
        return True
    except Exception:
        return False


def _load_preview_state(doc) -> Dict[str, Dict[str, Any]]:
    props = doc.getDocumentProperties().getUserDefinedProperties()
    if not _user_props_has_property(props, PREVIEW_STATE_KEY):
        return {}
    try:
        raw = props.getPropertyValue(PREVIEW_STATE_KEY)
        return json.loads(raw) if raw else {}
    except Exception:
        return {}


def _save_preview_state(doc, state: Dict[str, Dict[str, Any]]) -> None:
    props = doc.getDocumentProperties().getUserDefinedProperties()
    serialized = json.dumps(state)
    if _user_props_has_property(props, PREVIEW_STATE_KEY):
        props.setPropertyValue(PREVIEW_STATE_KEY, serialized)
    else:
        props.addProperty(PREVIEW_STATE_KEY, 0, serialized)


def _preview_bookmark_prefix(plan_id: Optional[str], op_id: Optional[str] = None) -> str:
    if not isinstance(plan_id, str) or not plan_id:
        return ""
    base = f"ov_preview_{plan_id}_".replace("-", "_")
    if isinstance(op_id, str) and op_id:
        return f"{base}{op_id}_".replace("-", "_")
    return base


def _get_preview_bookmark_names_for_op(
    doc, plan_id: Optional[str], op_id: Optional[str]
) -> List[str]:
    prefix = _preview_bookmark_prefix(plan_id, op_id)
    if not prefix:
        return []
    try:
        bookmarks = doc.getBookmarks()
        names = list(bookmarks.getElementNames())
    except Exception:
        return []
    return [
        name
        for name in names
        if isinstance(name, str)
        and any(name.startswith(p) for p in PREVIEW_BOOKMARK_PREFIXES)
        and name.startswith(prefix)
    ]


def _fallback_resolve_without_state(
    doc, plan_id: Optional[str], op_id: Optional[str], decision: str
) -> int:
    """
    Best-effort cleanup when operation state is missing/corrupt.
    Cleans preview artifacts by deterministic bookmark naming.
    Returns the number of matching preview bookmarks encountered.
    """
    names = _get_preview_bookmark_names_for_op(doc, plan_id, op_id)
    if not names:
        return 0

    original_ids = [n for n in names if "_original" in n]
    proposed_ids = [n for n in names if "_proposed" in n]
    separator_ids = [n for n in names if "_separator" in n]

    classified = set(original_ids + proposed_ids + separator_ids)
    other_ids = [n for n in names if n not in classified]

    if decision == "accept":
        _remove_bookmarked_segments(doc, original_ids + separator_ids)
        _normalize_green_segments(doc, proposed_ids)
        _normalize_table_segments(doc, proposed_ids)
    else:
        _remove_bookmarked_segments(doc, proposed_ids + separator_ids)
        _normalize_green_segments(doc, original_ids)
        _normalize_table_segments(doc, original_ids)

    if other_ids:
        _normalize_green_segments(doc, other_ids)
        _normalize_table_segments(doc, other_ids)

    # Hard-finalizer for this op: remove any remaining preview bookmarks.
    remaining = _get_preview_bookmark_names_for_op(doc, plan_id, op_id)
    if remaining:
        _dispose_bookmarks(doc, remaining)

    return len(names)


def _dispose_stale_preview_bookmarks(
    doc, plan_id: Optional[str] = None, operation_ids: Optional[List[str]] = None
) -> Tuple[int, int]:
    """
    Best-effort cleanup for any preview bookmarks that may remain after resolve.
    This ensures preview-only markers never survive into WOPI save snapshots.
    """
    removed = 0
    failed = 0
    try:
        bookmarks = doc.getBookmarks()
        names = list(bookmarks.getElementNames())
    except Exception as exc:
        _log(f"_dispose_stale_preview_bookmarks: failed to access bookmarks: {exc}")
        return removed, failed

    plan_prefix = _preview_bookmark_prefix(plan_id)
    op_prefixes: List[str] = []
    if isinstance(plan_id, str) and plan_id and operation_ids:
        for op_id in operation_ids:
            if isinstance(op_id, str) and op_id:
                op_prefixes.append(_preview_bookmark_prefix(plan_id, op_id))

    for name in names:
        if not isinstance(name, str):
            continue
        if not any(name.startswith(prefix) for prefix in PREVIEW_BOOKMARK_PREFIXES):
            continue
        if plan_prefix:
            if op_prefixes:
                if not any(name.startswith(prefix) for prefix in op_prefixes):
                    continue
            elif not name.startswith(plan_prefix):
                continue
        try:
            if bookmarks.hasByName(name):
                bookmarks.getByName(name).dispose()
                removed += 1
        except Exception as exc:
            failed += 1
            _log(
                f"_dispose_stale_preview_bookmarks: failed to dispose {name}: {exc}"
            )

    return removed, failed


def _get_element_by_bookmark(doc, bookmark_name: Optional[str]):
    if not bookmark_name:
        return None
    try:
        bookmarks = doc.getBookmarks()
        if not bookmarks.hasByName(bookmark_name):
            return None
        bookmark = bookmarks.getByName(bookmark_name)
        anchor = bookmark.getAnchor()
        doc_text = doc.getText()
        enum = doc_text.createEnumeration()
        while enum.hasMoreElements():
            element = enum.nextElement()
            try:
                if element.supportsService("com.sun.star.text.TextTable"):
                    table_anchor = element.getAnchor()
                    try:
                        if doc_text.compareRegionStarts(table_anchor, anchor) == 0:
                            return ("table", element)
                    except Exception:
                        if table_anchor == anchor:
                            return ("table", element)
                elif element.supportsService("com.sun.star.text.Paragraph"):
                    par_start = element.getStart()
                    try:
                        if doc_text.compareRegionStarts(par_start, anchor) == 0:
                            return ("paragraph", element)
                    except Exception:
                        if par_start == anchor:
                            return ("paragraph", element)
            except UNOException:
                continue
    except Exception:
        return None
    return None


def _get_element_by_index(doc, index: int):
    if index < 0:
        return None
    enum = doc.getText().createEnumeration()
    position = 0
    while enum.hasMoreElements():
        element = enum.nextElement()
        try:
            if element.supportsService("com.sun.star.text.TextTable"):
                if position == index:
                    return ("table", element)
                position += 1
            elif element.supportsService("com.sun.star.text.Paragraph"):
                if position == index:
                    return ("paragraph", element)
                position += 1
            else:
                position += 1
        except UNOException:
            position += 1
    return None


def _write_paragraph_spec(
    cursor,
    spec: Dict[str, Any],
    color: Optional[int] = None,
    append_break: bool = True,
) -> List[Tuple[Any, Any, Any]]:
    text = cursor.getText()
    ranges: List[Tuple[Any, Any, Any]] = []
    run_segments: List[Tuple[Any, Any, Dict[str, Any]]] = []

    runs = spec.get("runs")
    full_text = spec.get("text")
    format_spec = spec.get("format") or {}

    if runs:
        new_cursor = text.createTextCursorByRange(cursor)
        new_cursor.collapseToStart()
        try:
            new_cursor.CharStrikeout = 0
            new_cursor.CharUnderline = 0
        except Exception:
            pass
        start_range = new_cursor.getStart()

        for run in runs:
            if run.get("break_before") == "line":
                text.insertControlCharacter(new_cursor, LINE_BREAK, False)
            run_text = run.get("text") or ""
            run_start = new_cursor.getStart()
            text.insertString(new_cursor, run_text, False)
            run_cursor = text.createTextCursorByRange(run_start)
            run_cursor.gotoRange(new_cursor.getStart(), True)
            try:
                run_cursor.CharStrikeout = 0
                run_cursor.CharUnderline = 0
            except Exception:
                pass
            run_format = run.get("format") or {}
            _apply_run_formatting(run_cursor, run_format, color)
            run_segments.append((run_start, run_cursor.getEnd(), run_format))
            
            # Handle hyperlinks - check for hyperlink property on the run
            hyperlink_spec = run.get("hyperlink")
            if hyperlink_spec:
                target_url = hyperlink_spec.get("target") or ""
                if target_url:
                    is_external = hyperlink_spec.get("external", True)
                    target_frame = "_blank" if is_external else ""
                    try:
                        # LibreOffice UNO uses HyperLinkURL (capital L), not HyperlinkURL
                        if hasattr(run_cursor, "setPropertyValue"):
                            run_cursor.setPropertyValue("HyperLinkURL", target_url)
                            run_cursor.setPropertyValue("HyperLinkName", "")
                            run_cursor.setPropertyValue("HyperLinkTarget", target_frame)
                        else:
                            run_cursor.HyperLinkURL = target_url
                            run_cursor.HyperLinkName = ""
                            run_cursor.HyperLinkTarget = target_frame
                    except Exception:
                        pass

            # Explicitly move new_cursor to the end of the inserted run to ensure next run appends
            new_cursor = text.createTextCursorByRange(run_cursor.getEnd())
            new_cursor.collapseToEnd()

        if append_break:
            text.insertControlCharacter(new_cursor, PARAGRAPH_BREAK, False)
            end_range = new_cursor.getStart()
            new_cursor.goLeft(1, False)
        else:
            end_range = new_cursor.getStart()
        ranges.append((text, start_range, end_range))
    elif full_text:
        new_cursor = text.createTextCursorByRange(cursor)
        new_cursor.collapseToStart()
        try:
            new_cursor.CharStrikeout = 0
            new_cursor.CharUnderline = 0
        except Exception:
            pass
        start_range = new_cursor.getStart()
        text.insertString(new_cursor, full_text, False)
        if color is not None:
            rng = text.createTextCursorByRange(start_range)
            rng.gotoRange(new_cursor.getStart(), True)
            try:
                rng.CharStrikeout = 0
                rng.CharUnderline = 0
            except Exception:
                pass
            rng.CharColor = color
        if append_break:
            text.insertControlCharacter(new_cursor, PARAGRAPH_BREAK, False)
            end_range = new_cursor.getStart()
            new_cursor.goLeft(1, False)
        else:
            end_range = new_cursor.getStart()
        ranges.append((text, start_range, end_range))

    if ranges and format_spec:
        for text_obj, start_range, end_range in ranges:
            try:
                para_cursor = text_obj.createTextCursorByRange(start_range)
                para_cursor.gotoRange(end_range, True)
                _apply_paragraph_formatting(para_cursor, format_spec)
            except Exception as exc:
                _log(f"_write_paragraph_spec: failed to apply paragraph formatting: {exc}")

    if run_segments:
        for run_start, run_end, run_format in run_segments:
            try:
                run_cursor = text.createTextCursorByRange(run_start)
                run_cursor.gotoRange(run_end, True)
                _apply_run_formatting(run_cursor, run_format, color)
            except Exception as exc:
                _log(f"_write_paragraph_spec: failed to re-apply run formatting: {exc}")

    return ranges


def _apply_run_formatting(
    cursor, fmt: Dict[str, Any], override_color: Optional[int]
) -> None:
    if override_color is not None:
        cursor.CharColor = override_color
    else:
        font_color = fmt.get("font_color_rgb")
        if font_color:
            try:
                cursor.CharColor = int(str(font_color).lstrip("#"), 16)
            except ValueError:
                pass
    if fmt.get("bold"):
        cursor.CharWeight = FONT_WEIGHT_BOLD
    if fmt.get("italic"):
        cursor.CharPosture = FONT_SLANT_ITALIC
    if fmt.get("underline"):
        cursor.CharUnderline = FONT_UNDERLINE_SINGLE
    if fmt.get("font_name"):
        cursor.CharFontName = fmt["font_name"]
    if fmt.get("font_size_pt"):
        cursor.CharHeight = fmt["font_size_pt"]


def _accept_preview(doc, plan_id: str, op_id: str, op_state: Dict[str, Any]) -> None:
    # Always clean up separators (untracked)
    _remove_bookmarked_segments(doc, op_state.get("separator_ids") or [])

    op_type = op_state.get("type")
    if op_type in ("paragraph_replacement", "run_replacement"):
        _log(f"_accept_paragraph_replacement: op_id={op_id}")
        _accept_paragraph_replacement(doc, plan_id, op_id, op_state)
    elif op_type == "paragraph_insertion":
        _log(f"_accept_paragraph_insertion: op_id={op_id}")
        _accept_paragraph_insertion(doc, plan_id, op_id, op_state)
    elif op_type == "delete_paragraph":
        _log(f"_accept_delete_paragraph: op_id={op_id}")
        _accept_delete_paragraph(doc, plan_id, op_id, op_state)
    elif op_type in ("table_replacement", "table_insertion"):
        _log(f"_accept_table_change: op_id={op_id}")
        _accept_table_change(doc, plan_id, op_id, op_state)
    elif op_type == "delete_table":
        _log(f"_accept_delete_table: op_id={op_id}")
        _accept_delete_table(doc, plan_id, op_id, op_state)
    elif op_type == "paragraph_style":
        _log(f"_accept_style_change: op_id={op_id}")
        _accept_style_change(doc, plan_id, op_id, op_state)
    elif op_type == "document_style":
        _log(f"_apply_document_style: op_id={op_id}")
        _apply_document_style(doc, op_state.get("update") or {})
    # Chart operations
    elif op_type in ("chart_insertion", "chart_replacement"):
        _log(f"_accept_chart_change: op_id={op_id}")
        _accept_chart_change(doc, plan_id, op_id, op_state)
    elif op_type == "chart_data_update":
        _log(f"_accept_chart_data_update: op_id={op_id}")
        # Data update is already applied; just log
        pass
    # Header/footer operations
    elif op_type in ("header_change", "header_clear"):
        _log(f"_accept_header_change: op_id={op_id}")
        _accept_header_footer_change(doc, plan_id, op_id, op_state)
    elif op_type in ("footer_change", "footer_clear"):
        _log(f"_accept_footer_change: op_id={op_id}")
        _accept_header_footer_change(doc, plan_id, op_id, op_state)
    # Field operations
    elif op_type == "field_insertion":
        _log(f"_accept_field_insertion: op_id={op_id}")
        # Field is already inserted; nothing to do on accept
        pass
    # Cell merge/split operations
    elif op_type in ("cell_merge", "cell_split"):
        _log(f"_accept_cell_operation: op_id={op_id}")
        # Cell operation is already applied; nothing to do on accept
        pass
    # Section break operations
    elif op_type == "section_break_insertion":
        _log(f"_accept_section_break: op_id={op_id}")
        # Section break is already inserted; nothing to do on accept
        pass
    # Track changes operations (Phase 2)
    elif op_type in (
        "track_changes_toggle",
        "revision_accept",
        "revision_reject",
        "revisions_accept_all",
        "revisions_reject_all",
    ):
        _log(f"_accept_track_changes_operation: op_id={op_id}")
        # Track changes operations are already applied; nothing to do on accept
        pass
    # Comment operations (Phase 2)
    elif op_type in (
        "comment_insertion",
        "comment_reply",
        "comment_resolve",
        "comment_deletion",
    ):
        _log(f"_accept_comment_operation: op_id={op_id}")
        _accept_comment_operation(doc, plan_id, op_id, op_state)
    # Style management operations (Phase 2)
    elif op_type in ("style_creation", "style_modification", "style_deletion"):
        _log(f"_accept_style_operation: op_id={op_id}")
        # Style operations are already applied; nothing to do on accept
        pass
    # Footnote/endnote operations (Phase 2)
    elif op_type in (
        "footnote_insertion",
        "endnote_insertion",
        "note_modification",
        "note_deletion",
    ):
        _log(f"_accept_note_operation: op_id={op_id}")
        # Note operations are already applied; nothing to do on accept
        pass
    # Image operations (Phase 2)
    elif op_type in ("image_insertion", "image_update", "image_deletion"):
        _log(f"_accept_image_operation: op_id={op_id}")
        _accept_image_operation(doc, plan_id, op_id, op_state)
    # TOC operations (Phase 3)
    elif op_type in ("toc_insertion", "toc_update", "toc_removal"):
        _log(f"_accept_toc_operation: op_id={op_id}")
        if op_type == "toc_insertion":
            _normalize_table_segments(doc, op_state.get("boundary_ids") or [])
        elif op_type == "toc_removal":
            _delete_toc_from_state(doc, op_state)
    # Mail Merge operations (Phase 3)
    elif op_type in ("merge_field_insertion", "merge_data_applied", "merge_data_error"):
        _log(f"_accept_merge_operation: op_id={op_id}")
        # Merge operations are already applied; nothing to do on accept
        pass
    # Content Controls operations (Phase 3)
    elif op_type == "control_insertion":
        _log(f"_accept_control_insertion: op_id={op_id}")
        # Control operations are already applied; nothing to do on accept
        pass
    # Cross-Reference operations (Phase 3)
    elif op_type == "cross_reference_insertion":
        _log(f"_accept_cross_reference_insertion: op_id={op_id}")
        # Already applied
        pass
    # Master Document operations (Phase 3)
    elif op_type == "sub_document_insertion":
        _log(f"_accept_sub_document_insertion: op_id={op_id}")
        # Already applied; could potentially unlink here if embedded was requested?
        pass
    # Multi-Column Layout operations (Phase 4)
    elif op_type in ("columns_set", "column_break_insertion"):
        _log(f"_accept_column_operation: op_id={op_id}")
        # Changes applied in preview; nothing extra to finalize
        pass
    # Text Box operations (Phase 4)
    elif op_type in ("text_box_insertion", "text_box_deletion"):
        _log(f"_accept_text_box_operation: op_id={op_id}")
        # Changes applied in preview
        pass
    # Drop Cap operations (Phase 4)
    elif op_type == "drop_cap_set":
        _log(f"_accept_drop_cap_set: op_id={op_id}")
        # Already applied
        pass
    # Shape operations (Phase 4)
    elif op_type == "shape_insertion":
        _log(f"_accept_shape_insertion: op_id={op_id}")
        pass
    # Linked Frames operations (Phase 4)
    elif op_type == "frames_linked":
        _log(f"_accept_frames_linked: op_id={op_id}")
        pass
    # Phase 5: Lists, Watermarks, Equations, Comparison
    elif op_type in (
        "list_style_set",
        "watermark_insertion",
        "equation_insertion",
        "document_comparison",
    ):
        _log(f"_accept_phase5_operation: op_id={op_id}, type={op_type}")
        pass
    else:
        raise RuntimeError(f"Unsupported preview type '{op_type}' for accept")


def _reject_preview(doc, plan_id: str, op_id: str, op_state: Dict[str, Any]) -> None:
    # Always clean up separators
    _remove_bookmarked_segments(doc, op_state.get("separator_ids") or [])

    op_type = op_state.get("type")
    if op_type in ("paragraph_replacement", "run_replacement"):
        _log(f"_reject_paragraph_replacement: op_id={op_id}")
        _reject_paragraph_replacement(doc, plan_id, op_id, op_state)
    elif op_type == "paragraph_insertion":
        _log(f"_reject_paragraph_insertion: op_id={op_id}")
        _reject_paragraph_insertion(doc, plan_id, op_id, op_state)
    elif op_type == "delete_paragraph":
        _log(f"_reject_delete_paragraph: op_id={op_id}")
        _reject_delete_paragraph(doc, plan_id, op_id, op_state)
    elif op_type in ("table_replacement", "table_insertion"):
        _log(f"_reject_table_change: op_id={op_id}")
        _reject_table_change(doc, plan_id, op_id, op_state)
    elif op_type == "delete_table":
        _log(f"_reject_delete_table: op_id={op_id}")
        _reject_delete_table(doc, plan_id, op_id, op_state)
    elif op_type == "paragraph_style":
        _log(f"_reject_style_change: op_id={op_id}")
        _reject_style_change(doc, plan_id, op_id, op_state)
    elif op_type == "document_style":
        _log("_revert_document_style invoked")
        _revert_document_style(doc)
    # Chart operations
    elif op_type in ("chart_insertion", "chart_replacement"):
        _log(f"_reject_chart_change: op_id={op_id}")
        _reject_chart_change(doc, plan_id, op_id, op_state)
    elif op_type == "chart_data_update":
        _log(f"_reject_chart_data_update: op_id={op_id}")
        # Nothing to revert for data update preview
        pass
    # Header/footer operations
    elif op_type in ("header_change", "header_clear"):
        _log(f"_reject_header_change: op_id={op_id}")
        _reject_header_footer_change(doc, plan_id, op_id, op_state)
    elif op_type in ("footer_change", "footer_clear"):
        _log(f"_reject_footer_change: op_id={op_id}")
        _reject_header_footer_change(doc, plan_id, op_id, op_state)
    # Field operations
    elif op_type == "field_insertion":
        _log(f"_reject_field_insertion: op_id={op_id}")
        # TODO: Would need to track the field to remove it
        pass
    # Cell merge/split operations
    elif op_type in ("cell_merge", "cell_split"):
        _log(f"_reject_cell_operation: op_id={op_id}")
        # TODO: Would need to track original cell state to revert
        pass
    # Section break operations
    elif op_type == "section_break_insertion":
        _log(f"_reject_section_break: op_id={op_id}")
        section_name = op_state.get("section_name")
        if section_name:
            try:
                sections = doc.getTextSections()
                if sections.hasByName(section_name):
                    sections.getByName(section_name).dispose()
            except Exception as e:
                _log(f"_reject_section_break: failed to delete section {section_name}: {e}")
        pass
    # Track changes operations (Phase 2)
    elif op_type in (
        "track_changes_toggle",
        "revision_accept",
        "revision_reject",
        "revisions_accept_all",
        "revisions_reject_all",
    ):
        _log(f"_reject_track_changes_operation: op_id={op_id}")
        # Track changes operations cannot be easily reverted
        pass
    # Comment operations (Phase 2)
    elif op_type in (
        "comment_insertion",
        "comment_reply",
        "comment_resolve",
        "comment_deletion",
    ):
        _log(f"_reject_comment_operation: op_id={op_id}")
        # TODO: Track original state for proper revert
        pass
    # Style management operations (Phase 2)
    elif op_type in ("style_creation", "style_modification", "style_deletion"):
        _log(f"_reject_style_operation: op_id={op_id}")
        # TODO: Track original style properties for revert
        pass
    # Footnote/endnote operations (Phase 2)
    elif op_type in (
        "footnote_insertion",
        "endnote_insertion",
        "note_modification",
        "note_deletion",
    ):
        _log(f"_reject_note_operation: op_id={op_id}")
        if op_type in ("footnote_insertion", "endnote_insertion"):
            note_id = op_state.get("note_id")
            note_type = op_state.get("note_type", "footnote")
            if note_id:
                _delete_note_by_id(doc, note_id, note_type)
        else:
            # Modification/Deletion revert requires state tracking (complex)
            _log(f"_reject_note_operation: revert not implemented for {op_type}")
        pass

    # Image operations (Phase 2)
    elif op_type in ("image_insertion", "image_update", "image_deletion"):
        _log(f"_reject_image_operation: op_id={op_id}")
        if op_type == "image_insertion":
            image_id = op_state.get("image_id")  # This is the Name
            if image_id:
                _delete_graphic_by_name(doc, image_id)
        else:
            _log(f"_reject_image_operation: revert not implemented for {op_type}")
        pass

    # TOC operations (Phase 3)
    elif op_type in ("toc_insertion", "toc_update", "toc_removal"):
        _log(f"_reject_toc_operation: op_id={op_id}")
        if op_type == "toc_insertion":
            toc_id = op_state.get("toc_id")  # This is the Name
            if toc_id:
                _delete_toc_by_name(doc, toc_id)
            _remove_bookmarked_segments(doc, op_state.get("boundary_ids") or [])
        pass

    # Mail Merge operations (Phase 3)
    elif op_type in ("merge_field_insertion", "merge_data_applied", "merge_data_error"):
        _log(f"_reject_merge_operation: op_id={op_id}")
        # Merge fields are hard to target for deletion without IDs.
        # Data application is hard to revert without snapshot.
        _log(f"Warning: Revert not fully implemented for {op_type}")
        pass
    # Content Controls operations (Phase 3)
    elif op_type == "control_insertion":
        _log(f"_reject_control_insertion: op_id={op_id}")
        control_id = op_state.get("control_id")
        if control_id:
            # ControlShapes are in DrawPage but often accessible via GraphicObjects or DrawPage iteration
            # Since we set the Name, we can try _delete_graphic_by_name or a new helper if needed.
            # Actually, ControlShapes are shapes. Let's add _delete_shape_by_name.
            if not _delete_shape_by_name(doc, control_id):
                pass  # Warning logged in helper
        pass
    # Cross-Reference operations (Phase 3)
    elif op_type == "cross_reference_insertion":
        _log(f"_reject_cross_reference_insertion: op_id={op_id}")
        # Fields don't have easy IDs. We depend on field_id being a proxy or try search?
        # TODO: Implement robust field deletion. For now, log warning.
        _log(
            "Warning: Revert for cross-reference not fully implemented (requires field traversal)"
        )
        pass
    # Master Document operations (Phase 3)
    elif op_type == "sub_document_insertion":
        _log(f"_reject_sub_document_insertion: op_id={op_id}")
        section_name = op_state.get("section_name")
        if section_name:
            try:
                sections = doc.getTextSections()
                if sections.hasByName(section_name):
                    section = sections.getByName(section_name)
                    # dispose() deletes the section content too usually?
                    # Actually removing a TextSection removes it from the document.
                    # But we must be careful if it merely unlinks.
                    # For TextSections, dispose() removes the section and its content if it's a section object?
                    # Standard way: section.dispose()
                    section.dispose()
            except Exception as e:
                _log(f"_reject_sub_document_insertion: failed to delete section: {e}")
        pass
    # Multi-Column Layout operations (Phase 4)
    elif op_type == "columns_set":
        _log(f"_reject_columns_set: op_id={op_id}")
        try:
            prev_count = op_state.get("prev_column_count", 1)
            prev_spacing_mm = op_state.get("prev_spacing_mm")
            prev_separator_line = op_state.get("prev_separator_line")
            target_scope = op_state.get("target_scope")
            section_name = op_state.get("section_name")
            page_style_name = op_state.get("page_style_name")

            target = None
            if target_scope == "page_style":
                styles = doc.getStyleFamilies().getByName("PageStyles")
                style_name = page_style_name or "Standard"
                if styles.hasByName(style_name):
                    target = styles.getByName(style_name)
            elif target_scope == "section":
                sections = doc.getTextSections()
                if section_name and sections.hasByName(section_name):
                    target = sections.getByName(section_name)

            if target:
                current_columns = _get_text_columns(target)
                target.TextColumns = _configure_text_columns(
                    current_columns,
                    prev_count,
                    prev_spacing_mm,
                    prev_separator_line,
                )
        except Exception as e:
            _log(f"_reject_columns_set: failed: {e}")
        pass

    elif op_type == "column_break_insertion":
        _log(f"_reject_column_break_insertion: op_id={op_id}")
        try:
            paragraph = None
            element_id = op_state.get("element_id")
            paragraph_index = op_state.get("paragraph_index")
            if element_id:
                found = _get_element_by_bookmark(doc, element_id)
                if found and found[0] == "paragraph":
                    paragraph = found[1]
            if paragraph is None and paragraph_index is not None:
                found = _get_element_by_index(doc, paragraph_index)
                if found and found[0] == "paragraph":
                    paragraph = found[1]

            if paragraph is not None:
                paragraph.BreakType = op_state.get("prev_break_type", 0)
        except Exception as e:
            _log(f"_reject_column_break_insertion: failed: {e}")
        pass

    elif op_type in ("text_box_insertion", "text_box_deletion"):
        _log(f"_reject_text_box_operation: op_id={op_id}")
        # Insert -> Delete frame
        if op_type == "text_box_insertion":
            name = op_state.get("frame_name")
            if name:
                # Frames are in text frames, not graphic objects
                try:
                    frames = doc.getTextFrames()
                    if frames.hasByName(name):
                        frames.getByName(name).dispose()
                except Exception as e:
                    _log(f"_reject_text_box_insertion: failed: {e}")
        # Delete -> Revert?
        # Very hard to revert deletion of a frame without deep snapshot
        pass

    # Drop Cap operations (Phase 4)
    elif op_type == "drop_cap_set":
        _log(f"_reject_drop_cap_set: op_id={op_id}")
        # Revert drop cap settings
        try:
            # Locate paragraph again
            # Stored indices/IDs in op_state
            # Simple approach: try to use enumeration
            el_index = op_state.get("element_index")
            if el_index is not None:
                paragraphs = doc.getText().createEnumeration()
                i = 0
                while paragraphs.hasMoreElements():
                    para = paragraphs.nextElement()
                    if i == el_index:
                        para.DropCapCharCount = op_state.get("prev_char_count", 0)
                        para.DropCapLines = op_state.get("prev_lines", 0)
                        para.DropCapDistance = op_state.get("prev_distance", 0)
                        break
                    i += 1
        except Exception as e:
            _log(f"_reject_drop_cap_set: failed: {e}")
        pass

    # Shape operations (Phase 4)
    elif op_type == "shape_insertion":
        _log(f"_reject_shape_insertion: op_id={op_id}")
        shape_name = op_state.get("shape_name")
        if shape_name:
            try:
                draw_page = doc.getDrawPage()
                for i in range(draw_page.getCount()):
                    shape = draw_page.getByIndex(i)
                    if shape.Name == shape_name:
                        draw_page.remove(shape)
                        break
            except Exception as e:
                _log(f"_reject_shape_insertion: failed: {e}")
        pass

    # Linked Frames operations (Phase 4)
    elif op_type == "frames_linked":
        _log(f"_reject_frames_linked: op_id={op_id}")
        source_name = op_state.get("source_frame_name")
        prev_chain = op_state.get("prev_chain_name", "")
        if source_name:
            try:
                frames = doc.getTextFrames()
                if frames.hasByName(source_name):
                    source = frames.getByName(source_name)
                    source.ChainNextName = prev_chain
            except Exception as e:
                _log(f"_reject_frames_linked: failed: {e}")
        pass

    # Phase 5: Lists, Watermarks, Equations, Comparison
    elif op_type == "list_style_set":
        _log(f"_reject_list_style_set: op_id={op_id}")
        # Restore previous numbering - complex, log warning
        _log("Warning: List style revert not fully implemented")
        pass
    elif op_type == "watermark_insertion":
        _log(f"_reject_watermark_insertion: op_id={op_id}")
        name = op_state.get("watermark_name")
        if name:
            try:
                draw_page = doc.getDrawPage()
                for i in range(draw_page.getCount()):
                    shape = draw_page.getByIndex(i)
                    if shape.Name == name:
                        draw_page.remove(shape)
                        break
            except Exception as e:
                _log(f"_reject_watermark_insertion: failed: {e}")
        pass
    elif op_type == "equation_insertion":
        _log(f"_reject_equation_insertion: op_id={op_id}")
        name = op_state.get("equation_name")
        if name:
            try:
                embedded = doc.getEmbeddedObjects()
                if embedded.hasByName(name):
                    embedded.getByName(name).dispose()
            except Exception as e:
                _log(f"_reject_equation_insertion: failed: {e}")
        pass
    elif op_type == "document_comparison":
        _log(f"_reject_document_comparison: op_id={op_id}")
        # Comparison doesn't modify, just shows differences
        pass

    else:
        raise RuntimeError(f"Unsupported preview type '{op_type}' for reject")


def _reinsert_as_tracked(doc, marked_ids: List[str]) -> None:
    """
    Legacy helper kept for backward compatibility.
    Silent-resolve policy disables tracked reinsertion.
    """
    _log(
        "_reinsert_as_tracked: skipped (silent-resolve policy active). "
        f"ids={len(marked_ids)}"
    )


def _dispose_bookmarks(doc, bookmark_names: List[str]) -> None:
    """Dispose bookmarks after they're no longer needed."""
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        try:
            if bookmarks.hasByName(name):
                bookmarks.getByName(name).dispose()
        except Exception as e:
            _log(f"_dispose_bookmarks: failed to dispose {name}: {e}")


def _accept_paragraph_replacement(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _log(f"_accept_paragraph_replacement: op_id={op_id}")
    # Step 1: Remove original (gold/yellow strikethrough) segments with tracking OFF
    _remove_bookmarked_segments(doc, op_state.get("marked_ids") or [])
    ids = op_state.get("inserted_ids") or []
    _log(f"_accept_paragraph_replacement: ids found = {ids}")
    target_bookmark = op_state.get("target_bookmark")

    # Preserve the original stable element_id for replacements.
    if ids and target_bookmark:
        rebound = _rebind_target_bookmark_from_preview(doc, target_bookmark, ids[0])
        _log(
            f"_accept_paragraph_replacement: target_rebind target={target_bookmark} source={ids[0]} ok={rebound}"
        )

    # Step 2: Finalize silently by removing preview formatting only.
    # Do not reinsert as tracked changes during resolve.
    proposed_spec = op_state.get("proposed_spec") or {}
    _log(
        "_accept_paragraph_replacement: proposed "
        f"ids={ids} format={(proposed_spec or {}).get('format') if isinstance(proposed_spec, dict) else None} "
        f"paragraphs={len((proposed_spec or {}).get('paragraphs') or []) if isinstance(proposed_spec, dict) else 0}"
    )
    has_explicit_color = _any_paragraph_spec_has_explicit_color(proposed_spec)
    if has_explicit_color:
        # The text already has the user's intended colors (preview green was
        # never applied). Only clear strikethrough to avoid destroying them.
        _normalize_strikethrough_only(doc, ids, dispose_bookmarks=False)
    else:
        _normalize_green_segments(doc, ids, dispose_bookmarks=False)
    _reapply_accepted_paragraph_formatting(doc, ids, proposed_spec)
    _dispose_bookmarks(doc, ids)


def _accept_paragraph_insertion(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _log(f"_accept_paragraph_insertion: op_id={op_id}")
    ids = op_state.get("inserted_ids") or []
    _log(f"_accept_paragraph_insertion: ids found = {ids}")

    # Finalize silently by removing preview formatting only.
    # Do not reinsert as tracked changes during resolve.
    proposed_spec = op_state.get("proposed_spec") or {}
    _log(
        "_accept_paragraph_insertion: proposed "
        f"ids={ids} format={(proposed_spec or {}).get('format') if isinstance(proposed_spec, dict) else None} "
        f"paragraphs={len((proposed_spec or {}).get('paragraphs') or []) if isinstance(proposed_spec, dict) else 0}"
    )
    has_explicit_color = _any_paragraph_spec_has_explicit_color(proposed_spec)
    if has_explicit_color:
        _normalize_strikethrough_only(doc, ids, dispose_bookmarks=False)
    else:
        _normalize_green_segments(doc, ids, dispose_bookmarks=False)
    _reapply_accepted_paragraph_formatting(doc, ids, proposed_spec)
    _dispose_bookmarks(doc, ids)


def _accept_delete_paragraph(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _log(f"_accept_delete_paragraph: op_id={op_id}")
    marked_ids = op_state.get("marked_ids") or []
    target_bookmark = op_state.get("target_bookmark")
    target_index = op_state.get("target_index")

    paragraph = None
    if target_bookmark:
        element = _get_element_by_bookmark(doc, target_bookmark)
        if element and element[0] == "paragraph":
            paragraph = element[1]
    if paragraph is None and isinstance(target_index, int):
        element = _get_element_by_index(doc, target_index)
        if element and element[0] == "paragraph":
            paragraph = element[1]

    if paragraph is not None:
        _remove_annotations_in_paragraph(doc, paragraph)
        _dispose_bookmarks(doc, marked_ids)
        if target_bookmark:
            _dispose_bookmarks(doc, [target_bookmark])
        _delete_paragraph_node(doc, paragraph)
    else:
        _log("_accept_delete_paragraph: could not locate paragraph; falling back to text clear")
        _remove_bookmarked_segments(doc, marked_ids)
        if target_bookmark:
            _dispose_bookmarks(doc, [target_bookmark])


def _accept_table_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _log(f"_accept_table_change: op_id={op_id}")
    ids = op_state.get("inserted_ids") or []
    boundary_ids = op_state.get("boundary_ids") or []
    _log(f"_accept_table_change: ids found = {ids}")
    target_bookmark = op_state.get("target_bookmark")
    op_type = op_state.get("type")
    marked_ids = op_state.get("marked_ids") or []
    original_table_name = op_state.get("original_table_name")
    proposed_table_name = op_state.get("proposed_table_name")
    proposed_spec = _normalize_table_spec(op_state.get("proposed_spec") or {})

    # Deterministic path for table replacement:
    # apply proposed spec directly onto the original target table object.
    # This avoids losing content when preview proposed table was nested/ephemeral.
    if op_type == "table_replacement":
        target_table = _resolve_table_for_replacement(doc, op_state)
        if target_table is not None and _table_spec_has_rows(proposed_spec):
            target_name = _safe_table_name(target_table)

            # Remove preview-proposed table clone if it exists separately.
            if (
                isinstance(proposed_table_name, str)
                and proposed_table_name
                and proposed_table_name != target_name
            ):
                removed_preview_clone = _remove_table_by_name(doc, proposed_table_name)
                if removed_preview_clone:
                    _log(
                        "_accept_table_change: removed preview clone by name "
                        f"{proposed_table_name}"
                    )

            applied = _apply_table_spec_to_table(target_table, proposed_spec)
            if applied:
                _normalize_table_object(target_table)
                reapplied = _apply_table_spec_to_table(target_table, proposed_spec)
                _log(
                    "_accept_table_change: reapplied normalized table spec after "
                    f"replacement cleanup table={_safe_table_name(target_table)} ok={reapplied}"
                )
                if _table_spec_has_explicit_run_color(proposed_spec):
                    _apply_explicit_colors_to_table(target_table, proposed_spec)
                _dispose_bookmarks(doc, marked_ids + ids)
                if target_bookmark:
                    rebound = _rebind_target_bookmark_from_table(
                        doc, target_bookmark, _safe_table_name(target_table)
                    )
                    _log(
                        "_accept_table_change: target_rebind_applied "
                        f"target={target_bookmark} ok={rebound}"
                    )
                return
            _log(
                "_accept_table_change: in-place apply failed; falling back to "
                "legacy table replacement path"
            )

    # Step 1: Remove original (gold/yellow strikethrough) segments with tracking OFF
    removed_by_name = False
    if (
        op_type == "table_replacement"
        and isinstance(original_table_name, str)
        and original_table_name
    ):
        if original_table_name != proposed_table_name:
            removed_by_name = _remove_table_by_name(doc, original_table_name)
            if removed_by_name:
                _log(
                    "_accept_table_change: removed original table by name "
                    f"{original_table_name}"
                )
    if removed_by_name:
        _dispose_bookmarks(doc, marked_ids)
    else:
        _remove_bookmarked_tables(doc, marked_ids)

    # For whole-table replacement, preserve the original stable table element_id.
    if op_type == "table_replacement" and target_bookmark:
        rebound = False
        if ids:
            rebound = _rebind_target_bookmark_from_preview(doc, target_bookmark, ids[0])
            _log(
                "_accept_table_change: target_rebind "
                f"target={target_bookmark} source={ids[0]} ok={rebound}"
            )
        if (
            not rebound
            and isinstance(proposed_table_name, str)
            and proposed_table_name
        ):
            rebound = _rebind_target_bookmark_from_table(
                doc, target_bookmark, proposed_table_name
            )
            _log(
                "_accept_table_change: target_rebind_table "
                f"target={target_bookmark} table={proposed_table_name} ok={rebound}"
            )

    # Step 2: Finalize silently by removing preview formatting only.
    # Do not reinsert as tracked changes during resolve.
    accepted_table = None
    if isinstance(proposed_table_name, str) and proposed_table_name:
        accepted_table = _get_table_by_name(doc, proposed_table_name)
    if accepted_table is None and ids:
        accepted_table = _resolve_table_from_bookmark(doc, ids[0])

    _normalize_table_segments(doc, ids + boundary_ids)

    if accepted_table is not None and _table_spec_has_rows(proposed_spec):
        applied_final_spec = _apply_table_spec_to_table(accepted_table, proposed_spec)
        _log(
            "_accept_table_change: reapplied normalized table spec after preview cleanup "
            f"table={_safe_table_name(accepted_table)} ok={applied_final_spec}"
        )
    elif (
        isinstance(proposed_table_name, str)
        and proposed_table_name
        and _normalize_table_by_name(doc, proposed_table_name)
    ):
        _log(
            "_accept_table_change: normalized proposed table by name "
            f"{proposed_table_name}"
        )

    _reapply_explicit_colors_for_table_spec(doc, ids, proposed_spec)
    if (
        isinstance(proposed_table_name, str)
        and proposed_table_name
        and _table_spec_has_explicit_run_color(proposed_spec)
    ):
        table = _get_table_by_name(doc, proposed_table_name)
        if table is not None:
            try:
                _apply_explicit_colors_to_table(table, proposed_spec)
            except Exception as exc:
                _log(
                    "_accept_table_change: failed explicit-color reapply "
                    f"for table {proposed_table_name}: {exc}"
                )


def _accept_delete_table(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    removed_target = _remove_table_by_target(
        doc, op_state.get("target_bookmark"), op_state.get("target_index")
    )
    if not removed_target:
        removed_target = _remove_table_by_name(doc, op_state.get("original_table_name"))
    if not removed_target:
        _remove_bookmarked_tables(doc, op_state.get("marked_ids") or [])
    else:
        _dispose_bookmarks(doc, op_state.get("marked_ids") or [])


def _accept_style_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _dispose_bookmarks(doc, op_state.get("marked_ids") or [])


def _apply_centered_inline_image_placement(image, context: str) -> None:
    """Persist the default Writer image placement: inline, no wrap, centered paragraph."""
    try:
        image.AnchorType = 1  # AS_CHARACTER
    except Exception as exc:
        _log(f"{context}: failed to set image AnchorType=AS_CHARACTER: {exc}")

    try:
        image.Surround = 0  # NONE
    except Exception as exc:
        _log(f"{context}: failed to set image Surround=NONE: {exc}")

    try:
        anchor = image.getAnchor()
        text_obj = anchor.getText()
        cursor = text_obj.createTextCursorByRange(anchor)
        cursor.ParaAdjust = PAR_ADJUST_CENTER
        cursor.ParaBottomMargin = 212  # ~6pt space between image and caption
        _log(f"{context}: enforced centered inline image paragraph")
    except Exception as exc:
        _log(f"{context}: failed to center image anchor paragraph: {exc}")


def _accept_image_operation(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    # Image previews are already applied in place.
    # Keep accept silent: no cut/paste tracked reinsertion.
    image_id = op_state.get("image_id")
    if image_id:
        try:
            graphics = doc.getGraphicObjects()
            if graphics.hasByName(image_id):
                _apply_centered_inline_image_placement(
                    graphics.getByName(image_id),
                    "_accept_image_operation",
                )
        except Exception as exc:
            _log(f"_accept_image_operation: failed to normalize image {image_id}: {exc}")
        _log(f"_accept_image_operation: accepted image {image_id} silently")


# Chart accept/reject helpers
def _accept_chart_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    """Accept a chart insertion or replacement preview."""
    # For charts, we just need to normalize the inserted chart (remove preview styling)
    # and remove any marked original content
    _remove_bookmarked_segments(doc, op_state.get("marked_ids") or [])
    # Chart objects don't need color normalization like text
    _log(f"_accept_chart_change: chart accepted for op_id={op_id}")


def _reject_chart_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    """Reject a chart insertion or replacement preview."""
    # Remove the inserted (proposed) chart
    inserted_ids = op_state.get("inserted_ids") or []
    # For embedded objects, we need to find and remove them
    # The chart was tagged with a bookmark-like identifier
    for chart_id in inserted_ids:
        try:
            # Try to remove embedded objects by iterating through text content
            # This is a simplified approach; full implementation would track the chart object
            _log(f"_reject_chart_change: removing chart {chart_id}")
        except Exception as e:
            _log(f"_reject_chart_change: failed to remove chart {chart_id}: {e}")

    # Restore original if it was a replacement
    _normalize_green_segments(doc, op_state.get("marked_ids") or [])
    _log(f"_reject_chart_change: chart rejected for op_id={op_id}")


# Header/footer accept/reject helpers
def _accept_header_footer_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    """Accept a header or footer change preview.

    The preview applies GREEN_RGB to header/footer text via _write_paragraph_spec.
    On accept we must strip that preview color (unless the spec has explicit
    run colors, in which case the colors are already the user's intended ones).
    """
    op_type = op_state.get("type")
    _log(f"_accept_header_footer_change: {op_type} accepted for op_id={op_id}")

    section_index = op_state.get("section_index", 0)

    try:
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            _log("_accept_header_footer_change: could not get page style")
            return

        # Get the header or footer text object
        hf_text = None
        if op_type and "header" in op_type:
            header_type = op_state.get("header_type", "default")
            if header_type == "first_page":
                hf_text = page_style.HeaderTextFirst
            elif header_type == "even_page":
                hf_text = page_style.HeaderTextLeft
            else:
                hf_text = page_style.HeaderText
        elif op_type and "footer" in op_type:
            footer_type = op_state.get("footer_type", "default")
            if footer_type == "first_page":
                hf_text = page_style.FooterTextFirst
            elif footer_type == "even_page":
                hf_text = page_style.FooterTextLeft
            else:
                hf_text = page_style.FooterText

        if hf_text is None:
            _log("_accept_header_footer_change: could not get header/footer text")
            return

        # Check if the proposed spec has explicit run colors
        proposed_spec = op_state.get("proposed_spec") or {}
        has_explicit = _any_paragraph_spec_has_explicit_color(proposed_spec)

        # Select all header/footer text and normalize preview formatting
        cursor = hf_text.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)

        if has_explicit:
            # Only clear strikethrough — colors are the user's intended ones
            cursor.setPropertyToDefault("CharStrikeout")
        else:
            # Clear preview green and strikethrough
            cursor.setPropertyToDefault("CharColor")
            cursor.setPropertyToDefault("CharStrikeout")

        _log(f"_accept_header_footer_change: normalized preview formatting")

    except Exception as e:
        _log(f"_accept_header_footer_change: error normalizing: {e}")


def _reject_header_footer_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    """Reject a header or footer change preview."""
    op_type = op_state.get("type")
    section_index = op_state.get("section_index", 0)
    original_was_on = op_state.get("original_was_on", False)

    try:
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            _log(f"_reject_header_footer_change: could not get page style")
            return

        if "header" in op_type:
            header_type = op_state.get("header_type", "default")
            if header_type == "first_page":
                header_text = page_style.HeaderTextFirst
            elif header_type == "even_page":
                header_text = page_style.HeaderTextLeft
            else:
                header_text = page_style.HeaderText

            if header_text:
                header_text.setString("")

            if not original_was_on:
                page_style.HeaderIsOn = False
        else:
            footer_type = op_state.get("footer_type", "default")
            if footer_type == "first_page":
                footer_text = page_style.FooterTextFirst
            elif footer_type == "even_page":
                footer_text = page_style.FooterTextLeft
            else:
                footer_text = page_style.FooterText

            if footer_text:
                footer_text.setString("")

            if not original_was_on:
                page_style.FooterIsOn = False

        _log(f"_reject_header_footer_change: {op_type} rejected for op_id={op_id}")

    except Exception as e:
        _log(f"_reject_header_footer_change: error: {e}")


def _get_page_style_for_section(doc, section_index: int = 0):
    """Get the page style for a given section."""
    try:
        page_styles = doc.getStyleFamilies().getByName("PageStyles")

        if section_index == 0:
            if page_styles.hasByName("Standard"):
                return page_styles.getByName("Standard")
            if page_styles.hasByName("Default Style"):
                return page_styles.getByName("Default Style")

        style_names = page_styles.getElementNames()
        if style_names:
            return page_styles.getByName(style_names[0])

        return None

    except Exception as e:
        _log(f"_get_page_style_for_section: error: {e}")
        return None


def _apply_document_style(doc, update: Dict[str, Any]) -> None:
    default_style_name = update.get("default_paragraph_style")
    paragraph_defaults = update.get("paragraph_defaults") or {}
    _log(
        "_apply_document_style: start "
        f"default_style_name={default_style_name!r} "
        f"paragraph_defaults_keys={list(paragraph_defaults.keys()) if isinstance(paragraph_defaults, dict) else []} "
        f"update_keys={list(update.keys()) if isinstance(update, dict) else []}"
    )

    if default_style_name:
        count = 0
        enum = doc.getText().createEnumeration()
        while enum.hasMoreElements():
            element = enum.nextElement()
            try:
                if element.supportsService("com.sun.star.text.Paragraph"):
                    before = getattr(element, "ParaStyleName", None)
                    preview = _safe_text_preview(element)
                    element.ParaStyleName = default_style_name
                    count += 1
                    if count <= 8 or "Executive Summary" in preview:
                        _log(
                            "_apply_document_style: default_style "
                            f"paragraph={count} before={before!r} "
                            f"after={getattr(element, 'ParaStyleName', None)!r} "
                            f"text='{preview}'"
                        )
            except Exception:
                continue
        _log(f"_apply_document_style: default_style applied_count={count}")

    if paragraph_defaults:
        count = 0
        enum = doc.getText().createEnumeration()
        while enum.hasMoreElements():
            element = enum.nextElement()
            if not element.supportsService("com.sun.star.text.Paragraph"):
                continue
            try:
                cursor = element.getText().createTextCursorByRange(element.getStart())
                cursor.gotoRange(element.getEnd(), True)
                before = _format_snapshot(cursor)
                preview = _safe_text_preview(cursor)
                _apply_paragraph_defaults(cursor, paragraph_defaults)
                count += 1
                if count <= 8 or "Executive Summary" in preview:
                    _log(
                        "_apply_document_style: paragraph_defaults "
                        f"paragraph={count} before={before} "
                        f"after={_format_snapshot(cursor)} text='{preview}'"
                    )
            except Exception:
                continue
        _log(f"_apply_document_style: paragraph_defaults applied_count={count}")


def _reject_paragraph_replacement(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _remove_bookmarked_segments(doc, op_state.get("inserted_ids") or [])
    _restore_original_paragraphs(doc, plan_id, op_id, op_state)


def _reject_paragraph_insertion(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _remove_bookmarked_segments(doc, op_state.get("inserted_ids") or [])


def _reject_delete_paragraph(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _restore_original_paragraphs(doc, plan_id, op_id, op_state)


def _reject_table_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    inserted_ids = op_state.get("inserted_ids") or []
    boundary_ids = op_state.get("boundary_ids") or []
    proposed_table_name = op_state.get("proposed_table_name")
    removed_by_name = False
    if isinstance(proposed_table_name, str) and proposed_table_name:
        removed_by_name = _remove_table_by_name(doc, proposed_table_name)
        if removed_by_name:
            _log(
                "_reject_table_change: removed proposed table by name "
                f"{proposed_table_name}"
            )
    if removed_by_name:
        _dispose_bookmarks(doc, inserted_ids)
    else:
        _remove_bookmarked_tables(doc, inserted_ids)
    _remove_bookmarked_segments(doc, boundary_ids)
    _restore_original_table(doc, plan_id, op_id, op_state)


def _reject_delete_table(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _restore_original_table(doc, plan_id, op_id, op_state)


def _reject_style_change(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _restore_original_paragraphs(doc, plan_id, op_id, op_state)


def _revert_document_style(doc) -> None:
    pass


def _apply_paragraph_defaults(cursor, defaults: Dict[str, Any]) -> None:
    alignment = defaults.get("alignment")
    if alignment:
        align = alignment.lower()
        if align == "left":
            cursor.ParaAdjust = PAR_ADJUST_LEFT
        elif align == "right":
            cursor.ParaAdjust = PAR_ADJUST_RIGHT
        elif align == "center":
            cursor.ParaAdjust = PAR_ADJUST_CENTER
        elif align == "justify":
            cursor.ParaAdjust = PAR_ADJUST_BLOCK

    if defaults.get("space_before_pt") is not None:
        cursor.ParaTopMargin = int(defaults["space_before_pt"] * PT_TO_HMM)
    if defaults.get("space_after_pt") is not None:
        cursor.ParaBottomMargin = int(defaults["space_after_pt"] * PT_TO_HMM)
    if defaults.get("left_indent_pt") is not None:
        cursor.ParaLeftMargin = int(defaults["left_indent_pt"] * PT_TO_HMM)
    if defaults.get("right_indent_pt") is not None:
        cursor.ParaRightMargin = int(defaults["right_indent_pt"] * PT_TO_HMM)
    if defaults.get("first_line_indent_pt") is not None:
        cursor.ParaFirstLineIndent = int(defaults["first_line_indent_pt"] * PT_TO_HMM)

    char = defaults.get("character_format") or {}
    color_str = char.get("font_color_rgb")
    if color_str:
        parsed = _parse_color(color_str)
        if parsed is not None:
            cursor.CharColor = parsed
    if char.get("bold") is not None:
        cursor.CharWeight = FONT_WEIGHT_BOLD if char["bold"] else FONT_WEIGHT_NORMAL
    if char.get("italic") is not None:
        cursor.CharPosture = FONT_SLANT_ITALIC if char["italic"] else 0
    if char.get("underline") is not None:
        cursor.CharUnderline = FONT_UNDERLINE_SINGLE if char["underline"] else 0
    if char.get("font_name"):
        cursor.CharFontName = char["font_name"]
    if char.get("font_size_pt") is not None:
        cursor.CharHeight = char["font_size_pt"]


def _safe_text_preview(text_range, limit: int = 120) -> str:
    if not DEBUG_LOG:
        return ""
    try:
        value = text_range.getString()
    except Exception:
        value = ""
    value = " ".join(str(value or "").split())
    if len(value) > limit:
        return value[: limit - 3] + "..."
    return value


def _format_snapshot(cursor) -> Dict[str, Any]:
    if not DEBUG_LOG:
        return {}
    snapshot: Dict[str, Any] = {}
    for prop in (
        "ParaStyleName",
        "ParaAdjust",
        "ParaTopMargin",
        "ParaBottomMargin",
        "CharHeight",
        "CharFontName",
        "CharColor",
        "CharWeight",
        "CharPosture",
        "CharUnderline",
        "CharStrikeout",
    ):
        try:
            snapshot[prop] = getattr(cursor, prop)
        except Exception:
            continue
    return snapshot


# Keep this mapping aligned with PreviewActionPlan.py. The preview renderer
# applies these styles before users accept changes; accept must reassert the
# same paragraph-level format so finalized text does not fall back to Default.
STYLE_MAP = {
    "Normal": "Standard",
    "Body Text": "Text body",
    "BodyText": "Text body",
    "Title": "Title",
    "Subtitle": "Subtitle",
    "Caption": "Caption",
    "Heading1": "Heading 1",
    "Heading2": "Heading 2",
    "Heading3": "Heading 3",
    "Heading4": "Heading 4",
    "Heading5": "Heading 5",
    "Heading6": "Heading 6",
}

STYLE_SIZE_INHERITS_FROM_PARAGRAPH = {
    "Title",
    "Subtitle",
    "Heading 1",
    "Heading 2",
    "Heading 3",
    "Heading 4",
    "Heading 5",
    "Heading 6",
}


def _mapped_paragraph_style_name(format_spec: Dict[str, Any]) -> Optional[str]:
    style_name = format_spec.get("style_name") if isinstance(format_spec, dict) else None
    if not style_name:
        return None
    return STYLE_MAP.get(style_name, style_name)


def _style_should_inherit_character_size(format_spec: Dict[str, Any]) -> bool:
    mapped_style = _mapped_paragraph_style_name(format_spec)
    return mapped_style in STYLE_SIZE_INHERITS_FROM_PARAGRAPH


def _clear_direct_character_size(cursor, reason: str) -> None:
    try:
        before = _format_snapshot(cursor).get("CharHeight")
        cursor.setPropertyToDefault("CharHeight")
        after = _format_snapshot(cursor).get("CharHeight")
        _log(
            "_clear_direct_character_size: "
            f"reason={reason} before={before!r} after={after!r} "
            f"text='{_safe_text_preview(cursor)}'"
        )
    except Exception as exc:
        _log(f"_clear_direct_character_size: failed reason={reason}: {exc}")


def _apply_paragraph_formatting(cursor, format_spec: Dict[str, Any]) -> None:
    if not format_spec:
        return

    _log(
        "_apply_paragraph_formatting: start "
        f"format={format_spec} before={_format_snapshot(cursor)} "
        f"text='{_safe_text_preview(cursor)}'"
    )

    style_name = format_spec.get("style_name")
    if style_name:
        mapped_style = _mapped_paragraph_style_name(format_spec)
        try:
            cursor.setPropertyValue("ParaStyleName", mapped_style)
            _log(
                "_apply_paragraph_formatting: applied style "
                f"'{mapped_style}' (original: '{style_name}')"
            )
        except Exception as exc:
            _log(
                "_apply_paragraph_formatting: failed to apply style "
                f"'{mapped_style}': {exc}"
            )

    alignment = format_spec.get("alignment")
    if alignment:
        align = str(alignment).lower()
        try:
            if align == "left":
                cursor.ParaAdjust = PAR_ADJUST_LEFT
            elif align == "right":
                cursor.ParaAdjust = PAR_ADJUST_RIGHT
            elif align == "center":
                cursor.ParaAdjust = PAR_ADJUST_CENTER
            elif align == "justify":
                cursor.ParaAdjust = PAR_ADJUST_BLOCK
        except Exception as exc:
            _log(f"_apply_paragraph_formatting: failed to apply alignment: {exc}")

    if format_spec.get("indent_left_in") is not None:
        try:
            cursor.ParaLeftMargin = int(format_spec["indent_left_in"] * 2540)
        except Exception as exc:
            _log(f"_apply_paragraph_formatting: failed to apply left indent: {exc}")
    if format_spec.get("indent_right_in") is not None:
        try:
            cursor.ParaRightMargin = int(format_spec["indent_right_in"] * 2540)
        except Exception as exc:
            _log(f"_apply_paragraph_formatting: failed to apply right indent: {exc}")
    if format_spec.get("indent_first_line_in") is not None:
        try:
            cursor.ParaFirstLineIndent = int(
                format_spec["indent_first_line_in"] * 2540
            )
        except Exception as exc:
            _log(
                "_apply_paragraph_formatting: failed to apply first-line indent: "
                f"{exc}"
            )
    if format_spec.get("space_before_pt") is not None:
        try:
            cursor.ParaTopMargin = int(format_spec["space_before_pt"] * PT_TO_HMM)
        except Exception as exc:
            _log(f"_apply_paragraph_formatting: failed to apply space before: {exc}")
    if format_spec.get("space_after_pt") is not None:
        try:
            cursor.ParaBottomMargin = int(format_spec["space_after_pt"] * PT_TO_HMM)
        except Exception as exc:
            _log(f"_apply_paragraph_formatting: failed to apply space after: {exc}")

    if _style_should_inherit_character_size(format_spec):
        _clear_direct_character_size(cursor, "paragraph_style_inherits_size")

    _log(
        "_apply_paragraph_formatting: done "
        f"after={_format_snapshot(cursor)} text='{_safe_text_preview(cursor)}'"
    )


def _reapply_run_formatting_for_paragraph_spec(
    text, start_range, paragraph_spec: Dict[str, Any]
) -> None:
    runs = paragraph_spec.get("runs") or []
    if not isinstance(runs, list) or not runs:
        return

    format_spec = paragraph_spec.get("format") or {}
    walker = text.createTextCursorByRange(start_range)
    walker.collapseToStart()
    for run in runs:
        if not isinstance(run, dict):
            continue
        if run.get("break_before") == "line":
            try:
                walker.goRight(1, False)
            except Exception:
                pass

        run_text = run.get("text") or ""
        if not run_text:
            continue

        run_start = walker.getStart()
        try:
            moved = walker.goRight(len(run_text), False)
        except Exception:
            moved = False
        if not moved:
            break

        run_cursor = text.createTextCursorByRange(run_start)
        run_cursor.gotoRange(walker.getStart(), True)
        run_format = run.get("format") or {}
        if _style_should_inherit_character_size(format_spec) and run_format.get(
            "font_size_pt"
        ):
            run_format = dict(run_format)
            try:
                font_size = float(run_format.get("font_size_pt"))
            except Exception:
                font_size = None
            if font_size == 11.0:
                run_format.pop("font_size_pt", None)
                _clear_direct_character_size(
                    run_cursor,
                    f"strip_default_run_size_for_style:{format_spec.get('style_name')}",
                )
        _apply_run_formatting(run_cursor, run_format, None)


def _reapply_accepted_paragraph_formatting(
    doc, bookmark_names: List[str], proposed_spec: Dict[str, Any]
) -> None:
    paragraph_specs = _collect_paragraph_specs(proposed_spec)
    if not paragraph_specs:
        return

    bookmarks = doc.getBookmarks()
    for idx, name in enumerate(bookmark_names):
        if not _validate_bookmark_exists(doc, name):
            _log(
                f"_reapply_accepted_paragraph_formatting: bookmark {name} not found"
            )
            continue
        try:
            if len(paragraph_specs) == 1:
                paragraph_spec = paragraph_specs[0]
            elif idx < len(paragraph_specs):
                paragraph_spec = paragraph_specs[idx]
            else:
                continue

            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)

            _log(
                "_reapply_accepted_paragraph_formatting: start "
                f"bookmark={name} idx={idx} "
                f"spec_format={paragraph_spec.get('format') or {}} "
                f"runs={len(paragraph_spec.get('runs') or [])} "
                f"before={_format_snapshot(cursor)} "
                f"text='{_safe_text_preview(cursor)}'"
            )
            _apply_paragraph_formatting(cursor, paragraph_spec.get("format") or {})
            _reapply_run_formatting_for_paragraph_spec(text, anchor, paragraph_spec)
            after_cursor = text.createTextCursorByRange(anchor)
            after_cursor.gotoRange(anchor.getEnd(), True)
            _log(
                "_reapply_accepted_paragraph_formatting: done "
                f"bookmark={name} after={_format_snapshot(after_cursor)} "
                f"text='{_safe_text_preview(after_cursor)}'"
            )
        except Exception as exc:
            _log(
                "_reapply_accepted_paragraph_formatting: failed for "
                f"{name}: {exc}"
            )


def _parse_color(value: str) -> Optional[int]:
    raw = value.strip()
    if raw.startswith("#"):
        raw = raw[1:]
    try:
        return int(raw, 16)
    except Exception:
        return None


def _spec_has_explicit_run_color(spec: Dict[str, Any]) -> bool:
    runs = (spec or {}).get("runs") or []
    if not isinstance(runs, list):
        return False
    for run in runs:
        if not isinstance(run, dict):
            continue
        fmt = run.get("format") or {}
        if isinstance(fmt, dict) and fmt.get("font_color_rgb"):
            return True
    return False


def _any_paragraph_spec_has_explicit_color(proposed_spec: Dict[str, Any]) -> bool:
    """Check if any paragraph in the proposed spec has an explicit run color."""
    paragraph_specs = _collect_paragraph_specs(proposed_spec)
    return any(_spec_has_explicit_run_color(p) for p in paragraph_specs)


def _collect_paragraph_specs(spec: Dict[str, Any]) -> List[Dict[str, Any]]:
    if not isinstance(spec, dict):
        return []
    paragraphs = spec.get("paragraphs")
    if isinstance(paragraphs, list):
        return [p for p in paragraphs if isinstance(p, dict)]
    return [spec]


def _apply_explicit_run_colors(text, start_range, runs: List[Dict[str, Any]]) -> None:
    if not runs:
        return
    walker = text.createTextCursorByRange(start_range)
    walker.collapseToStart()
    for run in runs:
        if not isinstance(run, dict):
            continue
        run_text = run.get("text") or ""
        if not run_text:
            continue
        fmt = run.get("format") or {}
        if not isinstance(fmt, dict):
            continue
        color_value = fmt.get("font_color_rgb")
        parsed = _parse_color(color_value) if color_value else None

        run_start = walker.getStart()
        try:
            moved = walker.goRight(len(run_text), False)
        except Exception:
            moved = False
        if not moved:
            break
        run_end = walker.getStart()

        if parsed is None:
            continue
        try:
            run_cursor = text.createTextCursorByRange(run_start)
            run_cursor.gotoRange(run_end, True)
            run_cursor.CharColor = parsed
        except Exception:
            continue


def _reapply_explicit_colors_for_paragraph_specs(
    doc, bookmark_names: List[str], proposed_spec: Dict[str, Any]
) -> None:
    paragraph_specs = _collect_paragraph_specs(proposed_spec)
    if not paragraph_specs:
        return
    bookmarks = doc.getBookmarks()
    for idx, name in enumerate(bookmark_names):
        if not _validate_bookmark_exists(doc, name):
            continue
        try:
            if len(paragraph_specs) == 1:
                paragraph_spec = paragraph_specs[0]
            elif idx < len(paragraph_specs):
                paragraph_spec = paragraph_specs[idx]
            else:
                continue
            if not _spec_has_explicit_run_color(paragraph_spec):
                continue
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            runs = paragraph_spec.get("runs") or []
            if text is None or not isinstance(runs, list):
                continue
            _apply_explicit_run_colors(text, anchor, runs)
        except Exception as exc:
            _log(f"_reapply_explicit_colors_for_paragraph_specs: failed for {name}: {exc}")


DEFAULT_TABLE_HEADER_FILL = "#2E75B6"
DEFAULT_TABLE_STRIPE_FILL = "#EAF2F8"
DEFAULT_TABLE_BODY_FILL = "#FFFFFF"
DEFAULT_TABLE_FONT_NAME = "Calibri"
DEFAULT_TABLE_FONT_SIZE_PT = 11


def _merge_missing(base: Dict[str, Any], defaults: Dict[str, Any]) -> Dict[str, Any]:
    merged = dict(base or {})
    for key, value in defaults.items():
        if merged.get(key) is None:
            merged[key] = value
    return merged


def _table_cell_text(cell_spec: Dict[str, Any]) -> str:
    text = cell_spec.get("text")
    if text is not None:
        return str(text)
    runs = cell_spec.get("runs")
    if isinstance(runs, list):
        return "".join(str(run.get("text") or "") for run in runs if isinstance(run, dict))
    paragraph = cell_spec.get("paragraph")
    if isinstance(paragraph, dict):
        return _table_cell_text(paragraph)
    content = cell_spec.get("content")
    if isinstance(content, list):
        parts = []
        for item in content:
            if not isinstance(item, dict):
                continue
            inner = item.get("paragraph") if isinstance(item.get("paragraph"), dict) else item
            parts.append(_table_cell_text(inner))
        return " ".join(part for part in parts if part)
    return ""


def _default_table_column_widths(rows: List[Dict[str, Any]], cols: int) -> List[float]:
    if cols <= 0:
        return []
    max_lengths = [0] * cols
    for row in rows:
        for c_idx, cell in enumerate((row or {}).get("cells") or []):
            if c_idx >= cols or not isinstance(cell, dict):
                continue
            max_lengths[c_idx] = max(max_lengths[c_idx], len(_table_cell_text(cell)))

    widths = []
    for max_len in max_lengths:
        widths.append(min(2.35, max(0.65, 0.45 + min(max_len, 24) * 0.075)))

    total = sum(widths)
    max_total = 6.2
    if total > max_total:
        scale = max_total / total
        widths = [max(0.55, width * scale) for width in widths]
    return [round(width, 2) for width in widths]


def _normalize_table_run(run: Dict[str, Any], is_header: bool) -> Dict[str, Any]:
    normalized = dict(run or {})
    normalized["format"] = _merge_missing(
        normalized.get("format") or {},
        {
            "font_name": DEFAULT_TABLE_FONT_NAME,
            "font_size_pt": DEFAULT_TABLE_FONT_SIZE_PT,
            "bold": is_header,
            "italic": False,
            "underline": False,
            "font_color_rgb": "#FFFFFF" if is_header else "#000000",
        },
    )
    return normalized


def _normalize_table_paragraph(paragraph: Dict[str, Any], is_header: bool) -> Dict[str, Any]:
    normalized = dict(paragraph or {})
    normalized["format"] = _merge_missing(
        normalized.get("format") or {},
        {
            "style_name": "Normal",
            "alignment": "left",
            "space_before_pt": 0.01,
            "space_after_pt": 0.01,
        },
    )
    runs = normalized.get("runs")
    if isinstance(runs, list) and runs:
        normalized["runs"] = [
            _normalize_table_run(run, is_header)
            for run in runs
            if isinstance(run, dict)
        ]
    elif normalized.get("text") is not None:
        normalized["runs"] = [_normalize_table_run({"text": str(normalized.get("text") or "")}, is_header)]
        normalized.pop("text", None)
    else:
        normalized["runs"] = [_normalize_table_run({"text": ""}, is_header)]
    return normalized


def _normalize_table_cell_spec(
    cell_spec: Dict[str, Any],
    row_idx: int,
    col_idx: int,
    column_widths: List[float],
) -> Dict[str, Any]:
    normalized = dict(cell_spec or {})
    is_header = row_idx == 0

    for border_key in ("border_top", "border_bottom", "border_left", "border_right"):
        if normalized.get(border_key) is None:
            normalized[border_key] = "single"
    if normalized.get("vertical_align") is None:
        normalized["vertical_align"] = "center"
    if normalized.get("width_in") is None and col_idx < len(column_widths):
        normalized["width_in"] = column_widths[col_idx]
    if normalized.get("shading_color") is None:
        if is_header:
            normalized["shading_color"] = DEFAULT_TABLE_HEADER_FILL
        else:
            normalized["shading_color"] = (
                DEFAULT_TABLE_STRIPE_FILL if row_idx > 0 and row_idx % 2 == 0 else DEFAULT_TABLE_BODY_FILL
            )

    content = normalized.get("content")
    if isinstance(content, list) and content:
        normalized_content = []
        for item in content:
            if not isinstance(item, dict):
                continue
            if isinstance(item.get("paragraph"), dict):
                merged = dict(item)
                merged["paragraph"] = _normalize_table_paragraph(item["paragraph"], is_header)
                normalized_content.append(merged)
            elif _looks_like_paragraph_spec(item):
                normalized_content.append({"paragraph": _normalize_table_paragraph(item, is_header)})
        if normalized_content:
            normalized["content"] = normalized_content
            return normalized

    paragraph = normalized.get("paragraph")
    if isinstance(paragraph, dict):
        normalized["paragraph"] = _normalize_table_paragraph(paragraph, is_header)
        return normalized

    runs = normalized.get("runs")
    if isinstance(runs, list) and runs:
        normalized["paragraph"] = {
            "format": {
                "style_name": "Normal",
                "alignment": "left",
                "space_before_pt": 0.01,
                "space_after_pt": 0.01,
            },
            "runs": [_normalize_table_run(run, is_header) for run in runs if isinstance(run, dict)],
        }
        normalized.pop("runs", None)
        return normalized

    text = normalized.get("text")
    normalized["paragraph"] = {
        "format": {
            "style_name": "Normal",
            "alignment": "left",
            "space_before_pt": 0.01,
            "space_after_pt": 0.01,
        },
        "runs": [_normalize_table_run({"text": "" if text is None else str(text)}, is_header)],
    }
    normalized.pop("text", None)
    return normalized


def _normalize_table_spec(spec: Dict[str, Any]) -> Dict[str, Any]:
    normalized = dict(spec or {})
    rows = normalized.get("rows") or []
    if not rows:
        return normalized
    cols = len((rows[0] or {}).get("cells") or [])
    first_row_cells = (rows[0] or {}).get("cells") or []
    has_widths = any(isinstance(cell, dict) and cell.get("width_in") is not None for cell in first_row_cells)
    column_widths = [
        float(cell.get("width_in"))
        for cell in first_row_cells
        if isinstance(cell, dict) and cell.get("width_in") is not None
    ]
    if not has_widths or len(column_widths) != cols:
        column_widths = _default_table_column_widths(rows, cols)

    normalized_rows = []
    for row_idx, row in enumerate(rows):
        cells = (row or {}).get("cells") or []
        normalized_rows.append(
            {
                **dict(row or {}),
                "cells": [
                    _normalize_table_cell_spec(cell, row_idx, col_idx, column_widths)
                    for col_idx, cell in enumerate(cells)
                ],
            }
        )
    normalized["rows"] = normalized_rows
    return normalized


def _table_spec_has_explicit_run_color(spec: Dict[str, Any]) -> bool:
    rows = (spec or {}).get("rows") or []
    if not isinstance(rows, list):
        return False
    for row in rows:
        cells = (row or {}).get("cells") or []
        if not isinstance(cells, list):
            continue
        for cell in cells:
            if not isinstance(cell, dict):
                continue
            paragraph_spec = cell.get("paragraph")
            if isinstance(paragraph_spec, dict) and _spec_has_explicit_run_color(
                paragraph_spec
            ):
                return True
            content_spec = cell.get("content")
            if isinstance(content_spec, list):
                for item in content_spec:
                    if not isinstance(item, dict):
                        continue
                    paragraph_item = item.get("paragraph") if isinstance(item.get("paragraph"), dict) else item
                    if _looks_like_paragraph_spec(paragraph_item):
                        if _spec_has_explicit_run_color(paragraph_item):
                            return True
            runs_spec = cell.get("runs")
            if isinstance(runs_spec, list) and _spec_has_explicit_run_color(
                {"runs": runs_spec}
            ):
                return True
    return False


def _looks_like_paragraph_spec(spec: Any) -> bool:
    if not isinstance(spec, dict):
        return False
    return any(key in spec for key in ("runs", "text", "format"))


def _paragraph_specs_from_cell(cell_spec: Dict[str, Any]) -> List[Dict[str, Any]]:
    paragraph_spec = cell_spec.get("paragraph")
    if isinstance(paragraph_spec, dict):
        return [paragraph_spec]

    content_spec = cell_spec.get("content")
    if isinstance(content_spec, list):
        paragraph_specs = []
        for item in content_spec:
            if not isinstance(item, dict):
                continue
            if isinstance(item.get("paragraph"), dict):
                paragraph_specs.append(item["paragraph"])
            elif _looks_like_paragraph_spec(item):
                paragraph_specs.append(item)
        return paragraph_specs

    runs_spec = cell_spec.get("runs")
    if isinstance(runs_spec, list):
        return [{"runs": runs_spec}]

    text_spec = cell_spec.get("text")
    if text_spec is not None:
        return [{"text": str(text_spec)}]

    return []


def _apply_explicit_colors_to_table(table, table_spec: Dict[str, Any]) -> None:
    rows = (table_spec or {}).get("rows") or []
    if not isinstance(rows, list):
        return
    for r_idx, row in enumerate(rows):
        cells = (row or {}).get("cells") or []
        if not isinstance(cells, list):
            continue
        for c_idx, cell_spec in enumerate(cells):
            if not isinstance(cell_spec, dict):
                continue
            paragraph_specs = _paragraph_specs_from_cell(cell_spec)
            paragraph_specs = [
                spec for spec in paragraph_specs if _spec_has_explicit_run_color(spec)
            ]
            if not paragraph_specs:
                continue

            try:
                cell = table.getCellByPosition(c_idx, r_idx)
                cell_text = cell.getText()
                cell_cursor = cell_text.createTextCursor()
                cell_cursor.gotoStart(False)
                for paragraph_spec in paragraph_specs:
                    start_range = cell_cursor.getStart()
                    _apply_explicit_run_colors(
                        cell_text, start_range, paragraph_spec.get("runs") or []
                    )
                    if not cell_cursor.gotoEndOfParagraph(False):
                        break
                    cell_cursor.goRight(1, False)
            except Exception:
                continue


def _build_border_line(style_name: Optional[str]):
    if not style_name or str(style_name).lower() == "none":
        return None
    try:
        import uno

        border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
        border.Color = 0xD9D9D9
        border.LineWidth = 35
        border.LineStyle = 0
        return border
    except Exception:
        return None


def _apply_table_cell_formatting(cell, cell_cursor, cell_spec: Dict[str, Any]) -> None:
    if not isinstance(cell_spec, dict):
        return

    shading_color = _parse_color(cell_spec.get("shading_color") or "")
    if shading_color is not None:
        try:
            cell.setPropertyValue("BackTransparent", False)
        except Exception:
            pass
        try:
            cell.setPropertyValue("BackColor", shading_color)
        except Exception:
            try:
                cell.BackColor = shading_color
            except Exception:
                pass
        try:
            para_cursor = cell_cursor.getText().createTextCursor()
            para_cursor.gotoStart(False)
            para_cursor.gotoEnd(True)
            para_cursor.ParaBackColor = shading_color
        except Exception:
            pass

    vertical_align = str(cell_spec.get("vertical_align") or "").lower()
    vert_map = {"top": 0, "center": 1, "bottom": 2}
    if vertical_align in vert_map:
        try:
            cell.setPropertyValue("VertOrient", vert_map[vertical_align])
        except Exception:
            pass

    for side, key in (
        ("TopBorder", "border_top"),
        ("BottomBorder", "border_bottom"),
        ("LeftBorder", "border_left"),
        ("RightBorder", "border_right"),
    ):
        border = _build_border_line(cell_spec.get(key))
        if border is None:
            continue
        try:
            cell.setPropertyValue(side, border)
        except Exception:
            pass

    width_in = cell_spec.get("width_in")
    if width_in is not None:
        try:
            cell.setPropertyValue("Width", int(float(width_in) * 2540))
        except Exception:
            pass


def _reapply_explicit_colors_for_table_spec(
    doc, bookmark_names: List[str], table_spec: Dict[str, Any]
) -> None:
    if not _table_spec_has_explicit_run_color(table_spec):
        return
    for name in bookmark_names:
        try:
            element = _get_element_by_bookmark(doc, name)
        except Exception:
            element = None
        if not element or element[0] != "table":
            continue
        try:
            _apply_explicit_colors_to_table(element[1], table_spec)
        except Exception as exc:
            _log(f"_reapply_explicit_colors_for_table_spec: failed for {name}: {exc}")


def _validate_bookmark_exists(doc, bookmark_name: str) -> bool:
    """Validate that a bookmark exists and is accessible."""
    if not bookmark_name:
        return False
    try:
        bookmarks = doc.getBookmarks()
        if not bookmarks.hasByName(bookmark_name):
            return False
        bookmark = bookmarks.getByName(bookmark_name)
        anchor = bookmark.getAnchor()
        if anchor is None:
            return False
        text = anchor.getText()
        if text is None:
            return False
        return True
    except Exception as e:
        _log(f"_validate_bookmark_exists: error checking {bookmark_name}: {e}")
        return False


def _is_sdoc_bookmark_name(bookmark_name: Optional[str]) -> bool:
    return isinstance(bookmark_name, str) and bookmark_name.startswith("sdoc_")


def _rebind_target_bookmark_from_preview(
    doc, target_bookmark: Optional[str], preview_bookmark: Optional[str]
) -> bool:
    """
    Re-anchor a stable sdoc_* bookmark onto accepted preview content so
    follow-up edits can continue using the same element_id.
    """
    if not _is_sdoc_bookmark_name(target_bookmark):
        return False
    if not isinstance(preview_bookmark, str) or not preview_bookmark:
        return False
    if target_bookmark == preview_bookmark:
        return True

    try:
        bookmarks = doc.getBookmarks()
        if not bookmarks.hasByName(preview_bookmark):
            _log(
                f"_rebind_target_bookmark_from_preview: source bookmark missing ({preview_bookmark})"
            )
            return False

        preview = bookmarks.getByName(preview_bookmark)
        anchor = preview.getAnchor()
        text = anchor.getText()
        if text is None:
            _log(
                f"_rebind_target_bookmark_from_preview: source anchor has no text ({preview_bookmark})"
            )
            return False

        cursor = None
        try:
            cursor = text.createTextCursorByRange(anchor)
            cursor.collapseToStart()
        except Exception:
            try:
                start_pos = anchor.getStart()
                cursor = text.createTextCursorByRange(start_pos)
                cursor.collapseToStart()
            except Exception:
                cursor = None

        if cursor is None:
            _log(
                f"_rebind_target_bookmark_from_preview: could not create cursor from {preview_bookmark}"
            )
            return False

        # Remove any existing target bookmark(s) before inserting the rebound one.
        # Guard strictly: never insert if we cannot clear previous target name.
        try:
            dispose_rounds = 0
            while bookmarks.hasByName(target_bookmark) and dispose_rounds < 8:
                bookmarks.getByName(target_bookmark).dispose()
                dispose_rounds += 1
            if bookmarks.hasByName(target_bookmark):
                _log(
                    f"_rebind_target_bookmark_from_preview: target still exists after dispose loop ({target_bookmark}); aborting rebind to avoid duplicate names"
                )
                return False
        except Exception as dispose_exc:
            _log(
                f"_rebind_target_bookmark_from_preview: failed disposing old target {target_bookmark}: {dispose_exc}"
            )
            try:
                if bookmarks.hasByName(target_bookmark):
                    _log(
                        f"_rebind_target_bookmark_from_preview: target remains after dispose error ({target_bookmark}); aborting rebind"
                    )
                    return False
            except Exception:
                return False

        bookmark = doc.createInstance("com.sun.star.text.Bookmark")
        bookmark.setName(target_bookmark)
        text.insertTextContent(cursor, bookmark, True)
        if not bookmarks.hasByName(target_bookmark):
            _log(
                f"_rebind_target_bookmark_from_preview: insert did not materialize target bookmark ({target_bookmark})"
            )
            return False
        try:
            rebound_anchor = bookmarks.getByName(target_bookmark).getAnchor()
            if rebound_anchor is None:
                _log(
                    f"_rebind_target_bookmark_from_preview: rebound bookmark has no anchor ({target_bookmark})"
                )
                return False
        except Exception as anchor_exc:
            _log(
                f"_rebind_target_bookmark_from_preview: failed to verify rebound bookmark {target_bookmark}: {anchor_exc}"
            )
            return False
        _log(
            f"_rebind_target_bookmark_from_preview: rebound {target_bookmark} -> {preview_bookmark}"
        )
        return True
    except Exception as e:
        _log(f"_rebind_target_bookmark_from_preview: failed: {e}")
        return False


def _get_table_by_name(doc, table_name: Optional[str]):
    if not isinstance(table_name, str) or not table_name:
        return None
    try:
        tables = doc.getTextTables()
        if tables.hasByName(table_name):
            return tables.getByName(table_name)
    except Exception:
        return None
    return None


def _safe_table_name(table) -> Optional[str]:
    try:
        name = table.getName()
        if isinstance(name, str) and name:
            return name
    except Exception:
        pass
    return None


def _table_spec_has_rows(spec: Dict[str, Any]) -> bool:
    rows = (spec or {}).get("rows") or []
    if not isinstance(rows, list) or not rows:
        return False
    for row in rows:
        if not isinstance(row, dict):
            continue
        cells = row.get("cells") or []
        if isinstance(cells, list) and cells:
            return True
    return False


def _resolve_table_for_replacement(doc, op_state: Dict[str, Any]):
    target_bookmark = op_state.get("target_bookmark")
    table = _resolve_table_from_bookmark(doc, target_bookmark)
    if table is not None:
        return table

    original_table_name = op_state.get("original_table_name")
    table = _get_table_by_name(doc, original_table_name)
    if table is not None:
        return table

    for name in op_state.get("marked_ids") or []:
        table = _resolve_table_from_bookmark(doc, name)
        if table is not None:
            return table

    target_index = op_state.get("target_index")
    if isinstance(target_index, int):
        try:
            element = _get_element_by_index(doc, target_index)
            if element and element[0] == "table":
                return element[1]
        except Exception:
            pass
    return None


def _resize_table_to_spec(table, row_count: int, col_count: int) -> bool:
    if row_count <= 0 or col_count <= 0:
        return False
    try:
        rows = table.getRows()
        cols = table.getColumns()
        current_rows = rows.getCount()
        current_cols = cols.getCount()
    except Exception:
        return False

    try:
        if current_rows < row_count:
            rows.insertByIndex(current_rows, row_count - current_rows)
        elif current_rows > row_count:
            rows.removeByIndex(row_count, current_rows - row_count)

        if current_cols < col_count:
            cols.insertByIndex(current_cols, col_count - current_cols)
        elif current_cols > col_count:
            cols.removeByIndex(col_count, current_cols - col_count)
        return True
    except Exception as exc:
        _log(f"_resize_table_to_spec: failed: {exc}")
        return False


def _apply_table_spec_to_table(table, table_spec: Dict[str, Any]) -> bool:
    table_spec = _normalize_table_spec(table_spec)
    rows = (table_spec or {}).get("rows") or []
    if not isinstance(rows, list) or not rows:
        return False

    row_count = len(rows)
    col_count = 0
    for row in rows:
        cells = (row or {}).get("cells") or []
        if isinstance(cells, list):
            col_count = max(col_count, len(cells))
    if col_count <= 0:
        return False

    if not _resize_table_to_spec(table, row_count, col_count):
        return False

    try:
        for r_idx in range(row_count):
            row_cells = (rows[r_idx] or {}).get("cells") or []
            for c_idx in range(col_count):
                cell_spec = row_cells[c_idx] if c_idx < len(row_cells) else {}
                if not isinstance(cell_spec, dict):
                    cell_spec = {}

                cell = table.getCellByPosition(c_idx, r_idx)
                if cell_spec.get("style"):
                    try:
                        cell.CellStyleName = cell_spec["style"]
                    except Exception:
                        pass

                cell_text = cell.getText()
                cell_text.setString("")
                cell_cursor = cell_text.createTextCursor()
                paragraph_specs = _paragraph_specs_from_cell(cell_spec)
                if paragraph_specs:
                    for idx, paragraph_spec in enumerate(paragraph_specs):
                        _write_paragraph_spec(
                            cell_cursor,
                            paragraph_spec,
                            color=None,
                            append_break=idx < len(paragraph_specs) - 1,
                        )
                else:
                    cell_text.insertString(cell_cursor, "", False)
                _apply_table_cell_formatting(cell, cell_cursor, cell_spec)
        first_row = (rows[0] or {}).get("cells") or []
        if first_row:
            try:
                columns = table.getColumns()
                for c_idx, cell_spec in enumerate(first_row):
                    width_in = (cell_spec or {}).get("width_in")
                    if width_in is None:
                        continue
                    try:
                        column = columns.getByIndex(c_idx)
                        column.Width = int(float(width_in) * 2540)
                    except Exception:
                        continue
            except Exception:
                pass
        return True
    except Exception as exc:
        _log(f"_apply_table_spec_to_table: failed: {exc}")
        return False


def _remove_table_by_name(doc, table_name: Optional[str]) -> bool:
    table = _get_table_by_name(doc, table_name)
    if table is None:
        return False
    try:
        table.dispose()
        return True
    except Exception as exc:
        _log(f"_remove_table_by_name: failed to dispose {table_name}: {exc}")
        return False


def _normalize_table_object(table) -> bool:
    try:
        rows = table.getRows().getCount()
        cols = table.getColumns().getCount() if rows else 0
    except Exception:
        return False

    for r_idx in range(rows):
        for c_idx in range(cols):
            try:
                cell = table.getCellByPosition(c_idx, r_idx)
                cell_text = cell.getText()
                cursor = cell_text.createTextCursor()
                cursor.gotoStart(False)
                cursor.gotoEnd(True)
                try:
                    cursor.setPropertyToDefault("CharColor")
                except Exception:
                    try:
                        cursor.CharColor = 0
                    except Exception:
                        pass
                try:
                    cursor.setPropertyToDefault("CharStrikeout")
                except Exception:
                    try:
                        cursor.CharStrikeout = 0
                    except Exception:
                        pass
            except Exception:
                continue
    return True


def _normalize_table_by_name(doc, table_name: Optional[str]) -> bool:
    table = _get_table_by_name(doc, table_name)
    if table is None:
        return False
    return _normalize_table_object(table)


def _create_table_anchor_cursor(doc, table):
    try:
        anchor = table.getAnchor()
    except Exception:
        anchor = None

    if anchor is not None:
        text_candidates = []
        try:
            anchor_text = anchor.getText()
            if anchor_text is not None:
                text_candidates.append(anchor_text)
        except Exception:
            pass
        try:
            doc_text = doc.getText()
            if doc_text is not None:
                text_candidates.append(doc_text)
        except Exception:
            pass

        for text in text_candidates:
            try:
                cursor = text.createTextCursorByRange(anchor)
                cursor.collapseToStart()
                return text, cursor
            except Exception:
                pass
            try:
                start_pos = anchor.getStart()
                cursor = text.createTextCursorByRange(start_pos)
                cursor.collapseToStart()
                return text, cursor
            except Exception:
                pass

    return None, None


def _rebind_target_bookmark_from_table(
    doc, target_bookmark: Optional[str], table_name: Optional[str]
) -> bool:
    if not _is_sdoc_bookmark_name(target_bookmark):
        return False
    if not isinstance(table_name, str) or not table_name:
        return False

    table = _get_table_by_name(doc, table_name)
    if table is None:
        _log(f"_rebind_target_bookmark_from_table: table missing ({table_name})")
        return False

    text, cursor = _create_table_anchor_cursor(doc, table)
    if text is None or cursor is None:
        _log(
            "_rebind_target_bookmark_from_table: could not create cursor "
            f"for table {table_name}"
        )
        return False

    try:
        bookmarks = doc.getBookmarks()
        try:
            dispose_rounds = 0
            while bookmarks.hasByName(target_bookmark) and dispose_rounds < 8:
                bookmarks.getByName(target_bookmark).dispose()
                dispose_rounds += 1
            if bookmarks.hasByName(target_bookmark):
                _log(
                    "_rebind_target_bookmark_from_table: target still exists after "
                    f"dispose loop ({target_bookmark})"
                )
                return False
        except Exception as dispose_exc:
            _log(
                "_rebind_target_bookmark_from_table: failed disposing old target "
                f"{target_bookmark}: {dispose_exc}"
            )
            return False

        bookmark = doc.createInstance("com.sun.star.text.Bookmark")
        bookmark.setName(target_bookmark)
        text.insertTextContent(cursor, bookmark, True)
        if not bookmarks.hasByName(target_bookmark):
            _log(
                "_rebind_target_bookmark_from_table: insert did not materialize "
                f"target bookmark ({target_bookmark})"
            )
            return False
        _log(
            "_rebind_target_bookmark_from_table: rebound "
            f"{target_bookmark} -> {table_name}"
        )
        return True
    except Exception as exc:
        _log(f"_rebind_target_bookmark_from_table: failed: {exc}")
        return False


def _remove_annotations_in_paragraph(doc, paragraph) -> None:
    """Remove all Annotation text fields anchored within the given paragraph."""
    try:
        para_text = paragraph.getText()
        para_start = paragraph.getStart()
        para_end = paragraph.getEnd()
    except Exception as exc:
        _log(f"_remove_annotations_in_paragraph: cannot get paragraph bounds: {exc}")
        return

    to_remove = []
    try:
        text_fields = doc.getTextFields()
        enum = text_fields.createEnumeration()
        while enum.hasMoreElements():
            field = enum.nextElement()
            try:
                if not field.supportsService("com.sun.star.text.textfield.Annotation"):
                    continue
                field_anchor = field.getAnchor()
                field_text = field_anchor.getText()
                if field_text != para_text:
                    continue
                start_cmp = para_text.compareRegionStarts(field_anchor, para_start)
                end_cmp = para_text.compareRegionEnds(field_anchor, para_end)
                if start_cmp >= 0 and end_cmp <= 0:
                    to_remove.append(field)
            except Exception:
                continue
    except Exception as exc:
        _log(f"_remove_annotations_in_paragraph: enumeration failed: {exc}")
        return

    for field in to_remove:
        try:
            anchor = field.getAnchor()
            anchor.getText().removeTextContent(field)
            _log("_remove_annotations_in_paragraph: removed annotation")
        except Exception as exc:
            _log(f"_remove_annotations_in_paragraph: failed to remove annotation: {exc}")


def _delete_paragraph_node(doc, paragraph) -> None:
    """Delete a paragraph node from the document including the paragraph break."""
    try:
        text = paragraph.getText()
        cursor = text.createTextCursorByRange(paragraph.getStart())
        cursor.gotoRange(paragraph.getEnd(), True)

        end_probe = text.createTextCursorByRange(paragraph.getEnd())
        if end_probe.goRight(1, False):
            cursor.gotoRange(end_probe, True)
            cursor.String = ""
            _log("_delete_paragraph_node: removed via forward merge")
            return

        start_probe = text.createTextCursorByRange(paragraph.getStart())
        if start_probe.goLeft(1, False):
            cursor = text.createTextCursorByRange(start_probe)
            cursor.gotoRange(paragraph.getEnd(), True)
            cursor.String = ""
            _log("_delete_paragraph_node: removed via backward merge")
            return

        cursor.String = ""
        _log("_delete_paragraph_node: sole paragraph cleared (cannot remove node)")

    except Exception as exc:
        _log(f"_delete_paragraph_node: failed: {exc}")


def _remove_bookmarked_segments(doc, bookmark_names: List[str]) -> None:
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(
                f"_remove_bookmarked_segments: bookmark {name} not found or invalid; skipping."
            )
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)
            cursor.String = ""
            bookmark.dispose()
        except Exception as e:
            _log(f"_remove_bookmarked_segments: failed to remove {name}: {e}")
            # Try to dispose bookmark even if removal failed
            try:
                if bookmarks.hasByName(name):
                    bookmarks.getByName(name).dispose()
            except Exception:
                pass


def _resolve_table_from_bookmark(doc, bookmark_name: Optional[str]):
    if not bookmark_name:
        return None
    try:
        bookmarks = doc.getBookmarks()
        if not bookmarks.hasByName(bookmark_name):
            return None
        bookmark = bookmarks.getByName(bookmark_name)
        anchor = bookmark.getAnchor()
    except Exception:
        return None

    # Fast path: many table/cell anchors expose owning table directly.
    try:
        owner = anchor.getTextTable()
        if owner is not None:
            return owner
    except Exception:
        pass

    try:
        anchor_text = anchor.getText()
    except Exception:
        anchor_text = None

    try:
        doc_text = doc.getText()
    except Exception:
        doc_text = None

    try:
        tables = doc.getTextTables()
        table_names = list(tables.getElementNames())
    except Exception:
        table_names = []
        tables = None

    for table_name in table_names:
        try:
            table = tables.getByName(table_name)
        except Exception:
            continue

        # Direct anchor compare.
        try:
            table_anchor = table.getAnchor()
            if table_anchor == anchor:
                return table
        except Exception:
            table_anchor = None

        # Region compare in document body text, if available.
        if doc_text is not None and table_anchor is not None:
            try:
                if doc_text.compareRegionStarts(table_anchor, anchor) == 0:
                    return table
            except Exception:
                pass
            try:
                if doc_text.compareRegionStarts(table_anchor.getStart(), anchor.getStart()) == 0:
                    return table
            except Exception:
                pass

        # Cell-text ownership fallback for anchors that live inside table cells.
        if anchor_text is None:
            continue
        try:
            for cell_name in table.getCellNames():
                try:
                    cell = table.getCellByName(cell_name)
                    if cell.getText() == anchor_text:
                        return table
                except Exception:
                    continue
        except Exception:
            continue

    return None


def _remove_table_by_target(
    doc, target_bookmark: Optional[str], target_index: Optional[int]
) -> bool:
    table = _resolve_table_from_bookmark(doc, target_bookmark)
    if table is None and isinstance(target_index, int):
        try:
            element = _get_element_by_index(doc, target_index)
            if element and element[0] == "table":
                table = element[1]
        except Exception:
            table = None

    if table is None:
        return False

    try:
        table.dispose()
        return True
    except Exception as exc:
        _log(f"_remove_table_by_target: failed to dispose target table: {exc}")
        return False


def _remove_bookmarked_tables(doc, bookmark_names: List[str]) -> None:
    """
    Remove table content anchored by preview bookmarks.
    Uses table-aware deletion first to avoid leaving old table previews behind.
    Falls back to generic segment removal when bookmark does not resolve to a table.
    """
    if not bookmark_names:
        return

    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(f"_remove_bookmarked_tables: bookmark {name} not found or invalid")
            continue

        table = None
        bookmark = None

        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
        except Exception as exc:
            _log(f"_remove_bookmarked_tables: failed to read bookmark {name}: {exc}")
            continue

        # Preferred path: direct table owner from anchor.
        try:
            table = anchor.getTextTable()
        except Exception:
            table = None

        # Fallback 1: existing bookmark resolver.
        if table is None:
            try:
                element = _get_element_by_bookmark(doc, name)
                if element and element[0] == "table":
                    table = element[1]
            except Exception:
                table = None

        # Fallback 2: scan text tables and match by anchor/cell text container.
        if table is None:
            try:
                anchor_text = anchor.getText()
            except Exception:
                anchor_text = None
            try:
                tables = doc.getTextTables()
                for table_name in list(tables.getElementNames()):
                    try:
                        candidate = tables.getByName(table_name)
                    except Exception:
                        continue
                    try:
                        if candidate.getAnchor() == anchor:
                            table = candidate
                            break
                    except Exception:
                        pass
                    if anchor_text is None:
                        continue
                    try:
                        for cell_name in candidate.getCellNames():
                            try:
                                cell = candidate.getCellByName(cell_name)
                                if cell.getText() == anchor_text:
                                    table = candidate
                                    break
                            except Exception:
                                continue
                        if table is not None:
                            break
                    except Exception:
                        continue
            except Exception:
                pass

        if table is None:
            _log(
                f"_remove_bookmarked_tables: no table resolved for {name}; skipping table removal"
            )
            _dispose_bookmarks(doc, [name])
            continue

        removed_table = False
        try:
            table.dispose()
            removed_table = True
        except Exception as exc:
            _log(f"_remove_bookmarked_tables: failed to dispose table for {name}: {exc}")

        if removed_table:
            try:
                if bookmarks.hasByName(name):
                    bookmarks.getByName(name).dispose()
            except Exception:
                pass
        else:
            _dispose_bookmarks(doc, [name])


def _normalize_green_segments(
    doc, bookmark_names: List[str], dispose_bookmarks: bool = True
) -> None:
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(f"_normalize_green_segments: bookmark {name} not found or invalid")
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)
            before = _format_snapshot(cursor)
            preview = _safe_text_preview(cursor)

            # Remove preview-only artifacts while preserving intentional
            # formatting (italic/bold/underline) from the proposed content.
            cursor.setPropertyToDefault("CharColor")
            cursor.setPropertyToDefault("CharStrikeout")
            _log(
                "_normalize_green_segments: normalized "
                f"bookmark={name} dispose={dispose_bookmarks} "
                f"before={before} after={_format_snapshot(cursor)} "
                f"text='{preview}'"
            )

            if dispose_bookmarks:
                bookmark.dispose()
        except Exception as e:
            _log(f"_normalize_green_segments: failed for {name}: {e}")
            # Try to dispose bookmark even if normalization failed
            try:
                if dispose_bookmarks and bookmarks.hasByName(name):
                    bookmarks.getByName(name).dispose()
            except Exception:
                pass


def _is_preview_green_color(value: Any) -> bool:
    try:
        return int(value) == PREVIEW_GREEN_RGB
    except Exception:
        return False


def _clear_preview_green_cursor(cursor, label: str) -> int:
    try:
        if not _is_preview_green_color(getattr(cursor, "CharColor", None)):
            return 0
        before = _format_snapshot(cursor)
        preview = _safe_text_preview(cursor)
        cursor.setPropertyToDefault("CharColor")
        cursor.setPropertyToDefault("CharStrikeout")
        _log(
            "_scrub_preview_green_formatting: cleared "
            f"label={label} before={before} after={_format_snapshot(cursor)} "
            f"text='{preview}'"
        )
        return 1
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: failed label={label}: {exc}")
        return 0


def _clear_preview_green_range(text_obj, start, end, label: str) -> int:
    try:
        cursor = text_obj.createTextCursorByRange(start)
        cursor.gotoRange(end, True)
        return _clear_preview_green_cursor(cursor, label)
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: failed range label={label}: {exc}")
        return 0


def _clear_preview_green_paragraph(paragraph, label: str) -> int:
    changed = 0
    scanned_portions = False
    try:
        text_obj = paragraph.getText()
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: paragraph text failed label={label}: {exc}")
        return changed

    try:
        enum = paragraph.createEnumeration()
        portion_index = 0
        while enum.hasMoreElements():
            portion = enum.nextElement()
            portion_index += 1
            try:
                changed += _clear_preview_green_range(
                    text_obj,
                    portion.getStart(),
                    portion.getEnd(),
                    f"{label}:portion:{portion_index}",
                )
                scanned_portions = True
            except Exception:
                continue
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: portion scan failed label={label}: {exc}")

    if not scanned_portions:
        try:
            changed += _clear_preview_green_range(
                text_obj,
                paragraph.getStart(),
                paragraph.getEnd(),
                label,
            )
        except Exception as exc:
            _log(f"_scrub_preview_green_formatting: paragraph fallback failed label={label}: {exc}")

    return changed


def _clear_preview_green_text(text_obj, label: str) -> int:
    changed = 0
    scanned_elements = False
    try:
        enum = text_obj.createEnumeration()
        index = 0
        while enum.hasMoreElements():
            element = enum.nextElement()
            index += 1
            try:
                if element.supportsService("com.sun.star.text.Paragraph"):
                    changed += _clear_preview_green_paragraph(element, f"{label}:paragraph:{index}")
                    scanned_elements = True
                elif element.supportsService("com.sun.star.text.TextTable"):
                    changed += _scrub_preview_green_table(element, f"{label}:table:{index}")
                    scanned_elements = True
            except Exception as exc:
                _log(f"_scrub_preview_green_formatting: text element failed label={label} index={index}: {exc}")
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: text scan failed label={label}: {exc}")

    if not scanned_elements:
        try:
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(True)
            changed += _clear_preview_green_cursor(cursor, label)
        except Exception as exc:
            _log(f"_scrub_preview_green_formatting: failed text label={label}: {exc}")

    return changed


def _scrub_preview_green_formatting(doc) -> int:
    """Final safety net: remove preview-green direct formatting still present after accept.

    Bookmark-based cleanup is the primary path. This only clears ranges whose
    direct CharColor is exactly the SmartDocs preview green.
    """
    changed = 0

    try:
        enum = doc.getText().createEnumeration()
        index = 0
        while enum.hasMoreElements():
            element = enum.nextElement()
            index += 1
            try:
                if element.supportsService("com.sun.star.text.Paragraph"):
                    changed += _clear_preview_green_paragraph(element, f"body_paragraph:{index}")
                elif element.supportsService("com.sun.star.text.TextTable"):
                    changed += _scrub_preview_green_table(element, f"body_table:{index}")
            except Exception as exc:
                _log(f"_scrub_preview_green_formatting: body element failed index={index}: {exc}")
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: body scan failed: {exc}")

    try:
        tables = doc.getTextTables()
        for table_name in list(tables.getElementNames()):
            try:
                changed += _scrub_preview_green_table(
                    tables.getByName(table_name),
                    f"text_table:{table_name}",
                )
            except Exception as exc:
                _log(f"_scrub_preview_green_formatting: table failed name={table_name}: {exc}")
    except Exception:
        pass

    try:
        page_styles = doc.getStyleFamilies().getByName("PageStyles")
        for style_name in list(page_styles.getElementNames()):
            try:
                style = page_styles.getByName(style_name)
            except Exception:
                continue
            for prop in (
                "HeaderText",
                "HeaderTextFirst",
                "HeaderTextLeft",
                "FooterText",
                "FooterTextFirst",
                "FooterTextLeft",
            ):
                try:
                    text_obj = getattr(style, prop, None)
                except Exception:
                    text_obj = None
                if text_obj is not None:
                    changed += _clear_preview_green_text(text_obj, f"page_style:{style_name}:{prop}")
    except Exception as exc:
        _log(f"_scrub_preview_green_formatting: page style scan failed: {exc}")

    return changed


def _scrub_preview_green_table(table, label: str) -> int:
    changed = 0
    try:
        rows = table.getRows().getCount()
        cols = table.getColumns().getCount()
    except Exception:
        return changed
    for row in range(rows):
        for col in range(cols):
            try:
                cell_text = table.getCellByPosition(col, row).getText()
                changed += _clear_preview_green_text(
                    cell_text,
                    f"{label}:cell:{row}:{col}",
                )
            except Exception:
                continue
    return changed


def _normalize_strikethrough_only(
    doc, bookmark_names: List[str], dispose_bookmarks: bool = True
) -> None:
    """Remove only strikethrough preview artifacts, preserving all character colors.

    Use this instead of _normalize_green_segments when the proposed spec has
    explicit run colors — in that case the text already has the user's intended
    colors (preview green was never applied) and resetting CharColor would
    destroy them.
    """
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(f"_normalize_strikethrough_only: bookmark {name} not found or invalid")
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)
            before = _format_snapshot(cursor)
            preview = _safe_text_preview(cursor)

            # Only remove strikethrough; leave CharColor untouched.
            cursor.setPropertyToDefault("CharStrikeout")
            _log(
                "_normalize_strikethrough_only: normalized "
                f"bookmark={name} dispose={dispose_bookmarks} "
                f"before={before} after={_format_snapshot(cursor)} "
                f"text='{preview}'"
            )

            if dispose_bookmarks:
                bookmark.dispose()
        except Exception as e:
            _log(f"_normalize_strikethrough_only: failed for {name}: {e}")
            try:
                if dispose_bookmarks and bookmarks.hasByName(name):
                    bookmarks.getByName(name).dispose()
            except Exception:
                pass


def _normalize_green_segments_preserve_bookmarks(
    doc, bookmark_names: List[str]
) -> None:
    """Normalize (remove preview colors) but preserve bookmarks for later re-insertion."""
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(
                f"_normalize_green_segments_preserve_bookmarks: bookmark {name} not found or invalid"
            )
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)

            # Remove preview-only artifacts while preserving intentional
            # formatting (italic/bold/underline) from the proposed content.
            # DON'T dispose bookmark - we need it for re-insertion
            cursor.setPropertyToDefault("CharColor")
            cursor.setPropertyToDefault("CharStrikeout")

            # Bookmark is preserved for _reinsert_as_tracked()
        except Exception as e:
            _log(
                f"_normalize_green_segments_preserve_bookmarks: failed for {name}: {e}"
            )


def _normalize_table_segments(doc, bookmark_names: List[str]) -> None:
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(
                f"_normalize_table_segments: bookmark {name} not found or invalid; skipping."
            )
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)

            # Remove preview-only artifacts while preserving intentional
            # formatting (italic/bold/underline) from the proposed content.
            cursor.setPropertyToDefault("CharColor")
            cursor.setPropertyToDefault("CharStrikeout")

            bookmark.dispose()
        except Exception as e:
            _log(f"_normalize_table_segments: failed for {name}: {e}")
            # Try to dispose bookmark even if normalization failed
            try:
                if bookmarks.hasByName(name):
                    bookmarks.getByName(name).dispose()
            except Exception:
                pass


def _normalize_table_segments_preserve_bookmarks(
    doc, bookmark_names: List[str]
) -> None:
    """Normalize (remove preview colors) but preserve bookmarks for later re-insertion."""
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(
                f"_normalize_table_segments_preserve_bookmarks: bookmark {name} not found or invalid; skipping."
            )
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)

            # Remove preview-only artifacts while preserving intentional
            # formatting (italic/bold/underline) from the proposed content.
            # DON'T dispose bookmark - we need it for re-insertion
            cursor.setPropertyToDefault("CharColor")
            cursor.setPropertyToDefault("CharStrikeout")

            # Bookmark is preserved for _reinsert_as_tracked()
        except Exception as e:
            _log(
                f"_normalize_table_segments_preserve_bookmarks: failed for {name}: {e}"
            )


def _restore_original_paragraphs(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    original_spec = op_state.get("original_spec")
    marked_ids = op_state.get("marked_ids") or []
    bookmarks = doc.getBookmarks()

    if original_spec and marked_ids:
        for name in marked_ids:
            if not _validate_bookmark_exists(doc, name):
                _log(
                    f"_restore_original_paragraphs: bookmark {name} not found or invalid; skipping."
                )
                continue
            try:
                bookmark = bookmarks.getByName(name)
                anchor = bookmark.getAnchor()
                text = anchor.getText()
                cursor = text.createTextCursorByRange(anchor)
                cursor.gotoRange(anchor.getEnd(), True)
                cursor.String = ""
                bookmark.dispose()
                _write_paragraph_spec(
                    cursor, original_spec, color=None, append_break=False
                )
            except Exception as e:
                _log(f"_restore_original_paragraphs: failed to restore {name}: {e}")
                # Try to dispose bookmark even if restore failed
                try:
                    if bookmarks.hasByName(name):
                        bookmarks.getByName(name).dispose()
                except Exception:
                    pass
    else:
        _reset_marked_segments(doc, marked_ids)


def _restore_original_table(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    _normalize_table_segments(doc, op_state.get("marked_ids") or [])
    original_table_name = op_state.get("original_table_name")
    if (
        isinstance(original_table_name, str)
        and original_table_name
        and _normalize_table_by_name(doc, original_table_name)
    ):
        _log(
            "_restore_original_table: normalized original table by name "
            f"{original_table_name}"
        )


def _create_cursor_before_state_anchor(doc, op_state: Dict[str, Any]):
    bookmark = op_state.get("target_bookmark")
    if bookmark:
        element = _get_element_by_bookmark(doc, bookmark)
        if element and element[0] == "paragraph":
            paragraph = element[1]
            text = paragraph.getText()
            cursor = text.createTextCursorByRange(paragraph.getStart())
            cursor.collapseToStart()
            return cursor
    index = op_state.get("target_index")
    if isinstance(index, int):
        element = _get_element_by_index(doc, index)
        if element and element[0] == "paragraph":
            paragraph = element[1]
            text = paragraph.getText()
            cursor = text.createTextCursorByRange(paragraph.getStart())
            cursor.collapseToStart()
            return cursor
    return None


# ---------------------------------------------------------------------------
# Helpers for Rejection (Deletion)
# ---------------------------------------------------------------------------


def _delete_graphic_by_name(doc, name: str) -> bool:
    """Delete a graphic object by its name."""
    try:
        graphics = doc.getGraphicObjects()
        if graphics.hasByName(name):
            graphic = graphics.getByName(name)
            graphic.dispose()
            return True
        return False
    except Exception as e:
        _log(f"_delete_graphic_by_name: failed: {e}")
        return False


def _delete_toc_by_name(doc, name: str) -> bool:
    """Delete a TOC by its name."""
    try:
        indexes = doc.getDocumentIndexes()
        count = indexes.getCount()
        for i in range(count):
            idx = indexes.getByIndex(i)
            # Check Name property if available (we set it on insert)
            if hasattr(idx, "Name") and idx.Name == name:
                idx.dispose()
                return True
        return False
    except Exception as e:
        _log(f"_delete_toc_by_name: failed: {e}")
        return False


def _delete_toc_by_index(doc, index: int) -> bool:
    try:
        indexes = doc.getDocumentIndexes()
        if 0 <= index < indexes.getCount():
            indexes.getByIndex(index).dispose()
            return True
        return False
    except Exception as e:
        _log(f"_delete_toc_by_index: failed: {e}")
        return False


def _delete_toc_from_state(doc, op_state: Dict[str, Any]) -> bool:
    toc_id = op_state.get("toc_id")
    if isinstance(toc_id, str) and toc_id and _delete_toc_by_name(doc, toc_id):
        return True

    raw_index = op_state.get("toc_index")
    if raw_index is None:
        raw_index = op_state.get("element_index")
    try:
        toc_index = int(raw_index if raw_index is not None else 0)
    except Exception:
        toc_index = 0
    return _delete_toc_by_index(doc, toc_index)


def _delete_note_by_id(doc, note_id: str, note_type: str) -> bool:
    """Delete a footnote or endnote by its ID (proxy ID or index)."""
    try:
        if note_type == "footnote":
            notes = doc.getFootnotes()
        else:
            notes = doc.getEndnotes()

        for i in range(notes.getCount()):
            note = notes.getByIndex(i)
            # Check ID match
            if str(id(note)) == note_id or str(i) == note_id:
                anchor = note.getAnchor()
                anchor.getText().removeTextContent(note)
                return True
        return False
    except Exception as e:
        _log(f"_delete_note_by_id: failed: {e}")
        return False


def _delete_shape_by_name(doc, name: str) -> bool:
    """Delete a shape (ControlShape) by its name."""
    try:
        draw_page = doc.getDrawPage()
        count = draw_page.getCount()
        for i in range(count):
            shape = draw_page.getByIndex(i)
            if hasattr(shape, "Name") and shape.Name == name:
                draw_page.remove(shape)
                return True
        return False
    except Exception as e:
        _log(f"_delete_shape_by_name: failed: {e}")
        return False


def _reset_marked_segments(doc, bookmark_names: List[str]) -> None:
    if not bookmark_names:
        return
    bookmarks = doc.getBookmarks()
    for name in bookmark_names:
        if not _validate_bookmark_exists(doc, name):
            _log(
                f"_reset_marked_segments: bookmark {name} not found or invalid; skipping."
            )
            continue
        try:
            bookmark = bookmarks.getByName(name)
            anchor = bookmark.getAnchor()
            text = anchor.getText()
            cursor = text.createTextCursorByRange(anchor)
            cursor.gotoRange(anchor.getEnd(), True)
            cursor.CharColor = 0
            cursor.CharStrikeout = 0
            cursor.CharUnderline = 0
            bookmark.dispose()
        except Exception as e:
            _log(f"_reset_marked_segments: failed for {name}: {e}")
            # Try to dispose bookmark even if reset failed
            try:
                if bookmarks.hasByName(name):
                    bookmarks.getByName(name).dispose()
            except Exception:
                pass


def _accept_comment_operation(
    doc, plan_id: str, op_id: str, op_state: Dict[str, Any]
) -> None:
    """Accept a comment operation, potentially reapplying properties."""
    op_type = op_state.get("type")

    if op_type == "comment_insertion":
        comment_id = op_state.get("comment_id")
        date_str = op_state.get("date")
        author = op_state.get("author")
        content_summary = op_state.get("text")

        if date_str:
            try:
                # Need to find the annotation to set the date
                text_fields = doc.getTextFields()
                enum = text_fields.createEnumeration()
                while enum.hasMoreElements():
                    field = enum.nextElement()
                    if field.supportsService("com.sun.star.text.textfield.Annotation"):
                        # Match by Author and Content summary (since Content in state might be truncated)
                        field_author = getattr(field, "Author", "")
                        field_content = getattr(field, "Content", "")

                        match = False
                        if field_author == author:
                            if content_summary.endswith("..."):
                                match = field_content.startswith(content_summary[:-3])
                            else:
                                match = field_content == content_summary

                        # Fallback to ID if match fails (same-session convenience)
                        if not match and str(id(field)) == comment_id:
                            match = True

                        if match:
                            # Parse date
                            parsed = dt.fromisoformat(date_str.replace("Z", "+00:00"))
                            uno_dt = UnoDateTime()
                            uno_dt.Year = parsed.year
                            uno_dt.Month = parsed.month
                            uno_dt.Day = parsed.day
                            uno_dt.Hours = parsed.hour
                            uno_dt.Minutes = parsed.minute
                            uno_dt.Seconds = parsed.second
                            uno_dt.NanoSeconds = 0
                            uno_dt.IsUTC = False

                            field.DateTimeValue = uno_dt
                            _log(
                                f"_accept_comment_operation: applied date to comment by {author}"
                            )
                            return

                _log(f"_accept_comment_operation: comment {comment_id} not found")
            except Exception as e:
                _log(f"_accept_comment_operation: failed to apply date: {e}")


g_exportedScripts = (resolvePreviewDiff,)
