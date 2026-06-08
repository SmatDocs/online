"""
PreviewActionPlan.py
--------------------

Generates inline, color-coded diff previews for WordActionPlan actions inside
the active LibreOffice Writer document. Each action's original content is
rendered in gold with strike-through while the proposed replacement is
inserted in green directly beneath it. Preview metadata is persisted so later
accept/reject operations can resolve the diff without losing formatting.
"""

import os
import json
import copy
import re
import ssl
import sys
import tempfile
import traceback
import uuid
import urllib.request
from urllib.parse import urlparse
from datetime import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from com.sun.star.awt.FontWeight import BOLD as FONT_WEIGHT_BOLD
from com.sun.star.awt.FontUnderline import SINGLE as FONT_UNDERLINE_SINGLE
from com.sun.star.awt.FontSlant import ITALIC as FONT_SLANT_ITALIC
from com.sun.star.style.ParagraphAdjust import (
    LEFT as PAR_ADJUST_LEFT,
    RIGHT as PAR_ADJUST_RIGHT,
    CENTER as PAR_ADJUST_CENTER,
    BLOCK as PAR_ADJUST_BLOCK,
)
from com.sun.star.text.ControlCharacter import LINE_BREAK, PARAGRAPH_BREAK
from com.sun.star.uno import Exception as UNOException
from com.sun.star.text.TextContentAnchorType import (
    AS_CHARACTER,
    AT_PARAGRAPH,
)
from com.sun.star.util import DateTime as UnoDateTime

PT_TO_HMM = 2540.0 / 72.0
GOLD_RGB = 0xD49F00
GREEN_RGB = 0x008A00

PREVIEW_STATE_KEY = "sdoc_preview_state"
LOG_FILE = "/tmp/sdoc_preview.log"
SDOC_BOOKMARK_RE = re.compile(r"^sdoc_[A-Za-z0-9]+$")
# Preview markers must never overlap with canonical sdoc_* IDs.
PREVIEW_BOOKMARK_PREFIX = "ov_preview"
DEBUG_LOG = os.environ.get("SDOC_PREVIEW_DEBUG", "").lower() in {"1", "true", "yes", "debug"}
SUMMARY_LOG_PREFIXES = (
    "PreviewActionPlan: raw payload",
    "PreviewActionPlan: plan_json length=",
    "PreviewActionPlan: plan parsed",
    "PreviewActionPlan: applying ",
    "PreviewActionPlan: state saved",
    "PreviewActionPlan: returning receipt=",
)


def _should_log(message: str) -> bool:
    if DEBUG_LOG:
        return True
    lowered = message.lower()
    if "failed" in lowered or "error" in lowered or "warning" in lowered:
        return True
    return message.startswith(SUMMARY_LOG_PREFIXES)


def _bootstrap_extra_python_paths() -> None:
    candidates = []

    env_paths = os.environ.get("SMARTDOCS_LO_PYTHONPATH", "")
    if env_paths:
        candidates.extend(
            part.strip() for part in env_paths.split(os.pathsep) if part.strip()
        )

    try:
        repo_root = Path(__file__).resolve().parents[5]
        default_vendor = repo_root / "backend" / "services" / "libreoffice" / "vendor"
        candidates.append(str(default_vendor))
    except Exception:
        pass

    for raw_path in candidates:
        try:
            path = Path(raw_path).expanduser().resolve()
            if path.exists():
                path_str = str(path)
                if path_str not in sys.path:
                    sys.path.insert(0, path_str)
        except Exception:
            continue


_bootstrap_extra_python_paths()


ctx = ssl._create_unverified_context()


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


def _preview_result(ok: bool, **fields: Any) -> str:
    payload = {
        "event": "word_preview",
        "ok": bool(ok),
        **fields,
    }
    return json.dumps(payload, ensure_ascii=False, default=str)


def _safe_get_start(text_range):
    """Return a stable start range for bookmark/paragraph/table comparisons."""
    try:
        return text_range.getStart()
    except Exception:
        return text_range


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


def _get_text_columns_property(columns, name: str, default=None):
    try:
        return getattr(columns, name)
    except Exception:
        return default


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


def _insert_reasoning_comment(
    doc,
    action: Dict[str, Any],
    reasoning: str,
) -> bool:
    """
    Insert a Writer annotation (comment) explaining AI reasoning.

    This allows users to understand why the AI made specific changes.
    The comment is anchored near the target element.

    Args:
        doc: The Writer document
        action: The action dict containing target info
        reasoning: The reasoning text to display

    Returns:
        True if comment was inserted successfully
    """
    try:
        # Get target location
        element_index = action.get("element_index")
        target_bookmark = action.get("target_bookmark")

        text = doc.getText()
        cursor = None

        # Try to locate the anchor by bookmark
        if target_bookmark:
            try:
                bookmarks = doc.getBookmarks()
                if bookmarks.hasByName(target_bookmark):
                    bookmark = bookmarks.getByName(target_bookmark)
                    cursor = text.createTextCursorByRange(bookmark.getAnchor())
            except Exception:
                pass

        # Fall back to element index
        if cursor is None and element_index is not None:
            try:
                enum = text.createEnumeration()
                idx = 0
                while enum.hasMoreElements():
                    element = enum.nextElement()
                    if idx == element_index:
                        cursor = text.createTextCursorByRange(element.getStart())
                        break
                    idx += 1
            except Exception:
                pass

        # Fall back to end of document
        if cursor is None:
            cursor = text.createTextCursor()
            cursor.gotoEnd(False)

        # Create and insert annotation
        annotation = doc.createInstance("com.sun.star.text.TextField.Annotation")
        annotation.Content = reasoning
        annotation.Author = "AI Assistant"

        text.insertTextContent(cursor, annotation, False)
        _log(f"_insert_reasoning_comment: inserted comment for action")
        return True

    except Exception as e:
        _log(f"_insert_reasoning_comment: failed - {e}")
        return False


def previewActionPlan(plan_json: str) -> str:
    """
    Entry point invoked from Collabora Online. Accepts a JSON encoded
    WordActionPlan and annotates the active document with inline previews.
    """
    try:
        _log("PreviewActionPlan: raw payload received.")
        _log(f"PreviewActionPlan: plan_json length={len(plan_json)}")
        _log(f"PreviewActionPlan: plan_json raw={plan_json}")
        plan = json.loads(plan_json)
        _log(f"PreviewActionPlan: plan parsed successfully. keys={list(plan.keys())}")
    except Exception as exc:
        message = f"failed to parse plan JSON: {exc}"
        _log(
            f"PreviewActionPlan: {message}\n{traceback.format_exc()}"
        )
        return _preview_result(False, error="parse_failed", message=message)

    try:
        document = XSCRIPTCONTEXT.getDocument()  # type: ignore  # NOQA
        _log(f"PreviewActionPlan: obtained document object={document}")
    except Exception as exc:
        message = f"XSCRIPTCONTEXT.getDocument failed: {exc}"
        _log(
            f"PreviewActionPlan: {message}\n{traceback.format_exc()}"
        )
        return _preview_result(False, plan_id=plan.get("plan_id"), error="document_failed", message=message)

    if document is None:
        message = "could not obtain active document (None returned)."
        _log(f"PreviewActionPlan: {message}")
        return _preview_result(False, plan_id=plan.get("plan_id"), error="document_missing", message=message)

    plan_id = plan.get("plan_id")
    if not plan_id:
        message = "plan_id missing."
        _log(f"PreviewActionPlan: {message}")
        return _preview_result(False, error="plan_id_missing", message=message)

    actions = plan.get("actions", [])
    if not isinstance(actions, list):
        message = "plan.actions is not a list."
        _log(f"PreviewActionPlan: {message}")
        return _preview_result(False, plan_id=plan_id, error="actions_invalid", message=message)

    try:
        _log("PreviewActionPlan: loading previous state…")
        state = _load_preview_state(document)
        plan_state = state.setdefault(plan_id, {})
        _log(
            f"PreviewActionPlan: state entries for plan before render: {len(plan_state)}"
        )
    except Exception as exc:
        message = f"failed to load state: {exc}"
        _log(
            f"PreviewActionPlan: {message}\n{traceback.format_exc()}"
        )
        return _preview_result(False, plan_id=plan_id, error="state_load_failed", message=message)

    doc = document
    controller = None
    try:
        controller = doc.getCurrentController()
        _log(f"PreviewActionPlan: obtained controller={controller}")
    except Exception as exc:
        message = f"getCurrentController failed: {exc}"
        _log(
            f"PreviewActionPlan: {message}\n{traceback.format_exc()}"
        )
        return _preview_result(False, plan_id=plan_id, error="controller_failed", message=message)

    if controller is None:
        message = "getCurrentController returned None."
        _log(f"PreviewActionPlan: {message}")
        return _preview_result(False, plan_id=plan_id, error="controller_missing", message=message)

    controllers_locked = False
    original_view_data = None
    results: List[Dict[str, Any]] = []
    rendered_operation_ids: List[str] = []
    failed_operation_ids: List[str] = []
    skipped_operation_ids: List[str] = []
    final_result: Optional[str] = None
    
    # Capture original tracking state robustly
    original_record_changes = getattr(doc, "RecordChanges", False)
    original_display_type = getattr(doc, "RedlineDisplayType", 0)

    try:
        try:
            original_view_data = controller.getViewData()
            _log("PreviewActionPlan: captured initial view data.")
        except Exception as view_exc:
            _log(
                f"PreviewActionPlan: getViewData failed (continuing without restore): {view_exc}"
            )
            original_view_data = None

        # Pause RecordChanges and ShowChanges robustly BEFORE locking controllers
        try:
            doc.RecordChanges = False
            _log(f"PreviewActionPlan: paused RecordChanges (was {original_record_changes})")
        except Exception as e:
            _log(f"PreviewActionPlan: failed to pause RecordChanges: {e}")

        try:
            # RedlineDisplayType: 0=None, 1=Inserted, 2=Removed, 3=All
            doc.RedlineDisplayType = 0 
            _log(f"PreviewActionPlan: set RedlineDisplayType to 0 (was {original_display_type})")
        except Exception as e:
            _log(f"PreviewActionPlan: failed to set RedlineDisplayType: {e}")

        try:
            _log("PreviewActionPlan: locking controllers…")
            doc.lockControllers()
            controllers_locked = True
            _log("PreviewActionPlan: controllers locked.")
        except Exception as exc:
            message = f"failed to lock controllers: {exc}"
            _log(
                f"PreviewActionPlan: {message}\n{traceback.format_exc()}"
            )
            final_result = _preview_result(False, plan_id=plan_id, error="controller_lock_failed", message=message)
            return final_result

        try:
            # Ensure anchor bookmarks expected by the plan exist before preview rendering.
            # This handles sessions where the open live doc missed backend-injected sdoc bookmarks.
            _ensure_required_sdoc_bookmarks(doc, actions)

            try:
                bookmarks_count = document.getBookmarks().getCount()
            except Exception:
                bookmarks_count = "unknown"
            _log(
                f"PreviewActionPlan: applying {len(actions)} actions for plan {plan_id} …"
            )
            _log(f"PreviewActionPlan: existing bookmarks count={bookmarks_count}")

            # Step 3b: Auto-chain consecutive inserts that share the same after_element_id
            _auto_chain_insert_actions(actions)

            # Step 4: Apply actions
            try:
                for action in actions:
                    op_id = action.get("operation_id") or action.get(
                        "metadata", {}
                    ).get("title")
                    if not op_id:
                        _log("PreviewActionPlan: skipping action without operation_id.")
                        skipped_operation_ids.append("")
                        results.append(
                            {
                                "operation_id": None,
                                "kind": action.get("kind"),
                                "status": "skipped",
                                "reason": "operation_id_missing",
                            }
                        )
                        continue

                    if plan_state.get(op_id, {}).get("status") == "resolved":
                        _log(
                            f"PreviewActionPlan: skipping already resolved action {op_id}"
                        )
                        skipped_operation_ids.append(op_id)
                        results.append(
                            {
                                "operation_id": op_id,
                                "kind": action.get("kind"),
                                "status": "skipped",
                                "reason": "already_resolved",
                            }
                        )
                        continue

                    _log(
                        f"PreviewActionPlan: rendering action {op_id} kind={action.get('kind')}"
                    )
                    try:
                        preview_info = _render_action_preview(
                            doc, action, plan_id, op_id
                        )
                        if preview_info:
                            plan_state[op_id] = preview_info
                            rendered_operation_ids.append(op_id)
                            results.append(
                                {
                                    "operation_id": op_id,
                                    "kind": action.get("kind"),
                                    "status": "ok",
                                    "preview_type": preview_info.get("type"),
                                    "target_bookmark": preview_info.get("target_bookmark"),
                                    "inserted_count": len(preview_info.get("inserted_ids") or []),
                                    "marked_count": len(preview_info.get("marked_ids") or []),
                                    "separator_count": len(preview_info.get("separator_ids") or []),
                                }
                            )
                            _log(
                                f"PreviewActionPlan: action {op_id} rendered. preview keys={list(preview_info.keys())}"
                            )

                            # Insert reasoning comment if provided in action metadata
                            reasoning = (action.get("metadata") or {}).get("reasoning")
                            if reasoning:
                                # We pass the action and reasoning to a helper
                                _insert_reasoning_comment(doc, action, reasoning)
                        else:
                            failed_operation_ids.append(op_id)
                            results.append(
                                {
                                    "operation_id": op_id,
                                    "kind": action.get("kind"),
                                    "status": "error",
                                    "error": "preview_not_rendered",
                                    "message": "Preview renderer returned no preview metadata.",
                                }
                            )

                    except Exception as action_exc:
                        failed_operation_ids.append(op_id)
                        results.append(
                            {
                                "operation_id": op_id,
                                "kind": action.get("kind"),
                                "status": "error",
                                "error": "render_failed",
                                "message": str(action_exc),
                            }
                        )
                        _log(
                            f"PreviewActionPlan: failed to render preview for {op_id}: {action_exc}\n{traceback.format_exc()}"
                        )

                _save_preview_state(doc, state)
                _log(f"PreviewActionPlan: state saved for plan {plan_id}")
                final_result = _preview_result(
                    len(failed_operation_ids) == 0,
                    plan_id=plan_id,
                    action_count=len(actions),
                    rendered_ops=len(rendered_operation_ids),
                    failed_ops=len(failed_operation_ids),
                    skipped_ops=len(skipped_operation_ids),
                    rendered_operation_ids=rendered_operation_ids,
                    failed_operation_ids=failed_operation_ids,
                    skipped_operation_ids=skipped_operation_ids,
                    results=results,
                    all_done=len(failed_operation_ids) == 0,
                )

            except Exception as e:
                message = f"render loop failed: {e}"
                _log(
                    f"PreviewActionPlan: {message}\n{traceback.format_exc()}"
                )
                final_result = _preview_result(
                    False,
                    plan_id=plan_id,
                    error="render_loop_failed",
                    message=message,
                    action_count=len(actions),
                    rendered_ops=len(rendered_operation_ids),
                    failed_ops=max(1, len(failed_operation_ids)),
                    skipped_ops=len(skipped_operation_ids),
                    rendered_operation_ids=rendered_operation_ids,
                    failed_operation_ids=failed_operation_ids,
                    skipped_operation_ids=skipped_operation_ids,
                    results=results,
                    all_done=False,
                )
        finally:
            # This inner finally block is for any cleanup specific to the rendering loop
            # The outer finally block handles controller unlocking and view restoration.
            pass  # No specific cleanup needed here based on the provided edit, as RecordChanges is handled outside.

    finally:
        if controllers_locked:
            try:
                # Restore RecordChanges robustly
                try:
                    doc.RecordChanges = original_record_changes
                    _log(f"PreviewActionPlan: restored RecordChanges to {original_record_changes}")
                except Exception as e:
                    _log(f"PreviewActionPlan: failed to restore RecordChanges: {e}")

                # Keep redlines hidden after preview. Visibility should only be
                # changed via explicit user toggle (SetShowChanges).
                try:
                    doc.RedlineDisplayType = 0
                    _log(
                        "PreviewActionPlan: enforced RedlineDisplayType=0 after preview"
                    )
                except Exception as e:
                    _log(
                        f"PreviewActionPlan: failed to enforce RedlineDisplayType=0: {e}"
                    )

                doc.unlockControllers()
                _log("PreviewActionPlan: controllers unlocked.")
            except Exception as unlock_exc:
                _log(
                    f"PreviewActionPlan: unlockControllers failed: {unlock_exc}\n{traceback.format_exc()}"
                )

        if controller is not None and original_view_data is not None:
            try:
                controller.restoreViewData(original_view_data)
                _log("PreviewActionPlan: controller restored.")
            except Exception as restore_exc:
                _log(
                    f"PreviewActionPlan: controller restore failed: {restore_exc}\n{traceback.format_exc()}"
                )

    if final_result is None:
        final_result = _preview_result(
            False,
            plan_id=plan_id,
            error="preview_no_result",
            message="Preview script finished without a render receipt.",
            action_count=len(actions),
            rendered_ops=len(rendered_operation_ids),
            failed_ops=max(1, len(failed_operation_ids)),
            skipped_ops=len(skipped_operation_ids),
            rendered_operation_ids=rendered_operation_ids,
            failed_operation_ids=failed_operation_ids,
            skipped_operation_ids=skipped_operation_ids,
            results=results,
            all_done=False,
        )
    _log(f"PreviewActionPlan: returning receipt={final_result}")
    return final_result


def clearPreviewState(doc) -> None:
    """Utility hook to remove stored preview metadata."""
    _save_preview_state(doc, {})


# ---------------------------------------------------------------------------
# Preview rendering helpers
# ---------------------------------------------------------------------------


def _render_action_preview(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    kind = action.get("kind")

    if kind == "replace_paragraph":
        return _preview_replace_paragraph(doc, action, plan_id, op_id)
    if kind == "insert_paragraph":
        return _preview_insert_paragraph(doc, action, plan_id, op_id)
    if kind == "delete_element":
        return _preview_delete_element(doc, action, plan_id, op_id)
    if kind == "replace_table":
        return _preview_replace_table(doc, action, plan_id, op_id)
    if kind == "insert_table":
        return _preview_insert_table(doc, action, plan_id, op_id)
    if kind == "set_paragraph_style":
        return _preview_set_paragraph_style(doc, action, plan_id, op_id)
    if kind == "set_document_style":
        return _preview_document_style(doc, action, plan_id, op_id)
    if kind in ("insert_run", "replace_run"):
        return _preview_run_edit(doc, action, plan_id, op_id)
    if kind == "insert_chart":
        return _preview_insert_chart(doc, action, plan_id, op_id)
    if kind == "replace_chart":
        return _preview_replace_chart(doc, action, plan_id, op_id)
    if kind == "update_chart_data":
        return _preview_update_chart_data(doc, action, plan_id, op_id)
    # Header/footer actions
    if kind == "set_header":
        return _preview_set_header(doc, action, plan_id, op_id)
    if kind == "set_footer":
        return _preview_set_footer(doc, action, plan_id, op_id)
    if kind == "clear_header":
        return _preview_clear_header(doc, action, plan_id, op_id)
    if kind == "clear_footer":
        return _preview_clear_footer(doc, action, plan_id, op_id)
    # Field actions
    if kind == "insert_field":
        return _preview_insert_field(doc, action, plan_id, op_id)
    # Table cell merge/split actions
    if kind == "merge_cells":
        return _preview_merge_cells(doc, action, plan_id, op_id)
    if kind == "split_cell":
        return _preview_split_cell(doc, action, plan_id, op_id)
    # Section break actions
    if kind == "insert_section_break":
        return _preview_insert_section_break(doc, action, plan_id, op_id)
    # Track changes actions (Phase 2)
    if kind == "enable_track_changes":
        return _preview_enable_track_changes(doc, action, plan_id, op_id)
    if kind == "accept_revision":
        return _preview_accept_revision(doc, action, plan_id, op_id)
    if kind == "reject_revision":
        return _preview_reject_revision(doc, action, plan_id, op_id)
    if kind == "accept_all_revisions":
        return _preview_accept_all_revisions(doc, action, plan_id, op_id)
    if kind == "reject_all_revisions":
        return _preview_reject_all_revisions(doc, action, plan_id, op_id)
    # Comment actions (Phase 2)
    if kind == "insert_comment":
        return _preview_insert_comment(doc, action, plan_id, op_id)
    if kind == "reply_to_comment":
        return _preview_reply_to_comment(doc, action, plan_id, op_id)
    if kind == "resolve_comment":
        return _preview_resolve_comment(doc, action, plan_id, op_id)
    if kind == "delete_comment":
        return _preview_delete_comment(doc, action, plan_id, op_id)
    # Style management actions (Phase 2)
    if kind == "create_style":
        return _preview_create_style(doc, action, plan_id, op_id)
    if kind == "modify_style":
        return _preview_modify_style(doc, action, plan_id, op_id)
    if kind == "delete_style":
        return _preview_delete_style(doc, action, plan_id, op_id)
    # Footnote/Endnote actions (Phase 2)
    if kind == "insert_footnote":
        return _preview_insert_footnote(doc, action, plan_id, op_id)
    if kind == "insert_endnote":
        return _preview_insert_endnote(doc, action, plan_id, op_id)
    if kind == "modify_note":
        return _preview_modify_note(doc, action, plan_id, op_id)
    if kind == "delete_note":
        return _preview_delete_note(doc, action, plan_id, op_id)
    # Image actions (Phase 2)
    if kind == "insert_image":
        return _preview_insert_image(doc, action, plan_id, op_id)
    if kind == "replace_image":
        return _preview_replace_image(doc, action, plan_id, op_id)
    if kind == "delete_image":
        return _preview_delete_image(doc, action, plan_id, op_id)
    # Table of Contents actions (Phase 3)
    if kind == "insert_toc":
        return _preview_insert_toc(doc, action, plan_id, op_id)
    if kind == "update_toc":
        return _preview_update_toc(doc, action, plan_id, op_id)
    if kind == "remove_toc":
        return _preview_remove_toc(doc, action, plan_id, op_id)
    # Mail Merge actions (Phase 3)
    if kind == "insert_merge_field":
        return _preview_insert_merge_field(doc, action, plan_id, op_id)
    if kind == "apply_merge_data":
        return _preview_apply_merge_data(doc, action, plan_id, op_id)
    # Content Controls actions (Phase 3)
    if kind == "insert_control":
        return _preview_insert_control(doc, action, plan_id, op_id)
    # Cross-Reference actions (Phase 3)
    if kind == "insert_cross_reference":
        return _preview_insert_cross_reference(doc, action, plan_id, op_id)
    # Master Document actions (Phase 3)
    if kind == "insert_sub_document":
        return _preview_insert_sub_document(doc, action, plan_id, op_id)
    # Multi-Column Layout actions (Phase 4)
    if kind == "set_columns":
        return _preview_set_columns(doc, action, plan_id, op_id)
    if kind == "insert_column_break":
        return _preview_insert_column_break(doc, action, plan_id, op_id)
    # Text Box actions (Phase 4)
    if kind == "insert_text_box":
        return _preview_insert_text_box(doc, action, plan_id, op_id)
    if kind == "delete_text_box":
        return _preview_delete_text_box(doc, action, plan_id, op_id)
    # Drop Cap actions (Phase 4)
    if kind == "set_drop_cap":
        return _preview_set_drop_cap(doc, action, plan_id, op_id)
    # Shape actions (Phase 4)
    if kind == "insert_shape":
        return _preview_insert_shape(doc, action, plan_id, op_id)
    # Linked Frames actions (Phase 4)
    if kind == "link_text_frames":
        return _preview_link_text_frames(doc, action, plan_id, op_id)
    # Phase 5: Lists, Watermarks, Equations, Comparison
    if kind == "set_list_style":
        return _preview_set_list_style(doc, action, plan_id, op_id)
    if kind == "insert_watermark":
        return _preview_insert_watermark(doc, action, plan_id, op_id)
    if kind == "insert_equation":
        return _preview_insert_equation(doc, action, plan_id, op_id)
    if kind == "compare_documents":
        return _preview_compare_documents(doc, action, plan_id, op_id)

    _log(f"PreviewActionPlan: unsupported action kind '{kind}'")
    return None


def _preview_replace_paragraph(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_replace_paragraph: locating paragraph for op {op_id}")
    paragraph = _locate_paragraph(doc, action)
    if paragraph is None:
        raise RuntimeError("target paragraph not found")
    _log(f"_preview_replace_paragraph: paragraph located for op {op_id}")

    original_spec = _extract_paragraph_spec(paragraph)

    # Style original as Gold/Strikethrough
    _apply_char_style(paragraph, color=GOLD_RGB, strike=True)
    original_mark = _tag_preview_range(doc, paragraph, plan_id, op_id, role="original")

    # ONE empty line gap between Gold and Green
    text = paragraph.getText()
    cursor_after = text.createTextCursorByRange(paragraph.getEnd())
    spacing_mark = _insert_separator(
        doc, cursor_after, plan_id, op_id, role="separator_gap"
    )

    proposed_spec = action.get("paragraph") or {}
    inserted_paragraphs = _insert_paragraphs_from_spec(
        doc,
        cursor_after,
        proposed_spec,
        plan_id,
        op_id,
        add_separator=False,
        # Commented behavior: avoid appending an extra trailing paragraph break
        # for replace-paragraph previews; keep spacing before incoming change,
        # but not after it.
        append_break=False,
    )

    return {
        "type": "paragraph_replacement",
        "original_spec": original_spec,
        "proposed_spec": proposed_spec,
        "inserted_ids": inserted_paragraphs,
        "separator_ids": [spacing_mark],
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
        "marked_ids": [original_mark],
    }


def _preview_insert_paragraph(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_insert_paragraph: locating anchor for op {op_id}")
    pos_action = {k: v for k, v in action.items() if k != "element_id"}
    if action.get("after_element_id"):
        anchor_cursor = _create_cursor_after_action_target(doc, pos_action)
    elif action.get("before_element_id"):
        anchor_cursor = _create_cursor_before_action_target(doc, pos_action)
    elif action.get("target_element_id"):
        _log(f"_preview_insert_paragraph: target_element_id used as after-anchor fallback")
        anchor_cursor = _create_cursor_after_action_target(doc, pos_action)
    else:
        anchor_cursor = _create_cursor_before_action_target(doc, pos_action)
    if anchor_cursor is None:
        raise RuntimeError("could not resolve insertion anchor")
    _log(f"_preview_insert_paragraph: anchor located for op {op_id}")

    spec = action.get("paragraph") or {}

    anchor_cursor = _ensure_paragraph_boundary_for_insertion(anchor_cursor)

    inserted = _write_paragraph_spec(
        anchor_cursor,
        spec,
        color=GREEN_RGB,
        preserve_explicit_run_colors=True,
    )
    explicit_element_id = spec.get("element_id")
    if explicit_element_id and inserted:
        _tag_named_range(doc, inserted[0], explicit_element_id)
    root_element_id = action.get("element_id")
    if root_element_id and root_element_id != explicit_element_id and inserted:
        _tag_named_range(doc, inserted[0], root_element_id)
    inserted_ids = [
        _tag_preview_range(doc, p, plan_id, op_id, role=f"proposed_{idx}")
        for idx, p in enumerate(inserted)
    ]

    return {
        "type": "paragraph_insertion",
        "proposed_spec": spec,
        "inserted_ids": inserted_ids,
        "separator_ids": [],
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _preview_delete_element(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_delete_element: locating element for op {op_id}")
    element_info = _locate_element(doc, action)
    if element_info is None:
        raise RuntimeError("target element not found for delete preview")
    _log(
        f"_preview_delete_element: element located kind={element_info[0]} for op {op_id}"
    )

    kind, target = element_info
    if kind == "paragraph":
        original_spec = _extract_paragraph_spec(target)
        _apply_char_style(target, color=GOLD_RGB, strike=True)
        mark = _tag_preview_range(doc, target, plan_id, op_id, role="original")
        return {
            "type": "delete_paragraph",
            "original_spec": original_spec,
            "marked_ids": [mark],
            "target_bookmark": _resolve_target_bookmark(action),
            "target_index": action.get("element_index"),
        }

    if kind == "table":
        original_spec = _extract_table_spec(target)
        _colorize_table(target, GOLD_RGB, strike=True)
        mark = _tag_table(doc, target, plan_id, op_id, role="original")
        return {
            "type": "delete_table",
            "original_spec": original_spec,
            "marked_ids": [mark] if mark else [],
            "original_table_name": _safe_table_name(target),
            "target_bookmark": _resolve_target_bookmark(action),
            "target_index": action.get("element_index"),
        }

    raise RuntimeError(f"delete preview unsupported for element type '{kind}'")


def _preview_replace_table(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_replace_table: locating element for op {op_id}")
    element_info = _locate_element(doc, action)
    if element_info is None or element_info[0] != "table":
        raise RuntimeError("target table not found for replacement preview")
    _log(f"_preview_replace_table: table located for op {op_id}")

    target_table = element_info[1]
    original_spec = _extract_table_spec(target_table)
    original_table_name = _safe_table_name(target_table)

    # Style original as Gold/Strikethrough
    _colorize_table(target_table, GOLD_RGB, strike=True)
    original_mark = _tag_table(doc, target_table, plan_id, op_id, role="original")

    # Spacing logic: GAP between gold and green versions
    cursor_after = _create_cursor_after_table(doc, target_table)
    spacing_mark = _insert_separator(
        doc, cursor_after, plan_id, op_id, role="separator_gap"
    )

    table_spec = _normalize_table_spec(action.get("table") or {})
    inserted_table = _insert_table_from_spec(
        doc, cursor_after, table_spec, color=GREEN_RGB
    )
    proposed_table_name = _safe_table_name(inserted_table)
    inserted_mark = _tag_table(doc, inserted_table, plan_id, op_id, role="proposed")

    return {
        "type": "table_replacement",
        "original_spec": original_spec,
        "proposed_spec": table_spec,
        "inserted_ids": [inserted_mark] if inserted_mark else [],
        "separator_ids": [spacing_mark],
        "marked_ids": [original_mark] if original_mark else [],
        "original_table_name": original_table_name,
        "proposed_table_name": proposed_table_name,
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _preview_insert_table(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_insert_table: locating anchor for op {op_id}")
    pos_action = {k: v for k, v in action.items() if k != "element_id"}
    if action.get("after_element_id"):
        anchor_cursor = _create_cursor_after_action_target(doc, pos_action)
    elif action.get("before_element_id"):
        anchor_cursor = _create_cursor_before_action_target(doc, pos_action)
    elif action.get("target_element_id"):
        _log(f"_preview_insert_table: target_element_id used as after-anchor fallback")
        anchor_cursor = _create_cursor_after_action_target(doc, pos_action)
    else:
        anchor_cursor = _create_cursor_before_action_target(doc, pos_action)
    if anchor_cursor is None:
        raise RuntimeError("could not resolve insert anchor for table")
    _log(f"_preview_insert_table: anchor located for op {op_id}")

    table_spec = _normalize_table_spec(action.get("table") or {})
    text = anchor_cursor.getText()

    # Eliminated the gap insertion here for natural flow.
    table = _insert_table_from_spec(doc, anchor_cursor, table_spec, color=GREEN_RGB)
    proposed_table_name = _safe_table_name(table)

    # Create bookmark right after the table using the table's end position
    try:
        # Try to get cursor after table using anchor end
        anchor = table.getAnchor()
        # CRITICAL: Always use the text object from the anchor itself
        text_of_anchor = anchor.getText()
        if text_of_anchor is not None:
            # Create a cursor and move it to the end of the table anchor
            bookmark_cursor = text_of_anchor.createTextCursor()
            bookmark_cursor.gotoRange(anchor.getEnd(), False)
        else:
            bookmark_cursor = text.createTextCursor()
            bookmark_cursor.gotoRange(anchor.getEnd(), False)
        bookmark_cursor.collapseToStart()  # Position cursor right after table
    except Exception as e:
        _log(
            f"_preview_insert_table: failed to create cursor from table anchor end: {e}"
        )
        # Fallback: use the insertion cursor position
        try:
            bookmark_cursor = text.createTextCursorByRange(anchor_cursor)
            bookmark_cursor.gotoRange(table.getAnchor().getEnd(), False)
        except Exception:
            bookmark_cursor = anchor_cursor.getStart()
        try:
            # Move cursor to end of document or next position
            bookmark_cursor.gotoRange(anchor.getEnd(), False)
        except Exception:
            pass

    bookmark_name = (
        f"{PREVIEW_BOOKMARK_PREFIX}_{plan_id}_{op_id}_proposed".replace("-", "_")
    )
    _create_bookmark(doc, text, bookmark_cursor, bookmark_name)
    boundary_mark = _insert_table_trailing_boundary(
        doc, table, plan_id, op_id, role="boundary_after"
    )

    table_element_id = table_spec.get("element_id")
    if table_element_id:
        try:
            anchor = table.getAnchor()
            tag_cursor = anchor.getText().createTextCursorByRange(anchor.getStart())
            _tag_named_range(doc, tag_cursor, table_element_id)
        except Exception as e:
            _log(f"_preview_insert_table: failed to tag table element_id {table_element_id}: {e}")
    root_element_id = action.get("element_id")
    if root_element_id and root_element_id != table_element_id:
        try:
            anchor = table.getAnchor()
            tag_cursor = anchor.getText().createTextCursorByRange(anchor.getStart())
            _tag_named_range(doc, tag_cursor, root_element_id)
        except Exception as e:
            _log(f"_preview_insert_table: failed to tag root element_id {root_element_id}: {e}")

    return {
        "type": "table_insertion",
        "proposed_spec": table_spec,
        "inserted_ids": [bookmark_name],
        "boundary_ids": [boundary_mark] if boundary_mark else [],
        "separator_ids": [],
        "proposed_table_name": proposed_table_name,
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _preview_set_paragraph_style(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_set_paragraph_style: locating paragraph for op {op_id}")
    paragraph = _locate_paragraph(doc, action)
    if paragraph is None:
        raise RuntimeError("target paragraph not found for style preview")
    _log(f"_preview_set_paragraph_style: paragraph located for op {op_id}")

    original_spec = _extract_paragraph_spec(paragraph)
    _apply_style_preview(paragraph, action.get("update") or {})
    mark = _tag_preview_range(doc, paragraph, plan_id, op_id, role="original")

    return {
        "type": "paragraph_style",
        "original_spec": original_spec,
        "marked_ids": [mark],
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _preview_document_style(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    # No inline diff; record metadata for later accept/reject.
    return {
        "type": "document_style",
        "update": action.get("update") or {},
        "target_index": action.get("element_index"),
    }


def _preview_run_edit(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    _log(f"_preview_run_edit: locating paragraph for op {op_id}")
    paragraph = _locate_paragraph(doc, action)
    if paragraph is None:
        raise RuntimeError("parent paragraph not found for run edit")
    _log(f"_preview_run_edit: paragraph located for op {op_id}")
    original_spec = _extract_paragraph_spec(paragraph)
    proposed_spec = copy.deepcopy(original_spec)
    _apply_run_edit_to_spec(proposed_spec, action)

    _apply_char_style(paragraph, color=GOLD_RGB, strike=True)
    mark = _tag_preview_range(doc, paragraph, plan_id, op_id, role="original")

    # Insert a visible paragraph break between gold and green preview,
    # and track it separately so it can be removed on resolve.
    text = paragraph.getText()
    cursor_after = text.createTextCursorByRange(paragraph.getEnd())
    spacing_mark = _insert_separator(
        doc, cursor_after, plan_id, op_id, role="separator_gap"
    )
    inserted_paragraphs = _insert_paragraphs_from_spec(
        doc,
        cursor_after,
        proposed_spec,
        plan_id,
        op_id,
        add_separator=False,
        # Match replace_paragraph behavior: keep separator before incoming
        # content, but avoid adding an extra trailing break after it.
        append_break=False,
    )

    return {
        "type": "run_replacement",
        "original_spec": original_spec,
        "proposed_spec": proposed_spec,
        "inserted_ids": inserted_paragraphs,
        "separator_ids": [spacing_mark],
        "marked_ids": [mark],
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


# ---------------------------------------------------------------------------
# Chart preview functions
# ---------------------------------------------------------------------------


def _preview_insert_chart(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview inserting a chart at the specified location."""
    _log(f"_preview_insert_chart: locating anchor for op {op_id}")

    pos_action = {k: v for k, v in action.items() if k != "element_id"}
    if action.get("after_element_id"):
        anchor_cursor = _create_cursor_after_action_target(doc, pos_action)
    elif action.get("before_element_id"):
        anchor_cursor = _create_cursor_before_action_target(doc, pos_action)
    elif action.get("target_element_id"):
        _log(f"_preview_insert_chart: target_element_id used as after-anchor fallback")
        anchor_cursor = _create_cursor_after_action_target(doc, pos_action)
    else:
        anchor_cursor = _create_cursor_before_action_target(doc, pos_action)

    if anchor_cursor is None:
        raise RuntimeError("could not resolve insert anchor for chart")

    _log(f"_preview_insert_chart: anchor located for op {op_id}")

    chart_spec = action.get("chart") or {}
    chart_obj = _insert_chart_from_spec(doc, anchor_cursor, chart_spec, color=GREEN_RGB)

    if chart_obj is None:
        _log(f"_preview_insert_chart: failed to create chart for op {op_id}")
        return None

    # Tag the chart for tracking
    chart_mark = f"sdoc_chart_{plan_id}_{op_id}_proposed"

    return {
        "type": "chart_insertion",
        "inserted_ids": [chart_mark],
        "separator_ids": [],
        "chart_spec": chart_spec,
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _preview_replace_chart(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview replacing an existing chart."""
    _log(f"_preview_replace_chart: locating chart for op {op_id}")

    # For now, implement as insert_chart (full replace would need chart lookup)
    # This is a simplified implementation
    element = _locate_element(doc, action)
    if element is None:
        raise RuntimeError("target element not found for chart replacement")

    chart_spec = action.get("chart") or {}

    # Create cursor after the existing element
    text = doc.getText()
    cursor = text.createTextCursorByRange(element.getAnchor().getEnd())

    # Add smart spacing before the new chart and track for cleanup
    spacing_mark = _insert_separator(doc, cursor, plan_id, op_id, role="separator_gap")

    chart_obj = _insert_chart_from_spec(doc, cursor, chart_spec, color=GREEN_RGB)

    if chart_obj is None:
        _log(f"_preview_replace_chart: failed to create chart for op {op_id}")
        return None

    chart_mark = f"sdoc_chart_{plan_id}_{op_id}_proposed"

    return {
        "type": "chart_replacement",
        "chart_spec": chart_spec,
        "inserted_ids": [chart_mark],
        "separator_ids": [spacing_mark],
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _preview_update_chart_data(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview updating chart data (inline update, no visual diff needed)."""
    _log(f"_preview_update_chart_data: recording operation for op {op_id}")

    # Data update doesn't require visual preview - just record the intent
    return {
        "type": "chart_data_update",
        "data_spec": action.get("data") or {},
        "target_bookmark": _resolve_target_bookmark(action),
        "target_index": action.get("element_index"),
    }


def _insert_chart_from_spec(
    doc, cursor, spec: Dict[str, Any], color: Optional[int] = None
):
    """
    Insert a chart into a Writer document at the cursor position.
    Uses LibreOffice UNO API to create an embedded chart object.
    """
    _log(
        f"_insert_chart_from_spec: creating chart with spec: {spec.get('chart_type', 'bar')}"
    )

    try:
        # Get chart dimensions
        width_in = spec.get("width_in", 5.0)
        height_in = spec.get("height_in", 3.0)
        width_hmm = int(width_in * 2540)  # inches to 1/100 mm
        height_hmm = int(height_in * 2540)

        # Create the embedded chart object
        chart_obj = doc.createInstance("com.sun.star.text.TextEmbeddedObject")

        # Set the CLSID for charts in LibreOffice
        # This is the class ID for embedded chart objects
        chart_obj.CLSID = "12dcae26-281f-416f-a234-c3086127382e"

        # Set dimensions
        chart_obj.Width = width_hmm
        chart_obj.Height = height_hmm

        # Set anchor type based on spec
        anchor_type = spec.get("anchor_type", "paragraph")
        if anchor_type == "inline":
            chart_obj.AnchorType = AS_CHARACTER
        else:
            chart_obj.AnchorType = AT_PARAGRAPH

        # Insert the chart into the document
        text = cursor.getText()
        text.insertTextContent(cursor, chart_obj, False)

        # Get the embedded chart document and configure it
        chart_document = chart_obj.getEmbeddedObject()
        if chart_document is not None:
            if not _configure_chart(chart_document, spec):
                raise RuntimeError("chart document could not be configured with data")
        else:
            raise RuntimeError("embedded chart document was not available")

        _log(f"_insert_chart_from_spec: chart created successfully")
        return chart_obj

    except Exception as e:
        _log(f"_insert_chart_from_spec: failed to create chart: {e}")
        return None


def _configure_chart(chart_document, spec: Dict[str, Any]) -> bool:
    """Configure chart type, data, and styling."""
    _log(f"_configure_chart: applying chart configuration")

    try:
        _initialize_chart_document(chart_document)

        # Set chart data
        data_spec = spec.get("data") or {}
        data_ok = _set_chart_data(chart_document, data_spec)
        if not data_ok:
            _log("_configure_chart: chart data was not applied")

        # Some Writer chart data APIs recreate the default column chart while
        # binding data, so apply the requested type after the data is in place.
        chart_type = _normalize_chart_type(spec.get("chart_type", "bar"))
        type_ok = _apply_chart_type(chart_document, chart_type)
        if chart_type in {"pie", "pie3d", "doughnut"} and not type_ok:
            _log(f"_configure_chart: requested chart type '{chart_type}' was not applied")
            return False

        # Set title
        title = spec.get("title")
        if title:
            _set_chart_title(chart_document, title)

        # Configure legend
        show_legend = spec.get("show_legend", True)
        _configure_chart_legend(chart_document, show_legend)
        return data_ok

    except Exception as e:
        _log(f"_configure_chart: error during configuration: {e}")
        return False


def _initialize_chart_document(chart_document) -> bool:
    """Ensure embedded Writer charts have an internal data provider."""
    if chart_document is None:
        return False

    changed = False
    for method_name, args in (
        ("createDefaultChart", ()),
        ("createInternalDataProvider", (False,)),
    ):
        method = getattr(chart_document, method_name, None)
        if not callable(method):
            continue
        try:
            if method_name == "createInternalDataProvider":
                has_internal = getattr(chart_document, "hasInternalDataProvider", None)
                if callable(has_internal) and has_internal():
                    continue
            method(*args)
            changed = True
            _log(f"_initialize_chart_document: called {method_name}")
        except Exception as exc:
            _log(f"_initialize_chart_document: {method_name} failed: {exc}")
    return changed


_CHART_TYPE_ALIASES = {
    "donut": "doughnut",
}


def _normalize_chart_type(chart_type: str) -> str:
    value = str(chart_type or "").strip().lower().replace("-", "_")
    return _CHART_TYPE_ALIASES.get(value, value)


def _chart_type_to_uno(chart_type: str) -> str:
    chart_type = _normalize_chart_type(chart_type)
    mapping = {
        "bar": "com.sun.star.chart2.BarDiagram",
        "column": "com.sun.star.chart2.ColumnDiagram",
        "line": "com.sun.star.chart2.LineDiagram",
        "pie": "com.sun.star.chart2.PieDiagram",
        "pie3d": "com.sun.star.chart2.PieDiagram",
        "doughnut": "com.sun.star.chart2.DoughnutDiagram",
        "area": "com.sun.star.chart2.AreaDiagram",
        "scatter": "com.sun.star.chart2.XYDiagram",
        "radar": "com.sun.star.chart2.NetDiagram",
    }
    return mapping.get(chart_type, mapping["bar"])


def _chart_type_to_classic(chart_type: str) -> Optional[str]:
    chart_type = _normalize_chart_type(chart_type)
    mapping = {
        "bar": "com.sun.star.chart.BarDiagram",
        "column": "com.sun.star.chart.ColumnDiagram",
        "line": "com.sun.star.chart.LineDiagram",
        "pie": "com.sun.star.chart.PieDiagram",
        "pie3d": "com.sun.star.chart.PieDiagram",
        "doughnut": "com.sun.star.chart.DonutDiagram",
        "area": "com.sun.star.chart.AreaDiagram",
        "scatter": "com.sun.star.chart.XYDiagram",
        "radar": "com.sun.star.chart.NetDiagram",
    }
    return mapping.get(chart_type)


def _set_chart_3d_hint(diagram, chart_type: str) -> None:
    if diagram is None:
        return
    try:
        if hasattr(diagram, "Dim3D"):
            diagram.Dim3D = chart_type in {"pie3d"}
    except Exception:
        pass


def _chart_type_markers(chart_type: str) -> Tuple[str, ...]:
    chart_type = _normalize_chart_type(chart_type)
    if chart_type in {"pie", "pie3d"}:
        return ("pie",)
    if chart_type == "doughnut":
        return ("donut", "doughnut")
    if chart_type == "line":
        return ("line",)
    if chart_type == "area":
        return ("area",)
    if chart_type == "scatter":
        return ("xy", "scatter")
    if chart_type == "radar":
        return ("net", "radar")
    if chart_type == "column":
        return ("column", "bar")
    return ("bar", "column")


def _chart_type_matches(chart_document, chart_type: str) -> bool:
    try:
        diagram = chart_document.getDiagram()
    except Exception:
        diagram = None
    if diagram is None:
        return False

    markers = _chart_type_markers(chart_type)
    names = []
    for attr in ("DiagramType", "getDiagramType", "ImplementationName"):
        try:
            value = getattr(diagram, attr, None)
            if callable(value):
                value = value()
            if value:
                names.append(str(value))
        except Exception:
            continue
    try:
        service_names = getattr(diagram, "getSupportedServiceNames", None)
        if callable(service_names):
            names.extend(str(item) for item in service_names())
    except Exception:
        pass
    joined = " ".join(names).lower()
    if joined and any(marker in joined for marker in markers):
        return True
    for marker in markers:
        for service in (
            f"com.sun.star.chart2.{marker.capitalize()}Diagram",
            f"com.sun.star.chart.{marker.capitalize()}Diagram",
        ):
            try:
                if diagram.supportsService(service):
                    return True
            except Exception:
                continue
    _log(f"_chart_type_matches: expected={chart_type} diagram_names={names}")
    return False


def _apply_chart_type(chart_document, chart_type: str) -> bool:
    """Set the chart diagram type."""
    chart_type = _normalize_chart_type(chart_type)
    diagram_services = [_chart_type_to_uno(chart_type)]
    classic_service = _chart_type_to_classic(chart_type)
    if classic_service:
        diagram_services.append(classic_service)

    factories = []
    direct_factory = getattr(chart_document, "createInstance", None)
    if callable(direct_factory):
        factories.append(("chart_document", direct_factory))
    try:
        get_factory = getattr(chart_document, "getFactory", None)
        if callable(get_factory):
            factory = get_factory()
            create = getattr(factory, "createInstance", None)
            if callable(create):
                factories.append(("chart_factory", create))
    except Exception as exc:
        _log(f"_apply_chart_type: getFactory failed: {exc}")

    for service_name in diagram_services:
        for factory_name, factory in factories:
            try:
                new_diagram = factory(service_name)
                if new_diagram is None:
                    continue
                _set_chart_3d_hint(new_diagram, chart_type)
                chart_document.setDiagram(new_diagram)
                if _chart_type_matches(chart_document, chart_type):
                    _log(f"_apply_chart_type: set diagram to {service_name} via {factory_name}")
                    return True
                _log(f"_apply_chart_type: {service_name} via {factory_name} did not stick")
            except Exception as exc:
                _log(f"_apply_chart_type: {service_name} via {factory_name} failed: {exc}")

    try:
        import uno

        ctx = uno.getComponentContext()
        smgr = ctx.getServiceManager() if ctx is not None else None
        for service_name in diagram_services:
            try:
                new_diagram = smgr.createInstanceWithContext(service_name, ctx)
                if new_diagram is None:
                    continue
                _set_chart_3d_hint(new_diagram, chart_type)
                chart_document.setDiagram(new_diagram)
                if _chart_type_matches(chart_document, chart_type):
                    _log(f"_apply_chart_type: set diagram to {service_name} via service_manager")
                    return True
            except Exception as exc:
                _log(f"_apply_chart_type: service manager {service_name} failed: {exc}")
    except Exception as exc:
        _log(f"_apply_chart_type: service manager fallback failed: {exc}")

    try:
        manager = getattr(chart_document, "getChartTypeManager", None)
        if callable(manager):
            mgr = manager()
            change = getattr(mgr, "changeDiagramType", None)
            if callable(change):
                change(chart_document, diagram_services[0])
                if _chart_type_matches(chart_document, chart_type):
                    _log(f"_apply_chart_type: set diagram to {diagram_services[0]} via manager")
                    return True
    except Exception as exc:
        _log(f"_apply_chart_type: manager fallback failed: {exc}")

    _log(f"_apply_chart_type: failed to apply requested chart type {chart_type}")
    return False


def _set_chart_data(chart_document, data_spec: Dict[str, Any]) -> bool:
    """Set chart data from specification."""
    categories = data_spec.get("categories", [])
    series_list = data_spec.get("series", [])

    if not categories or not series_list:
        _log("_set_chart_data: no data to set")
        return False

    series_labels: List[str] = []
    series_matrix: List[Tuple[float, ...]] = []
    max_length = len(categories)
    for series in series_list:
        if not isinstance(series, dict):
            continue
        values = series.get("values", [])
        normalized_values = []
        for value in values if isinstance(values, list) else []:
            try:
                normalized_values.append(float(value))
            except Exception:
                normalized_values.append(0.0)
        max_length = max(max_length, len(normalized_values))
        series_labels.append(str(series.get("name") or f"Series {len(series_labels) + 1}"))
        series_matrix.append(tuple(normalized_values))

    if not series_matrix:
        _log("_set_chart_data: no valid series values")
        return False

    padded_series = []
    for row in series_matrix:
        if len(row) < max_length:
            row = row + (0.0,) * (max_length - len(row))
        padded_series.append(row)

    padded_rows = []
    for row_idx in range(max_length):
        padded_rows.append(
            tuple(
                values[row_idx] if row_idx < len(values) else 0.0
                for values in padded_series
            )
        )

    row_labels_tuple = tuple(
        str(categories[row_idx]) if row_idx < len(categories) else f"Category {row_idx + 1}"
        for row_idx in range(max_length)
    )
    column_labels_tuple = tuple(series_labels)
    full_table_matrix = [[""] + list(column_labels_tuple)]
    for row_idx, row_values in enumerate(padded_rows):
        full_table_matrix.append([row_labels_tuple[row_idx]] + list(row_values))
    padded_tuple = tuple(padded_rows)

    changed = False
    chart_data = None
    try:
        getter = getattr(chart_document, "getData", None)
        if callable(getter):
            chart_data = getter()
    except Exception as exc:
        _log(f"_set_chart_data: getData failed: {exc}")

    if chart_data is not None:
        for payload, methods in (
            (full_table_matrix, ("setData",)),
            (padded_tuple, ("setDataArray", "setData")),
        ):
            if changed:
                break
            for method_name in methods:
                method = getattr(chart_data, method_name, None)
                if not callable(method):
                    continue
                try:
                    method(payload)
                    changed = True
                    _log(f"_set_chart_data: wrote chart data via getData().{method_name}")
                    break
                except Exception as exc:
                    _log(f"_set_chart_data: getData().{method_name} failed: {exc}")
        _set_chart_labels(chart_data, row_labels_tuple, column_labels_tuple)

    provider = None
    try:
        provider_getter = getattr(chart_document, "getDataProvider", None)
        if callable(provider_getter):
            provider = provider_getter()
    except Exception as exc:
        _log(f"_set_chart_data: getDataProvider failed: {exc}")

    if provider is not None:
        for method_name in ("setData", "setDataArray"):
            method = getattr(provider, method_name, None)
            if not callable(method):
                continue
            try:
                method(padded_tuple)
                changed = True
                _log(f"_set_chart_data: wrote chart data via provider.{method_name}")
                break
            except Exception as exc:
                _log(f"_set_chart_data: provider.{method_name} failed: {exc}")
        _set_chart_labels(provider, row_labels_tuple, column_labels_tuple)
        try:
            attach = getattr(chart_document, "attachDataProvider", None)
            if callable(attach):
                attach(provider)
                changed = True
                _log("_set_chart_data: attached chart data provider")
        except Exception as exc:
            _log(f"_set_chart_data: attachDataProvider failed: {exc}")
        if _set_chart_receiver_arguments(
            chart_document,
            has_series_labels=bool(column_labels_tuple),
            has_categories=bool(row_labels_tuple),
        ):
            changed = True

    if changed:
        _log(
            f"_set_chart_data: set {len(categories)} categories and {len(series_list)} series"
        )
    else:
        _log(
            "_set_chart_data: no writable chart data interface found "
            f"has_data={chart_data is not None} has_provider={provider is not None}"
        )
    return changed


def _set_chart_labels(target, row_labels: Tuple[str, ...], column_labels: Tuple[str, ...]) -> bool:
    changed = False
    if row_labels:
        for method_name in ("setRowDescriptions", "setAnyRowDescriptions"):
            method = getattr(target, method_name, None)
            if not callable(method):
                continue
            try:
                method(row_labels)
                changed = True
                break
            except Exception:
                continue
    if column_labels:
        for method_name in ("setColumnDescriptions", "setAnyColumnDescriptions"):
            method = getattr(target, method_name, None)
            if not callable(method):
                continue
            try:
                method(column_labels)
                changed = True
                break
            except Exception:
                continue
    return changed


def _set_chart_receiver_arguments(chart_document, has_series_labels: bool, has_categories: bool) -> bool:
    setter = getattr(chart_document, "setArguments", None)
    if not callable(setter):
        return False
    try:
        from com.sun.star.beans import PropertyValue

        setter(
            (
                PropertyValue(Name="CellRangeRepresentation", Value="all"),
                PropertyValue(Name="DataRowSource", Value=1),
                PropertyValue(Name="FirstCellAsLabel", Value=bool(has_series_labels)),
                PropertyValue(Name="HasCategories", Value=bool(has_categories)),
            )
        )
        _log("_set_chart_receiver_arguments: applied chart receiver arguments")
        return True
    except Exception as exc:
        _log(f"_set_chart_receiver_arguments: failed: {exc}")
        return False


def _set_chart_title(chart_document, title: str) -> None:
    """Set the chart title."""
    try:
        chart_document.setProperty("HasMainTitle", True)
        chart_document.getTitle().String = title
        _log(f"_set_chart_title: set title to '{title}'")
    except Exception as e:
        _log(f"_set_chart_title: failed to set title: {e}")


def _configure_chart_legend(chart_document, show_legend: bool) -> None:
    """Configure chart legend visibility."""
    try:
        chart_document.setProperty("HasLegend", show_legend)
        _log(f"_configure_chart_legend: legend visibility set to {show_legend}")
    except Exception as e:
        _log(f"_configure_chart_legend: failed to configure legend: {e}")


# ---------------------------------------------------------------------------
# Header/Footer preview functions
# ---------------------------------------------------------------------------


def _preview_set_header(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview setting header content."""
    _log(f"_preview_set_header: setting header for op {op_id}")

    header_type = action.get("header_type", "default")
    section_index = action.get("section_index", 0)
    content = action.get("content", {})

    try:
        # Get page style for the section
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            raise RuntimeError(f"Could not get page style for section {section_index}")

        # Store original header state for potential rejection
        original_header_on = page_style.HeaderIsOn
        original_content = None

        # Enable header if not already
        page_style.HeaderIsOn = True

        # Get the appropriate header text based on type
        if header_type == "first_page":
            page_style.HeaderIsShared = False
            header_text = page_style.HeaderTextFirst
        elif header_type == "even_page":
            page_style.HeaderIsShared = False
            header_text = page_style.HeaderTextLeft
        else:
            header_text = page_style.HeaderText

        if header_text is None:
            raise RuntimeError(f"Could not access header text for type {header_type}")

        # Clear and write new content
        header_text.setString("")
        cursor = header_text.createTextCursor()
        _write_paragraph_spec(
            cursor,
            content,
            color=GREEN_RGB,
            append_break=False,
            preserve_explicit_run_colors=True,
        )

        _log(f"_preview_set_header: header set successfully for op {op_id}")

        return {
            "type": "header_change",
            "header_type": header_type,
            "section_index": section_index,
            "original_was_on": original_header_on,
            "proposed_spec": content,
        }

    except Exception as e:
        _log(f"_preview_set_header: failed to set header: {e}")
        return None


def _preview_set_footer(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview setting footer content."""
    _log(f"_preview_set_footer: setting footer for op {op_id}")

    footer_type = action.get("footer_type", "default")
    section_index = action.get("section_index", 0)
    content = action.get("content", {})

    try:
        # Get page style for the section
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            raise RuntimeError(f"Could not get page style for section {section_index}")

        # Store original footer state for potential rejection
        original_footer_on = page_style.FooterIsOn

        # Enable footer if not already
        page_style.FooterIsOn = True

        # Get the appropriate footer text based on type
        if footer_type == "first_page":
            page_style.FooterIsShared = False
            footer_text = page_style.FooterTextFirst
        elif footer_type == "even_page":
            page_style.FooterIsShared = False
            footer_text = page_style.FooterTextLeft
        else:
            footer_text = page_style.FooterText

        if footer_text is None:
            raise RuntimeError(f"Could not access footer text for type {footer_type}")

        # Clear and write new content
        footer_text.setString("")
        cursor = footer_text.createTextCursor()
        _write_paragraph_spec(
            cursor,
            content,
            color=GREEN_RGB,
            append_break=False,
            preserve_explicit_run_colors=True,
        )

        _log(f"_preview_set_footer: footer set successfully for op {op_id}")

        return {
            "type": "footer_change",
            "footer_type": footer_type,
            "section_index": section_index,
            "original_was_on": original_footer_on,
            "proposed_spec": content,
        }

    except Exception as e:
        _log(f"_preview_set_footer: failed to set footer: {e}")
        return None


def _preview_clear_header(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview clearing a header."""
    _log(f"_preview_clear_header: clearing header for op {op_id}")

    header_type = action.get("header_type", "default")
    section_index = action.get("section_index", 0)

    try:
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            raise RuntimeError(f"Could not get page style for section {section_index}")

        original_was_on = page_style.HeaderIsOn

        # Get and clear the appropriate header text
        if header_type == "first_page":
            header_text = page_style.HeaderTextFirst
        elif header_type == "even_page":
            header_text = page_style.HeaderTextLeft
        else:
            header_text = page_style.HeaderText

        if header_text:
            header_text.setString("")

        _log(f"_preview_clear_header: header cleared for op {op_id}")

        return {
            "type": "header_clear",
            "header_type": header_type,
            "section_index": section_index,
            "original_was_on": original_was_on,
        }

    except Exception as e:
        _log(f"_preview_clear_header: failed to clear header: {e}")
        return None


def _preview_clear_footer(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview clearing a footer."""
    _log(f"_preview_clear_footer: clearing footer for op {op_id}")

    footer_type = action.get("footer_type", "default")
    section_index = action.get("section_index", 0)

    try:
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            raise RuntimeError(f"Could not get page style for section {section_index}")

        original_was_on = page_style.FooterIsOn

        # Get and clear the appropriate footer text
        if footer_type == "first_page":
            footer_text = page_style.FooterTextFirst
        elif footer_type == "even_page":
            footer_text = page_style.FooterTextLeft
        else:
            footer_text = page_style.FooterText

        if footer_text:
            footer_text.setString("")

        _log(f"_preview_clear_footer: footer cleared for op {op_id}")

        return {
            "type": "footer_clear",
            "footer_type": footer_type,
            "section_index": section_index,
            "original_was_on": original_was_on,
        }

    except Exception as e:
        _log(f"_preview_clear_footer: failed to clear footer: {e}")
        return None


def _get_page_style_for_section(doc, section_index: int = 0):
    """
    Get the page style for a given section.
    In most documents, this is 'Standard'. For multi-section docs,
    we'd need to enumerate sections and their styles.
    """
    try:
        page_styles = doc.getStyleFamilies().getByName("PageStyles")

        # For section_index 0, use the default "Standard" page style
        # which covers most use cases
        if section_index == 0:
            if page_styles.hasByName("Standard"):
                return page_styles.getByName("Standard")
            # Try "Default Style" as fallback
            if page_styles.hasByName("Default Style"):
                return page_styles.getByName("Default Style")

        # For other sections, we'd need more complex logic
        # to find the actual section and its page style
        # For now, return the first available style
        style_names = page_styles.getElementNames()
        if style_names:
            return page_styles.getByName(style_names[0])

        return None

    except Exception as e:
        _log(f"_get_page_style_for_section: error getting page style: {e}")
        return None


# ---------------------------------------------------------------------------
# Field preview functions
# ---------------------------------------------------------------------------


def _preview_insert_field(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview inserting a document field."""
    _log(f"_preview_insert_field: inserting field for op {op_id}")

    field_type = action.get("field_type", "page_number")
    location = action.get("location", "body")

    try:
        # Determine where to insert the field
        if location == "header":
            text = _get_header_footer_text(doc, action, is_header=True)
        elif location == "footer":
            text = _get_header_footer_text(doc, action, is_header=False)
        else:
            text = doc.getText()

        if text is None:
            raise RuntimeError(f"Could not get text for location {location}")

        cursor = text.createTextCursor()
        cursor.gotoEnd(False)

        # Create and insert the field
        field = _create_field(doc, field_type, action)
        if field is None:
            raise RuntimeError(f"Could not create field of type {field_type}")

        text.insertTextContent(cursor, field, False)

        _log(f"_preview_insert_field: field {field_type} inserted for op {op_id}")

        return {
            "type": "field_insertion",
            "field_type": field_type,
            "location": location,
        }

    except Exception as e:
        _log(f"_preview_insert_field: failed to insert field: {e}")
        return None


def _create_field(doc, field_type: str, action: Dict[str, Any]):
    """Create a document field based on type."""
    try:
        if field_type == "page_number":
            field = doc.createInstance("com.sun.star.text.TextField.PageNumber")
            field.NumberingType = 4  # Arabic numerals
        elif field_type == "page_count":
            field = doc.createInstance("com.sun.star.text.TextField.PageCount")
            field.NumberingType = 4
        elif field_type == "date":
            field = doc.createInstance("com.sun.star.text.TextField.DateTime")
            field.IsDate = True
            field.IsFixed = False
        elif field_type == "time":
            field = doc.createInstance("com.sun.star.text.TextField.DateTime")
            field.IsDate = False
            field.IsFixed = False
        elif field_type == "date_time":
            field = doc.createInstance("com.sun.star.text.TextField.DateTime")
            field.IsDate = False  # Shows both date and time
            field.IsFixed = False
        elif field_type == "filename":
            field = doc.createInstance("com.sun.star.text.TextField.FileName")
            field.FileFormat = 1  # Name only, no path
        elif field_type == "filename_path":
            field = doc.createInstance("com.sun.star.text.TextField.FileName")
            field.FileFormat = 0  # Full path
        elif field_type == "author":
            field = doc.createInstance("com.sun.star.text.TextField.Author")
        elif field_type == "title":
            field = doc.createInstance("com.sun.star.text.TextField.DocInfo.Title")
        else:
            _log(f"_create_field: unknown field type {field_type}")
            return None

        return field

    except Exception as e:
        _log(f"_create_field: error creating field: {e}")
        return None


def _get_header_footer_text(doc, action: Dict[str, Any], is_header: bool):
    """Get header or footer text object for field insertion."""
    hf_type = action.get("header_footer_type", "default")
    section_index = action.get("section_index", 0)

    try:
        page_style = _get_page_style_for_section(doc, section_index)
        if page_style is None:
            return None

        if is_header:
            page_style.HeaderIsOn = True
            if hf_type == "first_page":
                page_style.HeaderIsShared = False
                return page_style.HeaderTextFirst
            elif hf_type == "even_page":
                page_style.HeaderIsShared = False
                return page_style.HeaderTextLeft
            else:
                return page_style.HeaderText
        else:
            page_style.FooterIsOn = True
            if hf_type == "first_page":
                page_style.FooterIsShared = False
                return page_style.FooterTextFirst
            elif hf_type == "even_page":
                page_style.FooterIsShared = False
                return page_style.FooterTextLeft
            else:
                return page_style.FooterText

    except Exception as e:
        _log(f"_get_header_footer_text: error: {e}")
        return None


# ---------------------------------------------------------------------------
# Table cell merge/split preview functions
# ---------------------------------------------------------------------------


def _preview_merge_cells(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview merging table cells."""
    _log(f"_preview_merge_cells: merging cells for op {op_id}")

    try:
        table = _locate_table(doc, action)
        if table is None:
            raise RuntimeError("Target table not found for merge")

        start_row = action.get("start_row", 0)
        start_col = action.get("start_col", 0)
        end_row = action.get("end_row", 0)
        end_col = action.get("end_col", 0)

        # Get cell range name (e.g., "A1:B2")
        start_cell_name = _cell_address_to_name(start_row, start_col)
        end_cell_name = _cell_address_to_name(end_row, end_col)
        range_name = f"{start_cell_name}:{end_cell_name}"

        _log(f"_preview_merge_cells: merging range {range_name}")

        # Get the cell range and merge
        cell_range = table.getCellRangeByName(range_name)
        if cell_range is not None:
            cell_range.merge(True)

        _log(f"_preview_merge_cells: cells merged for op {op_id}")

        return {
            "type": "cell_merge",
            "range_name": range_name,
            "start_row": start_row,
            "start_col": start_col,
            "end_row": end_row,
            "end_col": end_col,
        }

    except Exception as e:
        _log(f"_preview_merge_cells: failed: {e}")
        return None


def _preview_split_cell(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview splitting a table cell."""
    _log(f"_preview_split_cell: splitting cell for op {op_id}")

    try:
        table = _locate_table(doc, action)
        if table is None:
            raise RuntimeError("Target table not found for split")

        row = action.get("row", 0)
        col = action.get("col", 0)
        split_rows = action.get("split_rows", 1)
        split_cols = action.get("split_cols", 2)

        cell_name = _cell_address_to_name(row, col)
        _log(
            f"_preview_split_cell: splitting cell {cell_name} into {split_rows}x{split_cols}"
        )

        # Get the cell and split
        cell = table.getCellByName(cell_name)
        if cell is not None:
            # LibreOffice Writer tables support horizontal split via cursor
            # For more complex splits, we would need to manipulate the table structure
            cell_range = table.getCellRangeByName(cell_name)
            if cell_range is not None and split_cols > 1:
                cell_range.split(split_cols, split_rows > 1)

        _log(f"_preview_split_cell: cell split for op {op_id}")

        return {
            "type": "cell_split",
            "cell_name": cell_name,
            "row": row,
            "col": col,
            "split_rows": split_rows,
            "split_cols": split_cols,
        }

    except Exception as e:
        _log(f"_preview_split_cell: failed: {e}")
        return None


def _cell_address_to_name(row: int, col: int) -> str:
    """Convert row/column indices to cell name (e.g., 0,0 -> 'A1')."""
    # Column: A-Z, then AA, AB, etc.
    col_name = ""
    c = col
    while c >= 0:
        col_name = chr(ord("A") + (c % 26)) + col_name
        c = c // 26 - 1
        if c < 0:
            break
    # Row: 1-based
    return f"{col_name}{row + 1}"


def _locate_table(doc, action: Dict[str, Any]):
    """Locate a table by element_id (bookmark) or element_index."""
    element_id = action.get("element_id")
    element_index = action.get("element_index")

    if element_id:
        # Resolve using table-preferred bookmark semantics. This lets table IDs
        # anchored inside a first-cell paragraph still resolve to the owning table.
        element = _get_element_by_bookmark(doc, element_id, prefer_table=True)
        if element and element[0] == "table":
            return element[1]

    if element_index is not None:
        # Find table by index
        element = _get_element_by_index(doc, element_index)
        if element and element[0] == "table":
            return element[1]

    return None


# ---------------------------------------------------------------------------
# Section break preview functions
# ---------------------------------------------------------------------------


def _preview_insert_section_break(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Preview inserting a section break."""
    _log(f"_preview_insert_section_break: inserting section break for op {op_id}")

    break_type = action.get("break_type", "next_page")

    try:
        # Get cursor position
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            # Insert at end of document
            text = doc.getText()
            cursor = text.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text = paragraph.getText()
            cursor = text.createTextCursorByRange(paragraph.getEnd())

        # Insert the section break. In LibreOffice, a bounded multi-column
        # region must be a real TextSection with a TextColumns struct.
        section = doc.createInstance("com.sun.star.text.TextSection")
        section_name = action.get("section_name") or f"sdoc_section_{op_id}".replace(
            "-", "_"
        )
        try:
            section.setName(section_name)
        except Exception:
            pass

        # Configure break type
        if break_type == "next_page":
            # Insert a page break before section
            cursor.BreakType = 4  # PAGE_AFTER
        elif break_type == "continuous":
            # No break, section continues on same page
            pass
        elif break_type == "even_page":
            cursor.BreakType = 4
        elif break_type == "odd_page":
            cursor.BreakType = 4

        # Apply optional section properties
        orientation = action.get("orientation")
        columns = action.get("columns")

        if orientation == "landscape":
            # Would need to create a new page style
            _log("_preview_insert_section_break: landscape orientation requested")

        # Insert the section
        text.insertTextContent(cursor, section, False)

        if columns and columns > 1:
            try:
                section_columns = _get_text_columns(section)
                section.TextColumns = _configure_text_columns(
                    section_columns,
                    columns,
                )
            except Exception as column_exc:
                _log(
                    "_preview_insert_section_break: failed to set section columns "
                    f"for {section_name}: {column_exc}"
                )

        _log(f"_preview_insert_section_break: section break inserted for op {op_id}")

        return {
            "type": "section_break_insertion",
            "break_type": break_type,
            "orientation": orientation,
            "columns": columns,
            "section_name": section_name,
        }

    except Exception as e:
        _log(f"_preview_insert_section_break: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Track Changes preview functions (Phase 2)
# ---------------------------------------------------------------------------


def _preview_enable_track_changes(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Enable or disable track changes mode."""
    _log(f"_preview_enable_track_changes: for op {op_id}")

    enabled = action.get("enabled", True)
    author = action.get("author")

    try:
        # LibreOffice uses RecordChanges property
        doc.RecordChanges = enabled

        if author:
            # Set document properties for author
            props = doc.getDocumentProperties()
            props.Author = author

        _log(
            f"_preview_enable_track_changes: track changes {'enabled' if enabled else 'disabled'}"
        )

        return {
            "type": "track_changes_toggle",
            "enabled": enabled,
            "author": author,
        }

    except Exception as e:
        _log(f"_preview_enable_track_changes: failed: {e}")
        return None


def _preview_accept_revision(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Accept a single revision."""
    _log(f"_preview_accept_revision: for op {op_id}")

    revision_id = action.get("revision_id")

    try:
        # Get redlines supplier
        redlines = doc.getRedlines()
        if redlines is None:
            raise RuntimeError("Document does not support redlines")

        # Find and accept the revision
        for i in range(redlines.getCount()):
            redline = redlines.getByIndex(i)
            if str(redline.getPropertyValue("Identifier")) == revision_id:
                redline.accept()
                _log(f"_preview_accept_revision: accepted revision {revision_id}")
                return {
                    "type": "revision_accept",
                    "revision_id": revision_id,
                }

        _log(f"_preview_accept_revision: revision {revision_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_accept_revision: failed: {e}")
        return None


def _preview_reject_revision(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Reject a single revision."""
    _log(f"_preview_reject_revision: for op {op_id}")

    revision_id = action.get("revision_id")

    try:
        redlines = doc.getRedlines()
        if redlines is None:
            raise RuntimeError("Document does not support redlines")

        for i in range(redlines.getCount()):
            redline = redlines.getByIndex(i)
            if str(redline.getPropertyValue("Identifier")) == revision_id:
                redline.reject()
                _log(f"_preview_reject_revision: rejected revision {revision_id}")
                return {
                    "type": "revision_reject",
                    "revision_id": revision_id,
                }

        _log(f"_preview_reject_revision: revision {revision_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_reject_revision: failed: {e}")
        return None


def _preview_accept_all_revisions(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Accept all revisions in the document."""
    _log(f"_preview_accept_all_revisions: for op {op_id}")

    try:
        redlines = doc.getRedlines()
        if redlines is None:
            raise RuntimeError("Document does not support redlines")

        count = redlines.getCount()
        # Accept from end to start to avoid index shifting
        for i in range(count - 1, -1, -1):
            try:
                redline = redlines.getByIndex(i)
                redline.accept()
            except Exception:
                pass

        _log(f"_preview_accept_all_revisions: accepted {count} revisions")

        return {
            "type": "revisions_accept_all",
            "count": count,
        }

    except Exception as e:
        _log(f"_preview_accept_all_revisions: failed: {e}")
        return None


def _preview_reject_all_revisions(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Reject all revisions in the document."""
    _log(f"_preview_reject_all_revisions: for op {op_id}")

    try:
        redlines = doc.getRedlines()
        if redlines is None:
            raise RuntimeError("Document does not support redlines")

        count = redlines.getCount()
        # Reject from end to start to avoid index shifting
        for i in range(count - 1, -1, -1):
            try:
                redline = redlines.getByIndex(i)
                redline.reject()
            except Exception:
                pass

        _log(f"_preview_reject_all_revisions: rejected {count} revisions")

        return {
            "type": "revisions_reject_all",
            "count": count,
        }

    except Exception as e:
        _log(f"_preview_reject_all_revisions: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Comment preview functions (Phase 2)
# ---------------------------------------------------------------------------


def _preview_insert_comment(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a new comment anchored to text."""
    _log(f"_preview_insert_comment: for op {op_id}")

    author = action.get("author", "Unknown")
    text = action.get("text", "")
    date_str = action.get("date")

    try:
        # Locate the element to anchor the comment (paragraph preferred, table supported).
        element_info = _locate_element(doc, action)
        if element_info and element_info[0] == "paragraph":
            paragraph = element_info[1]
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
        elif element_info and element_info[0] == "table":
            table = element_info[1]
            anchor = table.getAnchor()
            text_obj = anchor.getText() or doc.getText()
            cursor = text_obj.createTextCursorByRange(anchor)
            cursor.collapseToStart()
        else:
            # Fallback to end of document
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)

        # Apply offsets if specified (only meaningful for text cursors).
        start_offset = action.get("start_offset", 0)
        end_offset = action.get("end_offset")
        if start_offset > 0:
            cursor.goRight(start_offset, False)
        if end_offset:
            cursor.goRight(end_offset - start_offset, True)

        # Create the annotation (comment)
        annotation = doc.createInstance("com.sun.star.text.textfield.Annotation")
        annotation.Author = author
        annotation.Content = text

        # Set date BEFORE insertion - setting after causes SIGSEGV crash
        try:
            import uno

            # Parse ISO date string or use current datetime if not provided
            if date_str:
                parsed = dt.fromisoformat(date_str.replace("Z", "+00:00"))
            else:
                parsed = dt.now()
            # Create UNO DateTime struct using uno.createUnoStruct (more compatible)
            uno_dt = uno.createUnoStruct("com.sun.star.util.DateTime")
            uno_dt.Year = parsed.year
            uno_dt.Month = parsed.month
            uno_dt.Day = parsed.day
            uno_dt.Hours = parsed.hour
            uno_dt.Minutes = parsed.minute
            uno_dt.Seconds = parsed.second
            uno_dt.NanoSeconds = 0
            uno_dt.IsUTC = False
            annotation.DateTimeValue = uno_dt
            _log(f"_preview_insert_comment: set date to {parsed}")
        except Exception as date_err:
            _log(f"_preview_insert_comment: failed to set date: {date_err}")

        # Now insert the annotation
        text_obj.insertTextContent(cursor, annotation, False)

        _log(f"_preview_insert_comment: comment inserted for op {op_id}")

        return {
            "type": "comment_insertion",
            "comment_id": str(id(annotation)),  # Return ID for rejection
            "author": author,
            "text": text[:50] + "..." if len(text) > 50 else text,
            "date": date_str,
        }

    except Exception as e:
        _log(f"_preview_insert_comment: failed: {e}")
        return None


def _preview_reply_to_comment(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Reply to an existing comment."""
    _log(f"_preview_reply_to_comment: for op {op_id}")

    comment_id = action.get("comment_id")
    author = action.get("author", "Unknown")
    reply_text = action.get("text", "")

    try:
        # Find the comment by ID
        text_fields = doc.getTextFields()
        enum = text_fields.createEnumeration()

        while enum.hasMoreElements():
            field = enum.nextElement()
            if field.supportsService("com.sun.star.text.textfield.Annotation"):
                # Check if this is the target comment
                field_name = str(
                    field.getPropertyValue("Name") if hasattr(field, "Name") else ""
                )
                if field_name == comment_id or str(id(field)) == comment_id:
                    # Append reply to comment content
                    existing = field.Content
                    field.Content = f"{existing}\n\n[Reply from {author}]: {reply_text}"

                    _log(f"_preview_reply_to_comment: replied to comment {comment_id}")
                    return {
                        "type": "comment_reply",
                        "comment_id": comment_id,
                        "author": author,
                    }

        _log(f"_preview_reply_to_comment: comment {comment_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_reply_to_comment: failed: {e}")
        return None


def _preview_resolve_comment(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Mark a comment as resolved."""
    _log(f"_preview_resolve_comment: for op {op_id}")

    comment_id = action.get("comment_id")
    resolved = action.get("resolved", True)

    try:
        text_fields = doc.getTextFields()
        enum = text_fields.createEnumeration()

        while enum.hasMoreElements():
            field = enum.nextElement()
            if field.supportsService("com.sun.star.text.textfield.Annotation"):
                field_name = str(
                    field.getPropertyValue("Name") if hasattr(field, "Name") else ""
                )
                if field_name == comment_id or str(id(field)) == comment_id:
                    # Mark as resolved by prefixing content
                    if resolved:
                        if not field.Content.startswith("[RESOLVED]"):
                            field.Content = f"[RESOLVED] {field.Content}"
                    else:
                        field.Content = field.Content.replace("[RESOLVED] ", "")

                    _log(
                        f"_preview_resolve_comment: {'resolved' if resolved else 'unresolved'} comment {comment_id}"
                    )
                    return {
                        "type": "comment_resolve",
                        "comment_id": comment_id,
                        "resolved": resolved,
                    }

        _log(f"_preview_resolve_comment: comment {comment_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_resolve_comment: failed: {e}")
        return None


def _preview_delete_comment(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Delete a comment."""
    _log(f"_preview_delete_comment: for op {op_id}")

    comment_id = action.get("comment_id")

    try:
        text_fields = doc.getTextFields()
        enum = text_fields.createEnumeration()

        while enum.hasMoreElements():
            field = enum.nextElement()
            if field.supportsService("com.sun.star.text.textfield.Annotation"):
                field_name = str(
                    field.getPropertyValue("Name") if hasattr(field, "Name") else ""
                )
                if field_name == comment_id or str(id(field)) == comment_id:
                    # Remove the annotation
                    anchor = field.getAnchor()
                    anchor.getText().removeTextContent(field)

                    _log(f"_preview_delete_comment: deleted comment {comment_id}")
                    return {
                        "type": "comment_deletion",
                        "comment_id": comment_id,
                    }

        _log(f"_preview_delete_comment: comment {comment_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_delete_comment: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Style management preview functions (Phase 2)
# ---------------------------------------------------------------------------


def _preview_create_style(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Create a new document style."""
    _log(f"_preview_create_style: for op {op_id}")

    style_name = action.get("style_name")
    style_type = action.get("style_type", "paragraph")

    try:
        # Get the appropriate style family
        style_families = doc.getStyleFamilies()

        family_name = {
            "paragraph": "ParagraphStyles",
            "character": "CharacterStyles",
            "table": "TableStyles",
            "list": "NumberingStyles",
        }.get(style_type, "ParagraphStyles")

        if not style_families.hasByName(family_name):
            raise RuntimeError(f"Style family {family_name} not found")

        family = style_families.getByName(family_name)

        # Create the new style
        if style_type == "paragraph":
            style = doc.createInstance("com.sun.star.style.ParagraphStyle")
        elif style_type == "character":
            style = doc.createInstance("com.sun.star.style.CharacterStyle")
        else:
            style = doc.createInstance("com.sun.star.style.ParagraphStyle")

        # Apply style properties
        _apply_style_properties(style, action)

        # Set parent style
        based_on = action.get("based_on")
        if based_on and family.hasByName(based_on):
            style.ParentStyle = based_on

        # Insert into family
        family.insertByName(style_name, style)

        _log(f"_preview_create_style: created style {style_name}")

        return {
            "type": "style_creation",
            "style_name": style_name,
            "style_type": style_type,
        }

    except Exception as e:
        _log(f"_preview_create_style: failed: {e}")
        return None


def _preview_modify_style(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Modify an existing document style."""
    _log(f"_preview_modify_style: for op {op_id}")

    style_name = action.get("style_name")

    try:
        style_families = doc.getStyleFamilies()

        # Search for the style in all families
        for family_name in ["ParagraphStyles", "CharacterStyles"]:
            if style_families.hasByName(family_name):
                family = style_families.getByName(family_name)
                if family.hasByName(style_name):
                    style = family.getByName(style_name)
                    _apply_style_properties(style, action)

                    _log(f"_preview_modify_style: modified style {style_name}")
                    return {
                        "type": "style_modification",
                        "style_name": style_name,
                    }

        _log(f"_preview_modify_style: style {style_name} not found")
        return None

    except Exception as e:
        _log(f"_preview_modify_style: failed: {e}")
        return None


def _preview_delete_style(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Delete a document style."""
    _log(f"_preview_delete_style: for op {op_id}")

    style_name = action.get("style_name")

    try:
        style_families = doc.getStyleFamilies()

        for family_name in ["ParagraphStyles", "CharacterStyles"]:
            if style_families.hasByName(family_name):
                family = style_families.getByName(family_name)
                if family.hasByName(style_name):
                    family.removeByName(style_name)

                    _log(f"_preview_delete_style: deleted style {style_name}")
                    return {
                        "type": "style_deletion",
                        "style_name": style_name,
                    }

        _log(f"_preview_delete_style: style {style_name} not found")
        return None

    except Exception as e:
        _log(f"_preview_delete_style: failed: {e}")
        return None


def _apply_style_properties(style, action: Dict[str, Any]) -> None:
    """Apply style properties from action to a style object."""
    try:
        if action.get("font_name"):
            style.CharFontName = action["font_name"]
        if action.get("font_size_pt"):
            style.CharHeight = action["font_size_pt"]
        if action.get("bold") is not None:
            style.CharWeight = 150 if action["bold"] else 100
        if action.get("italic") is not None:
            style.CharPosture = 2 if action["italic"] else 0
        if action.get("font_color"):
            style.CharColor = int(action["font_color"].lstrip("#"), 16)
        if action.get("background_color"):
            style.CharBackColor = int(action["background_color"].lstrip("#"), 16)
        if action.get("alignment"):
            align_map = {"left": 0, "center": 3, "right": 1, "justify": 2}
            style.ParaAdjust = align_map.get(action["alignment"], 0)
        if action.get("space_before_pt"):
            style.ParaTopMargin = int(
                action["space_before_pt"] * 35.28
            )  # pt to 1/100 mm
        if action.get("space_after_pt"):
            style.ParaBottomMargin = int(action["space_after_pt"] * 35.28)
    except Exception as e:
        _log(f"_apply_style_properties: error: {e}")


# ---------------------------------------------------------------------------
# Footnote/Endnote preview functions (Phase 2)
# ---------------------------------------------------------------------------


def _preview_insert_footnote(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a footnote at the specified position."""
    _log(f"_preview_insert_footnote: for op {op_id}")

    text_content = action.get("text", "")
    label = action.get("label")

    try:
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
            char_offset = action.get("char_offset", 0)
            if char_offset > 0:
                cursor.goRight(char_offset, False)

        # Create footnote
        footnote = doc.createInstance("com.sun.star.text.Footnote")
        if label:
            footnote.Label = label

        # Insert footnote reference
        text_obj.insertTextContent(cursor, footnote, False)

        # Add text to footnote
        footnote_text = footnote.getText()
        footnote_cursor = footnote_text.createTextCursor()
        footnote_text.insertString(footnote_cursor, text_content, False)

        _log(f"_preview_insert_footnote: inserted footnote for op {op_id}")

        return {
            "type": "footnote_insertion",
            "note_id": str(id(footnote)),  # Return ID for rejection
            "note_type": "footnote",
            "text": (
                text_content[:50] + "..." if len(text_content) > 50 else text_content
            ),
        }

    except Exception as e:
        _log(f"_preview_insert_footnote: failed: {e}")
        return None


def _preview_insert_endnote(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert an endnote at the specified position."""
    _log(f"_preview_insert_endnote: for op {op_id}")

    text_content = action.get("text", "")
    label = action.get("label")

    try:
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
            char_offset = action.get("char_offset", 0)
            if char_offset > 0:
                cursor.goRight(char_offset, False)

        # Create endnote
        endnote = doc.createInstance("com.sun.star.text.Endnote")
        if label:
            endnote.Label = label

        # Insert endnote reference
        text_obj.insertTextContent(cursor, endnote, False)

        # Add text to endnote
        endnote_text = endnote.getText()
        endnote_cursor = endnote_text.createTextCursor()
        endnote_text.insertString(endnote_cursor, text_content, False)

        _log(f"_preview_insert_endnote: inserted endnote for op {op_id}")

        return {
            "type": "endnote_insertion",
            "note_id": str(id(endnote)),  # Return ID for rejection
            "note_type": "endnote",
            "text": (
                text_content[:50] + "..." if len(text_content) > 50 else text_content
            ),
        }

    except Exception as e:
        _log(f"_preview_insert_endnote: failed: {e}")
        return None


def _preview_modify_note(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Modify an existing footnote or endnote."""
    _log(f"_preview_modify_note: for op {op_id}")

    note_id = action.get("note_id")
    note_type = action.get("note_type", "footnote")
    new_text = action.get("text")

    try:
        if note_type == "footnote":
            notes = doc.getFootnotes()
        else:
            notes = doc.getEndnotes()

        for i in range(notes.getCount()):
            note = notes.getByIndex(i)
            if str(i) == note_id or str(id(note)) == note_id:
                if new_text is not None:
                    note_text = note.getText()
                    note_cursor = note_text.createTextCursor()
                    note_cursor.gotoStart(False)
                    note_cursor.gotoEnd(True)
                    note_text.insertString(note_cursor, new_text, True)

                _log(f"_preview_modify_note: modified {note_type} {note_id}")
                return {
                    "type": "note_modification",
                    "note_type": note_type,
                    "note_id": note_id,
                }

        _log(f"_preview_modify_note: {note_type} {note_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_modify_note: failed: {e}")
        return None


def _preview_delete_note(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Delete a footnote or endnote."""
    _log(f"_preview_delete_note: for op {op_id}")

    note_id = action.get("note_id")
    note_type = action.get("note_type", "footnote")

    try:
        if note_type == "footnote":
            notes = doc.getFootnotes()
        else:
            notes = doc.getEndnotes()

        for i in range(notes.getCount()):
            note = notes.getByIndex(i)
            if str(i) == note_id or str(id(note)) == note_id:
                anchor = note.getAnchor()
                anchor.getText().removeTextContent(note)

                _log(f"_preview_delete_note: deleted {note_type} {note_id}")
                return {
                    "type": "note_deletion",
                    "note_type": note_type,
                    "note_id": note_id,
                }

        _log(f"_preview_delete_note: {note_type} {note_id} not found")
        return None

    except Exception as e:
        _log(f"_preview_delete_note: failed: {e}")
        return None

    except Exception as e:
        _log(f"_preview_delete_note: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Image preview functions (Phase 2)
# ---------------------------------------------------------------------------


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


def _preview_insert_image(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert an image with specific positioning."""
    _log(f"_preview_insert_image: for op {op_id}")

    image_url = action.get("image_url")
    width_mm = action.get("width_mm", 50)
    height_mm = action.get("height_mm", 50)
    # Always use as_character (inline) so images sit in the text flow and
    # paragraph centering applies. Read from top-level first (backend-forced),
    # then nested image dict, then default to as_character.
    _image_dict = action.get("image") or {}
    anchor_type_str = action.get("anchor_type") or _image_dict.get("anchor_type") or "as_character"
    wrap_type_str = action.get("wrap_type") or _image_dict.get("wrap_type") or "none"
    pos_x = action.get("pos_x_mm", 0)
    pos_y = action.get("pos_y_mm", 0)

    try:
        # Locate anchor
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
            char_offset = action.get("char_offset", 0)
            if char_offset > 0:
                cursor.goRight(char_offset, False)

        # Download and embed image properly
        temp_path = _download_image_to_temp(image_url)
        if temp_path:
            try:
                # Use GraphicProvider to embed the graphic
                graphic_provider = XSCRIPTCONTEXT.getComponentContext().ServiceManager.createInstanceWithContext( # type: ignore # NOQA
                    "com.sun.star.graphic.GraphicProvider", XSCRIPTCONTEXT.getComponentContext() # type: ignore # NOQA
                )
                from com.sun.star.beans import PropertyValue
                local_url = Path(temp_path).as_uri()
                props = (PropertyValue(Name="URL", Value=local_url),)
                graphic = graphic_provider.queryGraphic(props)
                
                # Create graphic object
                image = doc.createInstance("com.sun.star.text.TextGraphicObject")
                if graphic:
                    image.Graphic = graphic
                    _log(f"_preview_insert_image: graphic embedded for {image_url}")
                else:
                    image.GraphicURL = local_url
                    _log(f"_preview_insert_image: using GraphicURL fallback for {image_url}")
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass
        else:
            # Fallback to direct URL (might fail with "Read Error" if blocked)
            image = doc.createInstance("com.sun.star.text.TextGraphicObject")
            image.GraphicURL = image_url
            _log(f"_preview_insert_image: fallback to direct URL {image_url}")

        # All properties MUST be set before insertTextContent — AnchorType especially
        # is ignored if set after the object is already anchored in the document.

        # AT_PARAGRAPH=0, AS_CHARACTER=1, AT_PAGE=2, AT_FRAME=3, AT_CHARACTER=4
        anchor_map = {
            "at_paragraph": 0,
            "as_character": 1,
            "at_page": 2,
            "at_character": 4,
        }
        anchor_type_val = anchor_map.get(anchor_type_str, 1)  # default AS_CHARACTER
        image.AnchorType = anchor_type_val

        # NONE=0, THROUGH=1, PARALLEL=2, DYNAMIC=3
        wrap_map = {
            "none": 0,
            "through": 1,
            "parallel": 2,
            "dynamic": 3,
            "behind": 1,
            "in_front": 1,
        }
        image.Surround = wrap_map.get(wrap_type_str, 0)  # default none

        # Size (1/100 mm)
        image.Width = int(width_mm * 100)
        image.Height = int(height_mm * 100)

        if wrap_type_str == "behind":
            image.Opaque = False

        # For floating images, set position before insert
        if anchor_type_val != 1:  # not AS_CHARACTER
            if pos_x == 0 and pos_y == 0:
                image.HoriOrient = 1  # CENTER
                image.VertOrient = 0
            else:
                image.HoriOrientPosition = int(pos_x * 100)
                image.VertOrientPosition = int(pos_y * 100)
                image.HoriOrient = 0
                image.VertOrient = 0

        # Set Name before insert
        image_name = f"Image_{op_id}"
        image.Name = image_name

        # For AS_CHARACTER: center the paragraph before inserting so the image
        # inherits the centered paragraph context. Also add bottom margin so
        # the caption paragraph below has a visible gap from the image.
        if anchor_type_val == 1:
            try:
                cursor.ParaAdjust = PAR_ADJUST_CENTER
                cursor.ParaBottomMargin = 212  # ~6pt space between image and caption
            except Exception:
                pass

        # Insert — all properties already set
        text_obj.insertTextContent(cursor, image, False)
        _apply_centered_inline_image_placement(image, "_preview_insert_image")

        # Insert a paragraph break after the image so any following content
        # (e.g. a caption paragraph) starts on its own new line, not to the
        # right of the image in the remaining line space.
        text_obj.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

        _log(f"_preview_insert_image: inserted image {image_url}")

        return {
            "type": "image_insertion",
            "image_id": image_name,  # Return Name for rejection
            "image_url": image_url,
        }

    except Exception as e:
        _log(f"_preview_insert_image: failed: {e}")
        return None


def _preview_replace_image(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Replace an image or update its properties."""
    _log(f"_preview_replace_image: for op {op_id}")

    try:
        image, resolved_index, resolved_name = _resolve_graphic_target(doc, action)
        if image is not None:
            _image_dict = action.get("image") or {}
            anchor_type_str = action.get("anchor_type") or _image_dict.get("anchor_type") or "as_character"
            wrap_type_str = action.get("wrap_type") or _image_dict.get("wrap_type") or "none"
            has_explicit_position = (
                action.get("pos_x_mm") is not None or action.get("pos_y_mm") is not None
            )

            # Apply updates
            if action.get("image_url"):
                if not _embed_image_url_into_graphic(
                    image,
                    action["image_url"],
                    "_preview_replace_image",
                ):
                    raise RuntimeError(
                        "could not download/embed replacement image "
                        f"{action['image_url']!r}"
                    )

            if action.get("width_mm"):
                image.Width = int(action["width_mm"] * 100)
            if action.get("height_mm"):
                image.Height = int(action["height_mm"] * 100)

            anchor_map = {
                "at_paragraph": 0,
                "as_character": 1,
                "at_page": 2,
                "at_character": 4,
            }
            wrap_map = {
                "none": 0,
                "through": 1,
                "parallel": 2,
                "dynamic": 3,
                "behind": 1,
                "in_front": 1,
            }
            if not has_explicit_position:
                image.AnchorType = anchor_map.get(anchor_type_str, 1)
                image.Surround = wrap_map.get(wrap_type_str, 0)
            elif action.get("wrap_type"):
                image.Surround = wrap_map.get(wrap_type_str, 2)

            if action.get("pos_x_mm") is not None:
                image.HoriOrientPosition = int(action["pos_x_mm"] * 100)
                image.HoriOrient = 0
            if action.get("pos_y_mm") is not None:
                image.VertOrientPosition = int(action["pos_y_mm"] * 100)
                image.VertOrient = 0

            if not has_explicit_position and anchor_type_str == "as_character":
                _apply_centered_inline_image_placement(image, "_preview_replace_image")

            _log(
                "_preview_replace_image: updated image at "
                f"index {resolved_index} name={resolved_name}"
            )
            return {
                "type": "image_update",
                "element_index": resolved_index,
                "image_index": resolved_index,
                "image_id": resolved_name,
            }

        _log(
            "_preview_replace_image: image target not found "
            f"(image_id={action.get('image_id')}, image_index={action.get('image_index')}, element_index={action.get('element_index')})"
        )
        return None

    except Exception as e:
        _log(f"_preview_replace_image: failed: {e}")
        return None


def _preview_delete_image(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Delete an image."""
    _log(f"_preview_delete_image: for op {op_id}")

    try:
        image, resolved_index, resolved_name = _resolve_graphic_target(doc, action)
        if image is not None:
            image.dispose()  # Remove from document

            _log(
                "_preview_delete_image: deleted image at "
                f"index {resolved_index} name={resolved_name}"
            )
            return {
                "type": "image_deletion",
                "element_index": resolved_index,
                "image_index": resolved_index,
                "image_id": resolved_name,
            }

        _log(
            "_preview_delete_image: image target not found "
            f"(image_id={action.get('image_id')}, image_index={action.get('image_index')}, element_index={action.get('element_index')})"
        )
        return None

    except Exception as e:
        _log(f"_preview_delete_image: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Table of Contents preview functions (Phase 3)
# ---------------------------------------------------------------------------


def _preview_insert_toc(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a Table of Contents."""
    _log(f"_preview_insert_toc: for op {op_id}")

    toc_spec = _toc_spec_from_action(action)
    title = toc_spec["title"]
    levels = toc_spec["heading_levels"]

    try:
        # Create ContentIndex object
        toc = doc.createInstance("com.sun.star.text.ContentIndex")
        toc.Title = title
        toc.Level = levels
        # Ensure it's not protected so user can edit if needed
        toc.IsProtected = False

        # Set Name to op_id for robust rejection
        toc_name = f"TOC_{op_id}"
        toc.Name = toc_name

        # Locate insertion point
        text_obj = doc.getText()
        cursor = None

        if action.get("after_element_id"):
            cursor = _create_cursor_after_action_target(doc, action)
        elif action.get("before_element_id"):
            cursor = _create_cursor_before_action_target(doc, action)

        if cursor is None:
            # Default TOC placement is the beginning of the document.
            cursor = text_obj.createTextCursor()
            cursor.gotoStart(False)

        # Insert TOC
        text_obj.insertTextContent(cursor, toc, False)
        boundary_mark = _insert_index_trailing_boundary(
            doc, toc, plan_id, op_id, role="boundary_after"
        )

        # Update TOC to generate content
        toc.update()

        _log(f"_preview_insert_toc: inserted TOC '{title}'")

        return {
            "type": "toc_insertion",
            "toc_id": toc_name,  # Return Name for rejection
            "boundary_ids": [boundary_mark] if boundary_mark else [],
            "title": title,
            "toc": toc_spec,
        }

    except Exception as e:
        _log(f"_preview_insert_toc: failed: {e}")
        return None


def _preview_update_toc(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Update an existing Table of Contents."""
    _log(f"_preview_update_toc: for op {op_id}")

    try:
        toc, resolved_index, resolved_name = _resolve_toc_target(doc, action)
        if toc is not None:

            # Check if it supports updating (it should)
            if hasattr(toc, "update"):
                toc.update()
                _log(
                    "_preview_update_toc: updated TOC at "
                    f"index {resolved_index} name={resolved_name}"
                )
                return {
                    "type": "toc_update",
                    "toc_index": resolved_index,
                    "toc_id": resolved_name,
                }
            else:
                _log(
                    f"_preview_update_toc: index {resolved_index} does not support update"
                )
                return None

        _log(
            "_preview_update_toc: TOC target not found "
            f"(toc_id={action.get('toc_id')}, toc_index={action.get('toc_index')}, element_index={action.get('element_index')})"
        )
        return None

    except Exception as e:
        _log(f"_preview_update_toc: failed: {e}")
        return None


def _preview_remove_toc(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Queue removal of an existing Table of Contents."""
    _log(f"_preview_remove_toc: for op {op_id}")

    try:
        toc, resolved_index, resolved_name = _resolve_toc_target(doc, action)
        if toc is None:
            _log(
                "_preview_remove_toc: TOC target not found "
                f"(toc_id={action.get('toc_id')}, toc_index={action.get('toc_index')}, element_index={action.get('element_index')})"
            )
            return None

        return {
            "type": "toc_removal",
            "toc_index": resolved_index,
            "toc_id": resolved_name,
        }

    except Exception as e:
        _log(f"_preview_remove_toc: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Mail Merge preview functions (Phase 3)
# ---------------------------------------------------------------------------


def _preview_insert_merge_field(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a mail merge field (User Field)."""
    _log(f"_preview_insert_merge_field: for op {op_id}")

    field_name = action.get("field_name")
    if not field_name:
        return None

    try:
        # Get or create Master field
        masters = doc.getTextFieldMasters()
        master_name = f"com.sun.star.text.field.User.{field_name}"

        if not masters.hasByName(master_name):
            _log(f"Creating new user field master: {master_name}")
            master = doc.createInstance("com.sun.star.text.fieldmaster.User")
            master.Name = field_name
            master.Content = f"<{field_name}>"  # Default value
        else:
            _log(f"Using existing user field master: {master_name}")

        # Create field instance
        field = doc.createInstance("com.sun.star.text.textfield.User")
        field.Content = (
            f"<{field_name}>"  # This might be ignored if attached to master?
        )
        # Attach to master
        field.attachTextFieldMaster(masters.getByName(master_name))

        # Insert
        paragraph = _locate_paragraph(doc, action)
        text_obj = doc.getText()
        cursor = None

        if paragraph:
            if (
                action.get("after_element_id")
                or action.get("element_index") is not None
            ):
                cursor = text_obj.createTextCursorByRange(paragraph.getEnd())
            elif action.get("before_element_id"):
                cursor = text_obj.createTextCursorByRange(paragraph.getStart())

        if cursor is None:
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)

        text_obj.insertTextContent(cursor, field, False)

        _log(f"_preview_insert_merge_field: inserted field {field_name}")

        return {
            "type": "merge_field_insertion",
            "field_name": field_name,
        }

    except Exception as e:
        _log(f"_preview_insert_merge_field: failed: {e}")
        return None


def _preview_apply_merge_data(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Apply data to merge fields."""
    _log(f"_preview_apply_merge_data: for op {op_id}")

    data = action.get("data", {})
    preview_only = action.get("preview_only", False)

    try:
        masters = doc.getTextFieldMasters()
        updated_fields = 0

        for key, value in data.items():
            master_name = f"com.sun.star.text.field.User.{key}"
            if masters.hasByName(master_name):
                master = masters.getByName(master_name)
                # Update content
                master.Content = str(value)
                updated_fields += 1

        # Refresh fields to show new values
        doc.getTextFields().refresh()

        # If not preview_only, we might want to replace fields with text?
        # For now, let's keep them as fields but with the value set.
        # Replacing with text would require iterating all field instances.

        if not preview_only:
            # Iterate all user fields and replace with fixed text
            text_fields = doc.getTextFields()
            enum = text_fields.createEnumeration()
            while enum.hasMoreElements():
                tf = enum.nextElement()
                if tf.supportsService("com.sun.star.text.textfield.User"):
                    # Check if this field is one we have data for
                    # Master name isn't directly exposed easily on the instance sometimes
                    # But we can check Content if it matches?
                    # Actually, User fields share content via Master.

                    # Safe way: replace all User fields with their current content
                    anchor = tf.getAnchor()
                    content = tf.getTextFieldMaster().Content
                    anchor.setString(content)
                    # field is automatically removed when overwritten?
                    # No, setString replaces the range. The field is essentially the range.

            _log(f"_preview_apply_merge_data: flattened fields to text")

        _log(f"_preview_apply_merge_data: updated {updated_fields} fields")

        return {
            "type": "merge_data_applied",
            "updated_count": updated_fields,
        }

    except Exception as e:
        # Some errors might occur during iteration or replacement
        _log(f"_preview_apply_merge_data: failed: {e}")
        # Don't fail completely if some fields work
        return {"type": "merge_data_error", "error": str(e)}


# ---------------------------------------------------------------------------
# Content Controls preview functions (Phase 3)
# ---------------------------------------------------------------------------


def _preview_insert_control(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a form control (Checkbox, Dropdown, Text)."""
    _log(f"_preview_insert_control: for op {op_id}")

    control_type = action.get("control_type", "text")
    name = action.get("name", f"Control_{op_id}")
    label = action.get("label", name)
    options = action.get("options", [])
    default_text = action.get("default_text", "")
    width_mm = action.get("width_mm", 20.0)
    height_mm = action.get("height_mm", 5.0)

    try:
        # Locate anchor
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
            char_offset = action.get("char_offset", 0)
            if char_offset > 0:
                cursor.goRight(char_offset, False)

        # Create ControlShape
        shape = doc.createInstance("com.sun.star.drawing.ControlShape")
        shape.AnchorType = 1  # AS_CHARACTER
        shape.Width = int(width_mm * 100)
        shape.Height = int(height_mm * 100)

        # Set Name to op_id for robust rejection
        control_name = f"Control_{op_id}"
        shape.Name = control_name

        # Create Control Model
        model_service = {
            "checkbox": "com.sun.star.form.component.CheckBox",
            "dropdown": "com.sun.star.form.component.ComboBox",
            "text": "com.sun.star.form.component.TextField",
        }.get(control_type, "com.sun.star.form.component.TextField")

        model = doc.createInstance(model_service)
        model.Name = name

        if control_type == "checkbox":
            model.Label = label
            model.State = 0  # Unchecked
        elif control_type == "dropdown":
            model.StringItemList = tuple(options)
            model.Dropdown = True
        elif control_type == "text":
            model.Text = default_text

        # Assign model to shape
        shape.Control = model

        # Insert shape into DrawPage?
        # For AS_CHARACTER, we insert it into TEXT.
        text_obj.insertTextContent(cursor, shape, False)

        _log(f"_preview_insert_control: inserted {control_type} '{name}'")

        return {
            "type": "control_insertion",
            "control_type": control_type,
            "control_id": control_name,  # Return Name for rejection
            "name": name,
        }

    except Exception as e:
        _log(f"_preview_insert_control: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Cross-Reference preview functions (Phase 3)
# ---------------------------------------------------------------------------


def _preview_insert_cross_reference(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a cross-reference field."""
    _log(f"_preview_insert_cross_reference: for op {op_id}")

    ref_type = action.get("reference_type")
    target_name = action.get("target_name")
    display_format = action.get("display_format", "text")

    try:
        # Locate anchor
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
            char_offset = action.get("char_offset", 0)
            if char_offset > 0:
                cursor.goRight(char_offset, False)

        # Create GetReference field
        field = doc.createInstance("com.sun.star.text.textfield.GetReference")

        # Map SourceName (target)
        field.SourceName = target_name

        # Map ReferenceFieldSource (approximate mapping)
        # BOOKMARK=2, FOOTNOTE=3, ENDNOTE=4, SEQUENCE=0 (Figures/Tables usually)
        source_map = {
            "bookmark": 2,
            "footnote": 3,
            "endnote": 4,
            "figure": 0,
            "table": 0,
            "heading": 2,  # Often mapped to bookmarks in some contexts or separate source
        }
        field.ReferenceFieldSource = source_map.get(ref_type, 2)

        if ref_type == "figure":
            field.SequenceNumber = 0  # Need to find sequence ID via name? Complex.
            # fallback: assume target_name is the proper SourceName for the sequence if managed right.

        # Map ReferenceFieldPart
        # PAGE=1, TEXT=2, UP_DOWN=3, CHAPTER=4, NUMBER=0?
        # Actually: 0=PAGE, 1=CHAPTER, 2=TEXT, 3=UP_DOWN, 4=PAGE_DESC, 5=CATEGORY_AND_NUMBER
        # Warning: Constants vary. Let's use standard integer guesses or properties.
        # PAGE=1, TEXT=2 confirmed in some docs?
        # com.sun.star.text.ReferenceFieldPart.PAGE = 0
        # com.sun.star.text.ReferenceFieldPart.TEXT = 2
        # com.sun.star.text.ReferenceFieldPart.NUMBER = 3 (numbering)

        part_map = {
            "page_number": 0,  # PAGE
            "text": 2,  # TEXT
            "number": 3,  # NUMBER (e.g. "Figure 1")
            "chapter": 1,  # CHAPTER
        }
        field.ReferenceFieldPart = part_map.get(display_format, 2)

        # Insert field
        text_obj.insertTextContent(cursor, field, False)

        # Need to return an ID for rejection?
        # Fields don't have names usually.
        # We might need to wrap it or rely on location?
        # Or just accept that fields are hard to revert perfectly by ID.
        # But we can try to return the field object's implementation ID?
        # Note: field.dispose() works.

        field_id = str(id(field))

        _log(f"_preview_insert_cross_reference: inserted ref to {target_name}")

        return {
            "type": "cross_reference_insertion",
            "target_name": target_name,
            "field_id": field_id,  # Proxy ID
        }

    except Exception as e:
        _log(f"_preview_insert_cross_reference: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Master Document preview functions (Phase 3)
# ---------------------------------------------------------------------------


def _preview_insert_sub_document(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a sub-document as a linked section."""
    _log(f"_preview_insert_sub_document: for op {op_id}")

    file_path = action.get("file_path")
    link_mode = action.get("link_mode", "linked")

    # Needs absolute URL format for LibreOffice
    import uno

    file_url = uno.systemPathToFileUrl(file_path)

    try:
        # Locate anchor
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            text_obj = doc.getText()
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
        else:
            text_obj = paragraph.getText()
            cursor = text_obj.createTextCursorByRange(paragraph.getStart())
            # Sub-documents insert a section, essentially block-level, but technically inline

        # Create Section
        section_name = f"SubDoc_{op_id}"
        section = doc.createInstance("com.sun.star.text.TextSection")
        section.setName(section_name)

        # Link properties
        section.FileLink.FileURL = file_url

        # Insert section
        text_obj.insertTextContent(cursor, section, False)

        if link_mode == "embedded":
            # Break link to embed
            # section.FileLink.FileURL = "" # Might not be enough?
            # section.setPropertyValue("FileLink", uno.createUnoStruct("com.sun.star.text.SectionFileLink"))
            # Or use insertDocumentFromURL directly if we don't want a section container.
            # But the plan proposed using Sections for preview management.
            # Let's keep it linked for preview and maybe unlink on "Finalize"?
            # Or just set link to empty struct?

            # For now, if "embedded", we just leave it as linked because true embedding
            # involves copying content which the Section link does automatically visually,
            # but "breaking" the link makes it permanent.
            # To break link:
            link = section.FileLink
            link.FileURL = ""
            section.FileLink = link
            pass

        _log(f"_preview_insert_sub_document: inserted section {section_name}")

        return {
            "type": "sub_document_insertion",
            "section_name": section_name,
            "file_path": file_path,
        }

    except Exception as e:
        _log(f"_preview_insert_sub_document: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Multi-Column Layout preview functions (Phase 4)
# ---------------------------------------------------------------------------


def _preview_set_columns(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Set columns for page style or section."""
    _log(f"_preview_set_columns: for op {op_id}")

    col_count = action.get("column_count", 1)
    spacing_mm = action.get("spacing_mm", 5.0)
    target_scope = action.get("target_scope", "page_style")
    section_index = action.get("section_index")
    section_name = action.get("section_name")

    try:
        prev_columns = None
        prev_spacing = None
        prev_separator = None
        target = None
        page_style_name = None

        if target_scope == "page_style":
            # Apply to current page style
            # Locate anchor to find page style at cursor
            paragraph = _locate_paragraph(doc, action)
            if paragraph:
                cursor = paragraph.getText().createTextCursorByRange(
                    paragraph.getStart()
                )
                page_style_name = cursor.PageDescName
                if not page_style_name:
                    page_style_name = "Standard"  # Fallback

                styles = doc.getStyleFamilies().getByName("PageStyles")
                if styles.hasByName(page_style_name):
                    target = styles.getByName(page_style_name)
                    current_columns = _get_text_columns(target)
                    prev_columns = _get_text_columns_property(
                        current_columns, "ColumnCount", 1
                    )
                    prev_spacing = _get_text_columns_property(
                        current_columns, "AutomaticDistance", None
                    )
                    prev_separator = _get_text_columns_property(
                        current_columns, "SeparatorLineIsOn", None
                    )
            else:
                # Default to Standard style if no location
                styles = doc.getStyleFamilies().getByName("PageStyles")
                page_style_name = "Standard"
                target = styles.getByName(page_style_name)
                current_columns = _get_text_columns(target)
                prev_columns = _get_text_columns_property(
                    current_columns, "ColumnCount", 1
                )
                prev_spacing = _get_text_columns_property(
                    current_columns, "AutomaticDistance", None
                )
                prev_separator = _get_text_columns_property(
                    current_columns, "SeparatorLineIsOn", None
                )

        elif target_scope == "section":
            sections = doc.getTextSections()
            if section_name and sections.hasByName(section_name):
                target = sections.getByName(section_name)
            elif section_index is not None and sections.getCount() > section_index:
                target = sections.getByIndex(section_index)
                try:
                    section_name = target.getName()
                except Exception:
                    pass

            if target:
                current_columns = _get_text_columns(target)
                prev_columns = _get_text_columns_property(
                    current_columns, "ColumnCount", 1
                )
                prev_spacing = _get_text_columns_property(
                    current_columns, "AutomaticDistance", None
                )
                prev_separator = _get_text_columns_property(
                    current_columns, "SeparatorLineIsOn", None
                )

        if target:
            current_columns = _get_text_columns(target)
            target.TextColumns = _configure_text_columns(
                current_columns,
                col_count,
                spacing_mm,
                action.get("separator_line"),
            )
            _log(f"_preview_set_columns: set columns to {col_count} on {target_scope}")
        else:
            _log(f"_preview_set_columns: target not found")
            return None

        # Return state for revert
        return {
            "type": "columns_set",
            "target_scope": target_scope,
            "section_name": section_name,
            "page_style_name": page_style_name,
            "prev_column_count": prev_columns,
            "prev_spacing_mm": (
                float(prev_spacing) / 100.0 if prev_spacing is not None else None
            ),
            "prev_separator_line": prev_separator,
        }

    except Exception as e:
        _log(f"_preview_set_columns: failed: {e}")
        return None


def _preview_insert_column_break(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert manual column break."""
    _log(f"_preview_insert_column_break: for op {op_id}")

    try:
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            raise RuntimeError("Paragraph not found for column break")

        # Use BreakType
        import uno
        from com.sun.star.style import BreakType

        # Apply break AFTER the paragraph? Or insert empty paragraph with break?
        # Usually manual break is on a paragraph.
        # BreakType.COLUMN_AFTER = 2?
        # com.sun.star.style.BreakType.COLUMN_BEFORE = 1
        # com.sun.star.style.BreakType.COLUMN_AFTER = 2
        # com.sun.star.style.BreakType.PAGE_BEFORE = 4
        # com.sun.star.style.BreakType.PAGE_AFTER = 5

        # Safe way using constants if available via bridge, or hardcoded
        # Let's trust integers for this environment if enums tricky import
        # COLUMN_BEFORE = 1, COLUMN_AFTER = 2

        prev_break = paragraph.BreakType
        paragraph.BreakType = 2  # COLUMN_AFTER

        _log(f"_preview_insert_column_break: inserted column break")

        return {
            "type": "column_break_insertion",
            "element_id": action.get("element_id"),
            "paragraph_index": action.get("element_index"),
            "prev_break_type": prev_break,
        }

    except Exception as e:
        _log(f"_preview_insert_column_break: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Text Box (Frame) preview functions (Phase 4)
# ---------------------------------------------------------------------------


def _preview_insert_text_box(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a text box (TextFrame)."""
    _log(f"_preview_insert_text_box: for op {op_id}")

    width_mm = action.get("width_mm", 50.0)
    height_mm = action.get("height_mm")
    anchor_type = action.get("anchor_type", "paragraph")
    x_offset_mm = action.get("x_offset_mm", 0.0)
    y_offset_mm = action.get("y_offset_mm", 0.0)
    text_content = action.get("text_content")

    # MM to 1/100MM
    width = int(width_mm * 100)
    height = int(height_mm * 100) if height_mm else 0
    x_pos = int(x_offset_mm * 100)
    y_pos = int(y_offset_mm * 100)

    try:
        # Create Frame
        frame = doc.createInstance("com.sun.star.text.TextFrame")
        name = f"Frame_{op_id}"
        frame.setName(name)

        frame.Width = width
        if height > 0:
            frame.Height = height
            frame.SizeType = 1  # FIX (Fixed) = 1, VAR (Variable) = 0? 1=MIN?
            # actually SizeType: 0=VARIABLE, 1=FIX, 2=MIN?
            # Let's assume defaults mostly work or setting height implies fixed?

        # Mapping Anchor Type
        # AS_CHARACTER = 1, AT_CHARACTER = 4, AT_PARAGRAPH = 0, AT_PAGE = 2
        import uno
        from com.sun.star.text.TextContentAnchorType import (
            AS_CHARACTER,
            AT_CHARACTER,
            AT_PARAGRAPH,
            AT_PAGE,
        )

        anchor_map = {
            "as_character": AS_CHARACTER,
            "at_character": AT_CHARACTER,
            "at_paragraph": AT_PARAGRAPH,
            "at_page": AT_PAGE,
        }
        frame.AnchorType = anchor_map.get(anchor_type, AT_PARAGRAPH)

        # Positioning
        # HoriOrient/VertOrient = 0 (NONE) allows manual positioning
        frame.HoriOrient = 0
        frame.VertOrient = 0
        frame.HoriOrientPosition = x_pos
        frame.VertOrientPosition = y_pos

        # Style
        if action.get("background_color"):
            frame.BackColor = int(action.get("background_color").lstrip("#"), 16)

        # Insert
        paragraph = _locate_paragraph(doc, action)
        if paragraph:
            cursor = paragraph.getText().createTextCursorByRange(paragraph.getStart())
            paragraph.getText().insertTextContent(cursor, frame, False)
        else:
            cursor = _create_cursor_before_action_target(doc, action)
            if cursor is None:
                raise RuntimeError("Could not resolve insertion anchor for text box")
            cursor.getText().insertTextContent(cursor, frame, False)

        # Set Content
        if text_content:
            frame_cursor = frame.getText().createTextCursor()
            frame.getText().insertString(frame_cursor, text_content, False)

        _log(f"_preview_insert_text_box: inserted frame {name}")

        return {
            "type": "text_box_insertion",
            "frame_name": name,
        }

    except Exception as e:
        _log(f"_preview_insert_text_box: failed: {e}")
        return None


def _preview_delete_text_box(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Delete a text box."""
    _log(f"_preview_delete_text_box: for op {op_id}")
    name = action.get("frame_name")

    try:
        frames = doc.getTextFrames()
        if frames.hasByName(name):
            frame = frames.getByName(name)
            frame.dispose()
            _log(f"_preview_delete_text_box: deleted {name}")
            return {"type": "text_box_deletion", "frame_name": name}
        else:
            _log(f"_preview_delete_text_box: frame {name} not found")
            return None

    except Exception as e:
        _log(f"_preview_delete_text_box: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Drop Cap preview functions (Phase 4)
# ---------------------------------------------------------------------------


def _preview_set_drop_cap(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Set drop cap on a paragraph."""
    _log(f"_preview_set_drop_cap: for op {op_id}")

    char_count = action.get("char_count", 1)
    lines = action.get("lines", 3)
    distance_mm = action.get("distance_mm", 1.0)

    # Convert mm to 1/100mm
    distance = int(distance_mm * 100)

    try:
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            raise RuntimeError("Target paragraph not found for drop cap")

        # Store previous values for revert
        prev_char_count = paragraph.DropCapCharCount
        prev_lines = paragraph.DropCapLines
        prev_distance = paragraph.DropCapDistance

        # Apply drop cap
        paragraph.DropCapCharCount = char_count
        paragraph.DropCapLines = lines
        paragraph.DropCapDistance = distance

        _log(f"_preview_set_drop_cap: set drop cap ({char_count} chars, {lines} lines)")

        return {
            "type": "drop_cap_set",
            "prev_char_count": prev_char_count,
            "prev_lines": prev_lines,
            "prev_distance": prev_distance,
            "element_index": action.get("element_index"),
            "element_id": action.get("element_id"),
        }

    except Exception as e:
        _log(f"_preview_set_drop_cap: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Shape preview functions (Phase 4)
# ---------------------------------------------------------------------------


def _preview_insert_shape(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert a shape with optional text."""
    _log(f"_preview_insert_shape: for op {op_id}")

    shape_type = action.get("shape_type", "rectangle")
    width_mm = action.get("width_mm", 50.0)
    height_mm = action.get("height_mm", 30.0)
    x_offset_mm = action.get("x_offset_mm", 0.0)
    y_offset_mm = action.get("y_offset_mm", 0.0)

    # Convert mm to 1/100mm
    width = int(width_mm * 100)
    height = int(height_mm * 100)
    x_pos = int(x_offset_mm * 100)
    y_pos = int(y_offset_mm * 100)

    try:
        # Shape service mapping
        shape_map = {
            "rectangle": "com.sun.star.drawing.RectangleShape",
            "rounded_rectangle": "com.sun.star.drawing.RectangleShape",
            "oval": "com.sun.star.drawing.EllipseShape",
            "callout": "com.sun.star.drawing.CaptionShape",
        }
        service_name = shape_map.get(shape_type, "com.sun.star.drawing.RectangleShape")

        shape = doc.createInstance(service_name)

        # Size and Position
        import uno

        size = uno.createUnoStruct("com.sun.star.awt.Size")
        size.Width = width
        size.Height = height
        shape.Size = size

        position = uno.createUnoStruct("com.sun.star.awt.Point")
        position.X = x_pos
        position.Y = y_pos
        shape.Position = position

        # Colors
        if action.get("fill_color"):
            shape.FillColor = int(action["fill_color"].lstrip("#"), 16)
        if action.get("line_color"):
            shape.LineColor = int(action["line_color"].lstrip("#"), 16)

        # Rounded corners for rounded_rectangle
        if shape_type == "rounded_rectangle":
            try:
                shape.CornerRadius = 500  # 5mm corner radius
            except Exception:
                pass

        # Add to draw page
        draw_page = doc.getDrawPage()
        draw_page.add(shape)

        # Set Name
        name = f"Shape_{op_id}"
        shape.Name = name

        # Text content
        if action.get("text_content"):
            shape.getString()  # Initialize text
            shape.setString(action["text_content"])

        _log(f"_preview_insert_shape: inserted {shape_type} as {name}")

        return {
            "type": "shape_insertion",
            "shape_name": name,
        }

    except Exception as e:
        _log(f"_preview_insert_shape: failed: {e}")
        return None


def _preview_link_text_frames(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Link two text frames."""
    _log(f"_preview_link_text_frames: for op {op_id}")

    source_name = action.get("source_frame_name")
    target_name = action.get("target_frame_name")

    try:
        frames = doc.getTextFrames()

        if not frames.hasByName(source_name):
            raise RuntimeError(f"Source frame '{source_name}' not found")
        if not frames.hasByName(target_name):
            raise RuntimeError(f"Target frame '{target_name}' not found")

        source = frames.getByName(source_name)
        target = frames.getByName(target_name)

        # Store previous chain for revert
        prev_chain = source.ChainNextName

        # Link
        source.ChainNextName = target.Name

        _log(f"_preview_link_text_frames: linked {source_name} -> {target_name}")

        return {
            "type": "frames_linked",
            "source_frame_name": source_name,
            "target_frame_name": target_name,
            "prev_chain_name": prev_chain,
        }

    except Exception as e:
        _log(f"_preview_link_text_frames: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Phase 5: Lists, Watermarks, Equations, Comparison
# ---------------------------------------------------------------------------


def _preview_set_list_style(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Set list style on paragraph."""
    _log(f"_preview_set_list_style: for op {op_id}")

    list_type = action.get("list_type", "bullet")
    bullet_char = action.get("bullet_char")
    number_format = action.get("number_format", "1")
    start_value = action.get("start_value", 1)
    indent_mm = action.get("indent_mm", 6.0)

    try:
        paragraph = _locate_paragraph(doc, action)
        if paragraph is None:
            raise RuntimeError("Paragraph not found for list style")

        # Store previous state
        prev_numbering = paragraph.NumberingRules
        prev_level = paragraph.NumberingLevel

        if list_type == "none":
            paragraph.NumberingRules = None
        else:
            # Create or modify numbering rules
            numbering = doc.createInstance("com.sun.star.text.NumberingRules")

            # Get level 0 properties
            props = list(numbering.getByIndex(0))

            if list_type == "bullet":
                # Set bullet character
                for prop in props:
                    if prop.Name == "NumberingType":
                        prop.Value = 6  # CHAR_SPECIAL
                    if prop.Name == "BulletChar":
                        prop.Value = bullet_char or "•"
            else:  # numbered
                format_map = {"1": 4, "a": 0, "A": 1, "i": 2, "I": 3}
                for prop in props:
                    if prop.Name == "NumberingType":
                        prop.Value = format_map.get(number_format, 4)
                    if prop.Name == "StartWith":
                        prop.Value = start_value

            # Set indent
            indent = int(indent_mm * 100)
            for prop in props:
                if prop.Name == "LeftMargin":
                    prop.Value = indent

            numbering.replaceByIndex(0, tuple(props))
            paragraph.NumberingRules = numbering
            paragraph.NumberingLevel = 0

        _log(f"_preview_set_list_style: applied {list_type} list")

        return {
            "type": "list_style_set",
            "prev_numbering": prev_numbering,
            "prev_level": prev_level,
            "element_index": action.get("element_index"),
        }

    except Exception as e:
        _log(f"_preview_set_list_style: failed: {e}")
        return None


def _preview_insert_watermark(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert text or image watermark."""
    _log(f"_preview_insert_watermark: for op {op_id}")

    watermark_type = action.get("watermark_type", "text")
    text = action.get("text", "DRAFT")
    font_size_pt = action.get("font_size_pt", 72.0)
    color = action.get("color", "#C0C0C0")
    rotation = action.get("rotation_degrees", -45.0)
    opacity = action.get("opacity", 0.5)

    try:
        import uno

        # Create shape for watermark in header (so it appears on all pages)
        page_styles = doc.getStyleFamilies().getByName("PageStyles")
        page_style = page_styles.getByName("Standard")

        # Enable header if not already
        if not page_style.HeaderIsOn:
            page_style.HeaderIsOn = True

        header_text = page_style.HeaderText

        if watermark_type == "text":
            # Create text shape
            shape = doc.createInstance("com.sun.star.drawing.TextShape")
            shape.String = text

            # Size - large to cover page
            size = uno.createUnoStruct("com.sun.star.awt.Size")
            size.Width = 15000  # 150mm
            size.Height = 5000  # 50mm
            shape.Size = size

            # Position - center of page
            position = uno.createUnoStruct("com.sun.star.awt.Point")
            position.X = 3000
            position.Y = 10000
            shape.Position = position

            # Styling
            shape.CharHeight = font_size_pt
            shape.CharColor = int(color.lstrip("#"), 16)
            shape.RotateAngle = int(rotation * 100)  # 1/100 degrees

            # Transparency
            shape.FillTransparence = int((1 - opacity) * 100)

            # Add to draw page
            draw_page = doc.getDrawPage()
            draw_page.add(shape)

            # Name for removal
            name = f"Watermark_{op_id}"
            shape.Name = name

        else:
            # Image watermark - similar approach with GraphicObject
            name = f"Watermark_{op_id}"
            # Simplified - would need image loading
            _log("Image watermark not fully implemented")

        _log(f"_preview_insert_watermark: inserted {watermark_type} watermark")

        return {
            "type": "watermark_insertion",
            "watermark_name": name,
        }

    except Exception as e:
        _log(f"_preview_insert_watermark: failed: {e}")
        return None


def _preview_insert_equation(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Insert math equation."""
    _log(f"_preview_insert_equation: for op {op_id}")

    formula = action.get("formula", "a^2 + b^2 = c^2")
    formula_type = action.get("formula_type", "starmath")

    try:
        # Create embedded math object
        math_obj = doc.createInstance("com.sun.star.text.TextEmbeddedObject")
        math_obj.CLSID = "078B7ABA-54FC-457F-8551-6147e776a997"  # StarMath CLSID

        # Insert at position
        paragraph = _locate_paragraph(doc, action)
        if paragraph:
            cursor = paragraph.getText().createTextCursorByRange(paragraph.getEnd())
            paragraph.getText().insertTextContent(cursor, math_obj, False)
        else:
            text = doc.getText()
            text.insertTextContent(text.getEnd(), math_obj, False)

        # Access the embedded math component and set formula
        if math_obj.EmbeddedObject:
            model = math_obj.EmbeddedObject
            model.setPropertyValue("Formula", formula)

        name = f"Equation_{op_id}"
        math_obj.Name = name

        _log(f"_preview_insert_equation: inserted equation")

        return {
            "type": "equation_insertion",
            "equation_name": name,
        }

    except Exception as e:
        _log(f"_preview_insert_equation: failed: {e}")
        return None


def _preview_compare_documents(
    doc, action: Dict[str, Any], plan_id: str, op_id: str
) -> Optional[Dict[str, Any]]:
    """Compare two documents."""
    _log(f"_preview_compare_documents: for op {op_id}")

    original_path = action.get("original_path")
    modified_path = action.get("modified_path")

    try:
        import uno
        from com.sun.star.beans import PropertyValue

        # Open original document
        desktop = XSCRIPTCONTEXT.getDesktop() if "XSCRIPTCONTEXT" in dir() else None
        if not desktop:
            raise RuntimeError("Desktop not available")

        # Load properties
        props = (PropertyValue("Hidden", 0, True, 0),)

        original_url = uno.systemPathToFileUrl(original_path)
        modified_url = uno.systemPathToFileUrl(modified_path)

        # Use document's compareDocuments method
        doc.loadVersionFromFile(modified_url)

        # Track changes will show differences
        doc.RecordChanges = True

        _log(f"_preview_compare_documents: comparison initiated")

        return {
            "type": "document_comparison",
            "original_path": original_path,
            "modified_path": modified_path,
        }

    except Exception as e:
        _log(f"_preview_compare_documents: failed: {e}")
        return None


# ---------------------------------------------------------------------------
# Revert/Reject Helpers (Phase 2/3)
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


def _resolve_graphic_target(doc, action: Dict[str, Any]):
    """Resolve a graphic object by preview name, image_index, or legacy element_index."""
    try:
        graphics = doc.getGraphicObjects()
    except Exception as exc:
        _log(f"_resolve_graphic_target: failed to access graphic objects: {exc}")
        return None, None, None

    image_id = action.get("image_id")
    if image_id:
        try:
            if graphics.hasByName(image_id):
                graphic = graphics.getByName(image_id)
                resolved_index = None
                count = graphics.getCount()
                for idx in range(count):
                    try:
                        candidate = graphics.getByIndex(idx)
                        if candidate == graphic:
                            resolved_index = idx
                            break
                    except Exception:
                        continue
                return graphic, resolved_index, image_id
        except Exception as exc:
            _log(f"_resolve_graphic_target: name lookup failed for {image_id}: {exc}")

    for index_key in ("image_index", "element_index"):
        raw_index = action.get(index_key)
        if raw_index is None:
            continue
        try:
            graphic_index = int(raw_index)
        except Exception:
            continue
        try:
            if 0 <= graphic_index < graphics.getCount():
                graphic = graphics.getByIndex(graphic_index)
                return graphic, graphic_index, getattr(graphic, "Name", None)
        except Exception as exc:
            _log(
                f"_resolve_graphic_target: {index_key} lookup failed for {raw_index}: {exc}"
            )
    return None, None, None


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


def _toc_spec_from_action(action: Dict[str, Any]) -> Dict[str, Any]:
    raw = action.get("toc") if isinstance(action.get("toc"), dict) else {}
    title = raw.get("title") or action.get("title") or "Table of Contents"
    levels_raw = (
        raw.get("heading_levels")
        or raw.get("levels")
        or action.get("heading_levels")
        or action.get("levels")
        or 3
    )
    try:
        heading_levels = max(1, min(10, int(levels_raw)))
    except Exception:
        heading_levels = 3

    return {
        "title": str(title),
        "heading_levels": heading_levels,
        "show_page_numbers": bool(raw.get("show_page_numbers", True)),
        "right_align_page_numbers": bool(raw.get("right_align_page_numbers", True)),
        "use_hyperlinks": bool(raw.get("use_hyperlinks", True)),
        "tab_leader": str(raw.get("tab_leader") or "dot"),
    }


def _resolve_toc_target(doc, action: Dict[str, Any]):
    try:
        indexes = doc.getDocumentIndexes()
        count = indexes.getCount()
    except Exception:
        return None, None, None

    toc_id = action.get("toc_id")
    if isinstance(toc_id, str) and toc_id:
        for i in range(count):
            try:
                idx = indexes.getByIndex(i)
                if hasattr(idx, "Name") and idx.Name == toc_id:
                    return idx, i, toc_id
            except Exception:
                continue

    raw_index = action.get("toc_index")
    if raw_index is None:
        raw_index = action.get("element_index")
    try:
        toc_index = int(raw_index if raw_index is not None else 0)
    except Exception:
        toc_index = 0

    if 0 <= toc_index < count:
        try:
            idx = indexes.getByIndex(toc_index)
            name = getattr(idx, "Name", None)
            return idx, toc_index, name
        except Exception:
            return None, None, None

    return None, None, None


# ---------------------------------------------------------------------------
# Shared helpers (extraction, insertion, logging)
# ---------------------------------------------------------------------------


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


def _resolve_target_bookmark(action: Dict[str, Any]) -> Optional[str]:
    element_id = action.get("element_id") or action.get("target_element_id")
    if element_id:
        return element_id
    after_id = action.get("after_element_id")
    if after_id:
        return after_id
    before_id = action.get("before_element_id")
    if before_id:
        return before_id
    return None


def _is_sdoc_bookmark_name(bookmark_name: Optional[str]) -> bool:
    return bool(
        isinstance(bookmark_name, str) and SDOC_BOOKMARK_RE.match(bookmark_name)
    )


def _collect_required_sdoc_bookmarks(actions: List[Dict[str, Any]]) -> List[str]:
    required: Set[str] = set()
    for action in actions:
        if not isinstance(action, dict):
            continue
        for key in (
            "element_id",
            "target_element_id",
            "after_element_id",
            "before_element_id",
        ):
            bookmark_name = action.get(key)
            if _is_sdoc_bookmark_name(bookmark_name):
                required.add(bookmark_name)
    return sorted(required)


def _iter_body_elements(doc) -> List[Tuple[str, Any]]:
    elements: List[Tuple[str, Any]] = []
    enum = doc.getText().createEnumeration()
    while enum.hasMoreElements():
        element = enum.nextElement()
        try:
            if element.supportsService("com.sun.star.text.TextTable"):
                elements.append(("table", element))
            elif element.supportsService("com.sun.star.text.Paragraph"):
                elements.append(("paragraph", element))
        except UNOException:
            continue
    return elements


def _auto_chain_insert_actions(actions: List[Dict[str, Any]]) -> None:
    """Fix consecutive insert actions that share the same after_element_id.

    When two or more inserts target the same after_element_id, the second
    insert lands at the same anchor as the first, producing reversed order.
    This function detects the pattern and chains them: action N+1 gets
    after_element_id set to action N's element_id.

    Also ensures every insert action has a root element_id so chaining
    and bookmarking can work.
    """
    insert_kinds = {"insert_paragraph", "insert_table"}

    for action in actions:
        if action.get("kind") not in insert_kinds:
            continue
        if not action.get("element_id"):
            spec = action.get("paragraph") or action.get("table") or {}
            inner_eid = spec.get("element_id")
            if inner_eid:
                action["element_id"] = f"auto_{inner_eid}"
            else:
                action["element_id"] = f"auto_{uuid.uuid4().hex[:10]}"
            _log(
                f"_auto_chain_insert_actions: assigned element_id "
                f"'{action['element_id']}' to insert action missing one"
            )

    prev_action = None
    for action in actions:
        if action.get("kind") not in insert_kinds:
            prev_action = None
            continue

        if prev_action is not None:
            cur_after = action.get("after_element_id")
            prev_after = prev_action.get("after_element_id")
            prev_eid = prev_action.get("element_id")

            if (
                cur_after
                and prev_after
                and cur_after == prev_after
                and prev_eid
            ):
                _log(
                    f"_auto_chain_insert_actions: chaining action "
                    f"after_element_id '{cur_after}' -> '{prev_eid}' "
                    f"(was duplicate of previous action)"
                )
                action["after_element_id"] = prev_eid

        prev_action = action


def _ensure_required_sdoc_bookmarks(doc, actions: List[Dict[str, Any]]) -> None:
    required = _collect_required_sdoc_bookmarks(actions)
    if not required:
        return

    try:
        bookmarks = doc.getBookmarks()
    except Exception as exc:
        _log(f"_ensure_required_sdoc_bookmarks: failed to get bookmarks: {exc}")
        return

    missing = [name for name in required if not bookmarks.hasByName(name)]
    if not missing:
        return

    _log(
        "_ensure_required_sdoc_bookmarks: "
        f"required={len(required)} missing={len(missing)} "
        "cannot synthesize missing sdoc IDs from bookmark names; "
        "will rely on explicit action anchors (element_id/after/before) "
        "or explicit element_index if provided."
    )


def _locate_paragraph(doc, action: Dict[str, Any]):
    element_info = _locate_element(doc, action)
    if element_info and element_info[0] == "paragraph":
        return element_info[1]
    return None


def _locate_element(doc, action: Dict[str, Any]) -> Optional[Tuple[str, Any]]:
    kind = action.get("kind")
    table_target_kinds = {
        "replace_table",
        "delete_table",
        "merge_cells",
        "split_cell",
    }
    # delete_element is generic — it targets both paragraphs and tables.
    # Always prefer_table for it so table bookmarks resolve to the table,
    # not to a paragraph inside a cell. If the bookmark is on a paragraph
    # (not inside a table), the fast-path enumeration returns it as paragraph anyway.
    prefer_table = kind in table_target_kinds or kind == "delete_element"

    bookmark = _resolve_target_bookmark(action)
    if bookmark:
        element = _get_element_by_bookmark(doc, bookmark, prefer_table=prefer_table)
        if element:
            return element

    index = action.get("element_index")
    if isinstance(index, int):
        return _get_element_by_index(doc, index)
    return None


def _create_cursor_at_section_end(doc, section_name: Optional[str]):
    if not isinstance(section_name, str) or not section_name.strip():
        return None
    try:
        sections = doc.getTextSections()
        if not sections.hasByName(section_name):
            return None
        section = sections.getByName(section_name)
        anchor = section.getAnchor()
        text = anchor.getText() or doc.getText()
        cursor = text.createTextCursorByRange(anchor.getEnd())
        cursor.collapseToEnd()
        return cursor
    except Exception as exc:
        _log(f"_create_cursor_at_section_end: failed for {section_name}: {exc}")
        return None


def _create_cursor_before_action_target(doc, action: Dict[str, Any]):
    element = _locate_element(doc, action)
    if element is None:
        section_cursor = _create_cursor_at_section_end(doc, action.get("section_name"))
        if section_cursor is not None:
            return section_cursor
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        return cursor

    kind, obj = element
    if kind == "paragraph":
        text = obj.getText()
        cursor = text.createTextCursorByRange(obj.getStart())
        cursor.collapseToStart()
        return cursor

    getter = getattr(obj, "getAnchor", None)
    if getter:
        anchor = getter()
        text = anchor.getText()
        cursor = text.createTextCursorByRange(anchor)
        cursor.collapseToStart()
        return cursor
    return None


def _create_cursor_after_action_target(doc, action: Dict[str, Any]):
    element = _locate_element(doc, action)
    if element is None:
        section_cursor = _create_cursor_at_section_end(doc, action.get("section_name"))
        if section_cursor is not None:
            return section_cursor
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        return cursor

    kind, obj = element
    if kind == "paragraph":
        text = obj.getText()
        cursor = text.createTextCursorByRange(obj.getStart())
        cursor.gotoEndOfParagraph(False)
        if not cursor.goRight(1, False):
            cursor.gotoEnd(False)
        _log(f"_create_cursor_after_action_target: positioned after paragraph")
        return cursor

    if kind == "table":
        cursor = _create_cursor_after_table(doc, obj)
        if cursor is not None:
            _log(f"_create_cursor_after_action_target: positioned after table")
            return cursor

    getter = getattr(obj, "getAnchor", None)
    if getter:
        anchor = getter()
        text = anchor.getText()
        cursor = text.createTextCursorByRange(anchor.getEnd())
        cursor.collapseToEnd()
        return cursor
    return None


def _create_cursor_after_table(doc, table):
    """Create a text cursor positioned just after the given table.

    Tables whose bookmarks live inside a cell have anchors whose getText()
    returns None or belongs to the cell text, not the document body text.
    In those cases createTextCursorByRange(anchor.getEnd()) will fail.
    We fall back to enumerating body elements to find the paragraph that
    follows the target table and place the cursor at its start.
    """
    # Fast path: try the anchor directly
    anchor = table.getAnchor()
    text = anchor.getText()
    if text is not None:
        try:
            cursor = text.createTextCursorByRange(anchor.getEnd())
            cursor.collapseToStart()
            return cursor
        except Exception:
            pass

    # Fast path 2: try with document body text
    doc_text = doc.getText()
    try:
        cursor = doc_text.createTextCursorByRange(anchor.getEnd())
        cursor.collapseToStart()
        return cursor
    except Exception:
        pass

    # Fallback: enumerate body elements and find the element after this table
    _log("_create_cursor_after_table: anchor-based cursor failed, falling back "
         "to body enumeration")
    target_name = None
    try:
        target_name = table.getName()
    except Exception:
        pass

    found_table = False
    enum = doc_text.createEnumeration()
    while enum.hasMoreElements():
        element = enum.nextElement()
        try:
            if element.supportsService("com.sun.star.text.TextTable"):
                try:
                    if target_name and element.getName() == target_name:
                        found_table = True
                        continue
                except Exception:
                    pass
            if found_table:
                # This is the first element after the target table
                el_text = element.getText() or doc_text
                try:
                    cursor = el_text.createTextCursorByRange(element.getStart())
                    cursor.collapseToStart()
                    return cursor
                except Exception:
                    pass
        except Exception:
            continue

    # Last resort: cursor at the end of the document
    _log("_create_cursor_after_table: all strategies exhausted, using end of document")
    cursor = doc_text.createTextCursor()
    cursor.gotoEnd(False)
    return cursor


def _apply_char_style(paragraph, color: int, strike: bool = False) -> None:
    cursor = paragraph.getText().createTextCursorByRange(paragraph.getStart())
    cursor.gotoRange(paragraph.getEnd(), True)
    
    # Use direct assignments but wrap in safe calls if needed
    try:
        cursor.CharColor = color
        cursor.CharUnderline = FONT_UNDERLINE_SINGLE if strike else 0
        cursor.CharStrikeout = 2 if strike else 0
    except Exception as e:
        _log(f"_apply_char_style failed: {e}")


def _apply_style_preview(paragraph, update: Dict[str, Any]) -> None:
    cursor = paragraph.getText().createTextCursorByRange(paragraph.getStart())
    cursor.gotoRange(paragraph.getEnd(), True)
    if update.get("alignment"):
        align = update["alignment"].lower()
        if align == "left":
            cursor.ParaAdjust = PAR_ADJUST_LEFT
        elif align == "right":
            cursor.ParaAdjust = PAR_ADJUST_RIGHT
        elif align == "center":
            cursor.ParaAdjust = PAR_ADJUST_CENTER
        elif align == "justify":
            cursor.ParaAdjust = PAR_ADJUST_BLOCK
    if update.get("character_format"):
        char = update["character_format"]
        if char.get("font_color_rgb"):
            try:
                cursor.CharColor = int(char["font_color_rgb"].lstrip("#"), 16)
            except Exception:
                pass

    # Phase 4.3: Background Color
    if update.get("background_color"):
        try:
            cursor.ParaBackColor = int(update["background_color"].lstrip("#"), 16)
        except Exception:
            pass

    # Phase 4.3: Borders
    if update.get("border_color") or update.get("border_width_pt"):
        try:
            import uno

            border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
            if update.get("border_color"):
                border.Color = int(update["border_color"].lstrip("#"), 16)
            if update.get("border_width_pt"):
                # LineWidth is in 1/100mm; 1pt = 0.3528 mm = 35.28 1/100mm
                border.LineWidth = int(update["border_width_pt"] * 35.28)
            else:
                border.LineWidth = 35  # Default ~1pt

            # Style mapping (approximate)
            style_map = {"solid": 0, "double": 3, "dotted": 2, "dashed": 1, "none": 0}
            border_style_val = style_map.get(update.get("border_style", "solid"), 0)
            border.LineStyle = border_style_val

            cursor.TopBorder = border
            cursor.BottomBorder = border
            cursor.LeftBorder = border
            cursor.RightBorder = border
        except Exception:
            pass

    # Phase 4.3: Padding (BorderDistance)
    if update.get("padding_left_mm") is not None:
        cursor.LeftBorderDistance = int(update["padding_left_mm"] * 100)
    if update.get("padding_right_mm") is not None:
        cursor.RightBorderDistance = int(update["padding_right_mm"] * 100)
    if update.get("padding_top_mm") is not None:
        cursor.TopBorderDistance = int(update["padding_top_mm"] * 100)
    if update.get("padding_bottom_mm") is not None:
        cursor.BottomBorderDistance = int(update["padding_bottom_mm"] * 100)


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
            if isinstance(runs_spec, list):
                if _spec_has_explicit_run_color({"runs": runs_spec}):
                    return True
    return False


def _ensure_paragraph_boundary_for_insertion(cursor):
    """
    Reuse an existing empty paragraph when possible; otherwise split the
    current paragraph so inserted preview content does not glue onto adjacent
    text.
    """
    if cursor is None:
        return None

    text = cursor.getText()
    if text is None:
        return cursor

    working = text.createTextCursorByRange(cursor)
    working.collapseToStart()

    try:
        probe = text.createTextCursorByRange(working)
        probe.gotoStartOfParagraph(False)
        paragraph_start = probe.getStart()
        probe.gotoEndOfParagraph(True)
        paragraph_text = probe.getString() or ""
        if paragraph_text.strip() == "":
            anchored = text.createTextCursorByRange(paragraph_start)
            anchored.collapseToStart()
            _log("_ensure_paragraph_boundary_for_insertion: reusing empty paragraph")
            return anchored
    except Exception as exc:
        _log(
            "_ensure_paragraph_boundary_for_insertion: paragraph probe failed: "
            f"{exc}"
        )

    try:
        text.insertControlCharacter(working, PARAGRAPH_BREAK, False)
        working.goLeft(1, False)
        _log(
            "_ensure_paragraph_boundary_for_insertion: inserted paragraph break "
            "to isolate preview paragraph"
        )
    except Exception as exc:
        _log(
            "_ensure_paragraph_boundary_for_insertion: failed to insert break: "
            f"{exc}"
        )

    return working


def _insert_paragraphs_from_spec(
    doc,
    paragraph,
    spec: Dict[str, Any],
    plan_id: str,
    op_id: str,
    *,
    add_separator: bool = False,
    append_break: bool = True,
) -> List[str]:
    text = paragraph.getText()
    anchor_cursor = text.createTextCursorByRange(paragraph.getEnd())
    anchor_cursor.collapseToStart()
    inserted_ids: List[str] = []

    if add_separator:
        inserted_ids.append(
            _insert_separator(doc, anchor_cursor, plan_id, op_id, role="separator")
        )

    inserted_ranges = _write_paragraph_spec(
        anchor_cursor,
        spec,
        color=GREEN_RGB,
        append_break=append_break,
        preserve_explicit_run_colors=True,
    )
    inserted_ids.extend(
        _tag_preview_range(doc, entry, plan_id, op_id, role=f"proposed_{idx}")
        for idx, entry in enumerate(inserted_ranges)
    )

    return inserted_ids


def _write_spec_after(
    doc,
    paragraph,
    spec: Dict[str, Any],
    plan_id: str,
    op_id: str,
    *,
    add_separator: bool = False,
) -> List[str]:
    return _insert_paragraphs_from_spec(
        doc,
        paragraph,
        spec,
        plan_id,
        op_id,
        add_separator=add_separator,
    )


def _write_paragraph_spec(
    cursor,
    spec: Dict[str, Any],
    color: Optional[int] = None,
    append_break: bool = True,
    preserve_explicit_run_colors: bool = False,
) -> List[Tuple[Any, Any, Any]]:
    text = cursor.getText()
    ranges: List[Tuple[Any, Any, Any]] = []
    run_segments: List[Tuple[Any, Any, Dict[str, Any]]] = []

    # If a preview color is provided, apply paragraph formatting first and
    # then apply the preview color once at the end to avoid style overrides.
    apply_preview_color = color is not None

    # Extract paragraph format spec (contains style_name, alignment, etc.)
    format_spec = spec.get("format") or {}

    runs = spec.get("runs")
    if runs:
        _log(
            f"_write_paragraph_spec: processing {len(runs)} runs. First few text: {[r.get('text', '')[:20] for r in runs[:3]]}"
        )
    has_explicit_run_color = _spec_has_explicit_run_color(spec)
    full_text = spec.get("text")

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
            image_spec = run.get("image")
            if image_spec:
                anchor = _insert_image_run(new_cursor, image_spec)
                if anchor:
                    ranges.append(anchor)
                continue

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
            if format_spec.get("style_name") in {"Title", "Subtitle"} and run_format.get("font_size_pt"):
                run_format = dict(run_format)
                run_format.pop("font_size_pt", None)
            _apply_run_formatting(
                run_cursor,
                run_format,
                None if apply_preview_color else color,
            )
            run_end = run_cursor.getEnd()
            run_segments.append((run_start, run_end, run_format))
            
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
                        _log(f"_write_paragraph_spec: applied hyperlink to run: {target_url[:50]} (external={is_external})")
                    except Exception as e:
                        _log(f"_write_paragraph_spec: failed to apply hyperlink: {e}")

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
        if color is not None and not apply_preview_color:
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

    # Apply paragraph formatting (style_name, alignment, etc.) to all inserted paragraphs
    if ranges and format_spec:
        for text_obj, start_range, end_range in ranges:
            try:
                para_cursor = text_obj.createTextCursorByRange(start_range)
                para_cursor.gotoRange(end_range, True)
                _apply_paragraph_formatting(para_cursor, format_spec)
            except Exception as e:
                _log(f"_write_paragraph_spec: failed to apply paragraph formatting: {e}")

    # Paragraph style application can reset run-level text attributes.
    # Re-apply run formatting so explicit run directives (italic/bold/underline)
    # survive style assignment.
    if run_segments:
        for run_start, run_end, run_format in run_segments:
            try:
                run_cursor = text.createTextCursorByRange(run_start)
                run_cursor.gotoRange(run_end, True)
                _apply_run_formatting(
                    run_cursor,
                    run_format,
                    None if apply_preview_color else color,
                )
            except Exception as e:
                _log(f"_write_paragraph_spec: failed to re-apply run formatting: {e}")

    # Apply preview color AFTER paragraph formatting to avoid style overrides
    if apply_preview_color and ranges:
        skip_preview_color = preserve_explicit_run_colors and has_explicit_run_color
        if skip_preview_color:
            return ranges
        for text_obj, start_range, end_range in ranges:
            try:
                color_cursor = text_obj.createTextCursorByRange(start_range)
                color_cursor.gotoRange(end_range, True)
                color_cursor.CharColor = color
            except Exception as e:
                _log(f"_write_paragraph_spec: failed to apply preview color: {e}")

    return ranges


def _download_image_to_temp(image_url: str) -> Optional[str]:
    if not image_url:
        return None
    suffix = os.path.splitext(urlparse(image_url).path or "")[1] or ".png"
    try:
        _log(f"_download_image_to_temp: downloading {image_url} → suffix {suffix}")
        
        # Add User-Agent to avoid being blocked by some servers
        req = urllib.request.Request(
            image_url,
            headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) SmartDocs/1.0"}
        )
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as handle:
            with urllib.request.urlopen(req, timeout=30, context=ctx) as response:
                handle.write(response.read())
            temp_name = handle.name
        _log(f"_download_image_to_temp: saved to {temp_name}")
        return temp_name
    except Exception as exc:
        _log(f"_download_image_to_temp: failed for {image_url}: {exc}")
        return None


def _embed_image_url_into_graphic(graphic_obj, image_url: str, log_prefix: str) -> bool:
    """Embed image_url into an existing Writer graphic object.

    Do not assign remote HTTP(S) URLs directly to GraphicURL here. Collabora/LO
    can block remote hosts asynchronously, which makes the preview report OK
    even though the rendered image later fails. Download first, then assign the
    loaded Graphic object (preferred) or a local file URL fallback.
    """
    if not image_url:
        return False

    parsed = urlparse(image_url)
    temp_path: Optional[str] = None
    local_url: Optional[str] = None
    keep_temp_for_graphic_url = False

    try:
        if parsed.scheme in ("http", "https"):
            temp_path = _download_image_to_temp(image_url)
            if not temp_path:
                _log(f"{log_prefix}: failed to download replacement image {image_url}")
                return False
            local_url = Path(temp_path).as_uri()
        elif parsed.scheme == "file":
            local_url = image_url
        elif os.path.exists(image_url):
            local_url = Path(image_url).resolve().as_uri()
        else:
            _log(f"{log_prefix}: unsupported replacement image URL {image_url}")
            return False

        graphic_provider = XSCRIPTCONTEXT.getComponentContext().ServiceManager.createInstanceWithContext( # type: ignore # NOQA
            "com.sun.star.graphic.GraphicProvider", XSCRIPTCONTEXT.getComponentContext() # type: ignore # NOQA
        )
        from com.sun.star.beans import PropertyValue

        props = (PropertyValue(Name="URL", Value=local_url),)
        graphic = graphic_provider.queryGraphic(props)
        if graphic:
            graphic_obj.Graphic = graphic
            _log(f"{log_prefix}: embedded replacement graphic from {image_url}")
            return True

        # Local file URL fallback is still safe: LO reads a local temp/file path,
        # not the remote origin that Collabora may deny asynchronously.
        graphic_obj.GraphicURL = local_url
        keep_temp_for_graphic_url = bool(temp_path)
        _log(f"{log_prefix}: used local GraphicURL fallback for {image_url}")
        return True
    except Exception as exc:
        _log(f"{log_prefix}: failed to embed replacement image {image_url}: {exc}")
        return False
    finally:
        if temp_path and not keep_temp_for_graphic_url and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def _insert_image_run(cursor, image_spec: Dict[str, Any]):
    try:
        doc = XSCRIPTCONTEXT.getDocument()  # type: ignore  # NOQA
    except Exception as exc:
        _log(f"_insert_image_run: failed to obtain document: {exc}")
        return None

    image_url = image_spec.get("url")
    temp_path = _download_image_to_temp(image_url)
    if not temp_path:
        return None

    try:
        _log(f"_insert_image_run: embedding graphic for {image_url}")
        
        # Use GraphicProvider to embed the graphic
        graphic_provider = XSCRIPTCONTEXT.getComponentContext().ServiceManager.createInstanceWithContext( # type: ignore # NOQA
            "com.sun.star.graphic.GraphicProvider", XSCRIPTCONTEXT.getComponentContext() # type: ignore # NOQA
        )
        from com.sun.star.beans import PropertyValue
        local_url = Path(temp_path).as_uri()
        props = (PropertyValue(Name="URL", Value=local_url),)
        graphic = graphic_provider.queryGraphic(props)

        graphic_obj = doc.createInstance("com.sun.star.text.TextGraphicObject")
        if graphic:
            graphic_obj.Graphic = graphic
            _log(f"_insert_image_run: graphic embedded for {image_url}")
        else:
            graphic_obj.GraphicURL = local_url
            _log(f"_insert_image_run: using GraphicURL fallback for {image_url}")

        # Clean up temp file now that it's embedded
        try:
            os.remove(temp_path)
        except Exception:
            pass

        width_in = image_spec.get("width_in")
        height_in = image_spec.get("height_in")
        if width_in:
            graphic_obj.Width = int(width_in * 2540)
        if height_in:
            graphic_obj.Height = int(height_in * 2540)

        is_block = bool(image_spec.get("block"))
        if is_block:
            graphic_obj.AnchorType = AT_PARAGRAPH
        else:
            graphic_obj.AnchorType = AS_CHARACTER
        _log(
            f"_insert_image_run: anchor_type={'AT_PARAGRAPH' if is_block else 'AS_CHARACTER'}"
        )

        text = cursor.getText()
        text.insertTextContent(cursor, graphic_obj, False)
        anchor = graphic_obj.getAnchor()
        try:
            cursor.goRight(1, False)
        except Exception:
            pass
        if anchor is None:
            _log("_insert_image_run: graphic anchor is None after insertion.")
        else:
            _log("_insert_image_run: graphic anchor obtained successfully.")
        return anchor
    except Exception as exc:
        _log(f"_insert_image_run: failed to insert image {image_url}: {exc}")
        return None
    finally:
        try:
            os.unlink(temp_path)
            _log(f"_insert_image_run: deleted temp file {temp_path}")
        except OSError:
            pass


# Style name mapping between Word/AI terminology and LibreOffice Writer
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
}


def _apply_paragraph_formatting(cursor, format_spec: Dict[str, Any]) -> None:
    """Apply paragraph-level formatting including style name and alignment.
    
    This function is critical for applying styles like 'Heading 1', 'Title', etc.
    which define the visual appearance of paragraphs in the document.
    """
    if not format_spec:
        return
    
    # Apply paragraph style (most important for headings, titles, etc.)
    style_name = format_spec.get("style_name")
    if style_name:
        # Map Word style names to LibreOffice counterparts
        mapped_style = STYLE_MAP.get(style_name, style_name)
        try:
            # Use setPropertyValue for better UNO bridge stability
            cursor.setPropertyValue("ParaStyleName", mapped_style)
            _log(f"_apply_paragraph_formatting: applied style '{mapped_style}' (original: '{style_name}')")
        except Exception as e:
            _log(f"_apply_paragraph_formatting: failed to apply style '{mapped_style}': {e} (Type: {type(e).__name__})")
    
    # Apply alignment
    alignment = format_spec.get("alignment")
    if alignment:
        alignment_map = {
            "left": 0,      # com.sun.star.style.ParagraphAdjust.LEFT
            "center": 3,    # com.sun.star.style.ParagraphAdjust.CENTER
            "right": 1,     # com.sun.star.style.ParagraphAdjust.RIGHT
            "justify": 2,   # com.sun.star.style.ParagraphAdjust.BLOCK
        }
        if alignment.lower() in alignment_map:
            try:
                cursor.setPropertyValue("ParaAdjust", alignment_map[alignment.lower()])
                _log(f"_apply_paragraph_formatting: applied alignment '{alignment}'")
            except Exception as e:
                _log(f"_apply_paragraph_formatting: failed to apply alignment: {e}")
    
    # Apply indentation (convert inches to 1/100 mm)
    if format_spec.get("indent_left_in"):
        try:
            cursor.setPropertyValue("ParaLeftMargin", int(format_spec["indent_left_in"] * 2540))
        except Exception as e:
            _log(f"_apply_paragraph_formatting: failed to apply left indent: {e}")
    
    if format_spec.get("indent_right_in"):
        try:
            cursor.setPropertyValue("ParaRightMargin", int(format_spec["indent_right_in"] * 2540))
        except Exception as e:
            _log(f"_apply_paragraph_formatting: failed to apply right indent: {e}")
    
    if format_spec.get("indent_first_line_in"):
        try:
            cursor.setPropertyValue("ParaFirstLineIndent", int(format_spec["indent_first_line_in"] * 2540))
        except Exception as e:
            _log(f"_apply_paragraph_formatting: failed to apply first line indent: {e}")
    
    # Apply spacing (convert points to 1/100 mm)
    if format_spec.get("space_before_pt"):
        try:
            cursor.setPropertyValue("ParaTopMargin", int(format_spec["space_before_pt"] * 2540 / 72))
        except Exception as e:
            _log(f"_apply_paragraph_formatting: failed to apply space before: {e}")
    
    if format_spec.get("space_after_pt"):
        try:
            cursor.setPropertyValue("ParaBottomMargin", int(format_spec["space_after_pt"] * 2540 / 72))
        except Exception as e:
            _log(f"_apply_paragraph_formatting: failed to apply space after: {e}")


def _apply_run_formatting(
    cursor, fmt: Dict[str, Any], override_color: Optional[int]
) -> None:
    if override_color is not None:
        cursor.CharColor = override_color
    else:
        if fmt.get("font_color_rgb"):
            try:
                cursor.CharColor = int(fmt["font_color_rgb"].lstrip("#"), 16)
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


def _looks_like_paragraph_spec(spec: Any) -> bool:
    if not isinstance(spec, dict):
        return False
    return any(key in spec for key in ("runs", "text", "format"))


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
    fmt = _merge_missing(
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
    normalized["format"] = fmt
    return normalized


def _normalize_table_paragraph(paragraph: Dict[str, Any], is_header: bool) -> Dict[str, Any]:
    normalized = dict(paragraph or {})
    normalized["format"] = _merge_missing(
        normalized.get("format") or {},
        {
            "style_name": "Normal",
            "alignment": "left",
            # Tiny positive values force UNO to write compact zero-ish margins.
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


def _write_table_cell_spec(cell, cell_cursor, cell_spec: Dict[str, Any], color: Optional[int]) -> None:
    cell_text = cell_cursor.getText()
    content_spec = cell_spec.get("content")
    if isinstance(content_spec, list) and content_spec:
        paragraph_items = []
        for item in content_spec:
            if not isinstance(item, dict):
                continue
            inner = item.get("paragraph")
            if isinstance(inner, dict) and inner.get("runs"):
                merged = dict(inner)
                if "format" not in merged and "format" in item:
                    merged["format"] = item["format"]
                paragraph_items.append(merged)
            elif _looks_like_paragraph_spec(item):
                paragraph_items.append(item)
        for idx, item in enumerate(paragraph_items):
            _write_paragraph_spec(
                cell_cursor,
                item,
                color=color,
                append_break=idx < len(paragraph_items) - 1,
                preserve_explicit_run_colors=True,
            )
        _apply_table_cell_formatting(cell, cell_cursor, cell_spec)
        return

    paragraph_spec = cell_spec.get("paragraph")
    if isinstance(paragraph_spec, dict):
        _write_paragraph_spec(
            cell_cursor,
            paragraph_spec,
            color=color,
            append_break=False,
            preserve_explicit_run_colors=True,
        )
        _apply_table_cell_formatting(cell, cell_cursor, cell_spec)
        return

    runs_spec = cell_spec.get("runs")
    if isinstance(runs_spec, list):
        _write_paragraph_spec(
            cell_cursor,
            {"runs": runs_spec},
            color=color,
            append_break=False,
            preserve_explicit_run_colors=True,
        )
        _apply_table_cell_formatting(cell, cell_cursor, cell_spec)
        return

    text_spec = cell_spec.get("text")
    if text_spec:
        cell_text.insertString(cell_cursor, text_spec, False)
        if color is not None:
            range_cursor = cell_text.createTextCursor()
            range_cursor.gotoStart(False)
            range_cursor.gotoEnd(True)
            range_cursor.CharColor = color
    _apply_table_cell_formatting(cell, cell_cursor, cell_spec)


def _parse_color_value(value):
    if not isinstance(value, str):
        return None
    raw = value.strip()
    if raw.startswith("#"):
        raw = raw[1:]
    try:
        return int(raw, 16)
    except Exception:
        return None


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

    shading_color = _parse_color_value(cell_spec.get("shading_color"))
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


def _insert_table_from_spec(
    doc, cursor, spec: Dict[str, Any], color: Optional[int] = None
):
    spec = _normalize_table_spec(spec)
    has_explicit_cell_color = _table_spec_has_explicit_run_color(spec)
    rows = len(spec.get("rows") or [])
    cols = len(spec.get("rows")[0]["cells"]) if rows else 0
    table = doc.createInstance("com.sun.star.text.TextTable")
    if rows and cols:
        table.initialize(rows, cols)
    if spec.get("style"):
        try:
            table.TableStyleName = spec["style"]
        except Exception:
            pass
    text = cursor.getText()
    text.insertTextContent(cursor, table, False)

    for r_idx, row in enumerate(spec.get("rows") or []):
        for c_idx, cell_spec in enumerate(row.get("cells") or []):
            cell = table.getCellByPosition(c_idx, r_idx)
            cell_text = cell.getText()
            cell_cursor = cell_text.createTextCursor()
            cell_text.setString("")
            if cell_spec.get("style"):
                try:
                    cell.CellStyleName = cell_spec["style"]
                except Exception:
                    pass
            _write_table_cell_spec(cell, cell_cursor, cell_spec, color)
    first_row = (spec.get("rows") or [{}])[0].get("cells") or []
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
    if color is not None and not has_explicit_cell_color:
        _colorize_table(table, color)
    return table


def _colorize_table(table, color: int, strike: bool = False) -> None:
    rows = table.getRows().getCount()
    cols = table.getColumns().getCount()
    for r in range(rows):
        for c in range(cols):
            cell = table.getCellByPosition(c, r)
            cell_text = cell.getText()
            cursor = cell_text.createTextCursor()
            cursor.gotoEnd(True)
            cursor.CharColor = color
            cursor.CharUnderline = FONT_UNDERLINE_SINGLE if strike else 0
            cursor.CharStrikeout = 2 if strike else 0


def _tag_preview_range(doc, target, plan_id: str, op_id: str, role: str) -> str:
    bookmark_name = (
        f"{PREVIEW_BOOKMARK_PREFIX}_{plan_id}_{op_id}_{role}".replace("-", "_")
    )
    try:
        if isinstance(target, tuple):
            text, start_range, end_range = target
            cursor = text.createTextCursorByRange(start_range)
            cursor.gotoRange(end_range, True)
        else:
            text = target.getText()
            # For paragraphs, getStart/getEnd can be fussy with bookmarks
            # Try to create a cursor that spans the content
            cursor = text.createTextCursorByRange(target.getStart())
            cursor.gotoRange(target.getEnd(), True)
    except Exception as e:
        _log(f"_tag_preview_range: failed to create cursor: {e}")
        # Final fallback
        if hasattr(target, "getText"):
            text = target.getText()
            cursor = text.createTextCursor()
            if hasattr(target, "getStart"):
                cursor.gotoRange(target.getStart(), False)
                if hasattr(target, "getEnd"):
                    cursor.gotoRange(target.getEnd(), True)
        else:
            # Absolute fallback if target is unknown
            return ""

    _create_bookmark(doc, text, cursor, bookmark_name)
    return bookmark_name


def _tag_named_range(doc, target, bookmark_name: Optional[str]) -> Optional[str]:
    if not isinstance(bookmark_name, str) or not bookmark_name.strip():
        return None

    try:
        if isinstance(target, tuple):
            text, start_range, end_range = target
            cursor = text.createTextCursorByRange(start_range)
            cursor.gotoRange(end_range, True)
        else:
            text = target.getText()
            cursor = text.createTextCursorByRange(target.getStart())
            cursor.gotoRange(target.getEnd(), True)
        _create_bookmark(doc, text, cursor, bookmark_name)
        return bookmark_name
    except Exception as exc:
        _log(f"_tag_named_range: failed to create bookmark {bookmark_name}: {exc}")
        return None


def _tag_table(doc, table, plan_id: str, op_id: str, role: str, text=None) -> str:
    bookmark_name = (
        f"{PREVIEW_BOOKMARK_PREFIX}_{plan_id}_{op_id}_{role}".replace("-", "_")
    )
    anchor = table.getAnchor()

    # Prioritize the text object from the anchor itself to ensure compatibility
    anchor_text = anchor.getText()
    if anchor_text is not None:
        text = anchor_text
    # If text is still None, try to get it from the anchor's start point
    if text is None:
        try:
            # If the anchor is inside a table cell, it will have a 'Text' parent
            text = anchor.getStart().Text
        except Exception:
            try:
                text = anchor.getStart().getText()
            except Exception:
                pass

    # Final fallback to doc.getText()
    if text is None:
        text = doc.getText()
        _log(f"_tag_table: total fallback to doc.getText() for {op_id}")

    cursor = None
    try:
        # Try to create cursor from anchor range first
        cursor = text.createTextCursorByRange(anchor)
        cursor.gotoRange(anchor.getEnd(), True)
    except Exception as e:
        _log(f"_tag_table: failed to create cursor from anchor range: {e}")
        try:
            # Fallback: create cursor at anchor start position
            start_pos = anchor.getStart()
            cursor = text.createTextCursorByRange(start_pos)
            # Try to expand to end, but if that fails, just use start position
            try:
                cursor.gotoRange(anchor.getEnd(), True)
            except Exception:
                # If we can't expand to end, just use start position
                pass
        except Exception as e2:
            _log(f"_tag_table: failed to create cursor from anchor start: {e2}")
            try:
                # Last resort: anchor inside the first table cell paragraph.
                cell_names = list(table.getCellNames())
                if cell_names:
                    first_cell = table.getCellByName(cell_names[0])
                    cell_text = first_cell.getText()
                    cell_enum = cell_text.createEnumeration()
                    while cell_enum.hasMoreElements():
                        para = cell_enum.nextElement()
                        try:
                            if para.supportsService("com.sun.star.text.Paragraph"):
                                cursor = cell_text.createTextCursorByRange(
                                    para.getStart()
                                )
                                cursor.collapseToStart()
                                text = cell_text
                                break
                        except Exception:
                            continue
                    if cursor is None:
                        cursor = cell_text.createTextCursor()
                        cursor.gotoStart(False)
                        text = cell_text
            except Exception as e3:
                _log(f"_tag_table: all cursor creation methods failed: {e3}")
                cursor = None

    if cursor is None:
        _log(f"_tag_table: could not create cursor for table bookmark {bookmark_name}")
        return ""

    _create_bookmark(doc, text, cursor, bookmark_name)
    try:
        bookmarks = doc.getBookmarks()
        if bookmarks.hasByName(bookmark_name):
            return bookmark_name
    except Exception:
        pass
    _log(f"_tag_table: bookmark did not materialize for {bookmark_name}")
    return ""


def _insert_separator(doc, cursor, plan_id: str, op_id: str, role: str) -> str:
    """
    Inserts a paragraph break to create spacing, but skips if we are already
    preceded by an empty paragraph. This ensures "flow" as requested.
    """
    text = cursor.getText()
    bookmark_name = (
        f"{PREVIEW_BOOKMARK_PREFIX}_{plan_id}_{op_id}_{role}".replace("-", "_")
    )

    # Smart check for redundant breaks
    try:
        check = text.createTextCursorByRange(cursor)
        if not check.goLeft(1, False):
            # At start of doc, maybe skip? Let's skip for now to avoid top empty lines
            return bookmark_name

        # Check if the paragraph at the cursor (or just before it) is empty
        check.gotoStartOfParagraph(False)
        check.gotoEndOfParagraph(True)
        if check.getString().strip() == "":
            _log(f"_insert_separator: skipping redundant break for {op_id}")
            return bookmark_name
    except Exception as e:
        _log(f"_insert_separator: smart check failed: {e}")

    # Insert a real empty line
    text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)

    # Tag the new empty line (select the break char)
    tag_cursor = text.createTextCursorByRange(cursor)
    tag_cursor.goLeft(1, True)
    _create_bookmark(doc, text, tag_cursor, bookmark_name)

    return bookmark_name


def _insert_table_trailing_boundary(
    doc, table, plan_id: str, op_id: str, role: str
) -> str:
    try:
        cursor = _create_cursor_after_table(doc, table)
        if cursor is None:
            return ""
        text = cursor.getText()
        bookmark_name = (
            f"{PREVIEW_BOOKMARK_PREFIX}_{plan_id}_{op_id}_{role}".replace("-", "_")
        )
        text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
        tag_cursor = text.createTextCursorByRange(cursor)
        if not tag_cursor.goLeft(1, True):
            return ""
        _create_bookmark(doc, text, tag_cursor, bookmark_name)
        return bookmark_name
    except Exception as exc:
        _log(f"_insert_table_trailing_boundary: failed for {op_id}: {exc}")
        return ""


def _insert_index_trailing_boundary(
    doc, index, plan_id: str, op_id: str, role: str
) -> str:
    try:
        anchor = index.getAnchor()
        text = anchor.getText() or doc.getText()
        cursor = text.createTextCursorByRange(anchor.getEnd())
        cursor.collapseToStart()
        bookmark_name = (
            f"{PREVIEW_BOOKMARK_PREFIX}_{plan_id}_{op_id}_{role}".replace("-", "_")
        )
        text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
        tag_cursor = text.createTextCursorByRange(cursor)
        if not tag_cursor.goLeft(1, True):
            return ""
        _create_bookmark(doc, text, tag_cursor, bookmark_name)
        return bookmark_name
    except Exception as exc:
        _log(f"_insert_index_trailing_boundary: failed for {op_id}: {exc}")
        return ""


def _create_bookmark(doc, text, cursor, name: str) -> None:
    bookmarks = doc.getBookmarks()
    if bookmarks.hasByName(name):
        dispose_rounds = 0
        while bookmarks.hasByName(name) and dispose_rounds < 8:
            existing = bookmarks.getByName(name)
            existing.dispose()
            dispose_rounds += 1
        if bookmarks.hasByName(name):
            raise RuntimeError(
                f"_create_bookmark: could not clear existing bookmark name '{name}'"
            )
    bookmark = doc.createInstance("com.sun.star.text.Bookmark")
    bookmark.setName(name)
    text.insertTextContent(cursor, bookmark, True)


def _extract_paragraph_spec(paragraph) -> Dict[str, Any]:
    spec = {"runs": []}
    enum = paragraph.createEnumeration()
    while enum.hasMoreElements():
        portion = enum.nextElement()
        try:
            portion_type = portion.TextPortionType
        except Exception:
            portion_type = None
        if portion_type == "LineBreak":
            spec["runs"].append({"text": "", "break_before": "line"})
            continue
        if not hasattr(portion, "getString"):
            continue
        text = portion.getString()
        if not text:
            continue
        cursor = paragraph.getText().createTextCursorByRange(portion.getStart())
        cursor.gotoRange(portion.getEnd(), True)
        run_spec = {
            "text": text,
            "format": {
                "font_name": cursor.CharFontName,
                "font_size_pt": cursor.CharHeight,
                "font_color_rgb": (
                    f"{cursor.CharColor:06X}" if cursor.CharColor is not None else None
                ),
                "bold": cursor.CharWeight == FONT_WEIGHT_BOLD,
                "italic": cursor.CharPosture == FONT_SLANT_ITALIC,
                "underline": cursor.CharUnderline == FONT_UNDERLINE_SINGLE,
            },
        }
        spec["runs"].append(run_spec)
    return spec


def _extract_table_spec(table) -> Dict[str, Any]:
    rows = table.getRows().getCount()
    cols = table.getColumns().getCount() if rows else 0
    spec_rows: List[Dict[str, Any]] = []
    for r in range(rows):
        row_cells: List[Dict[str, Any]] = []
        for c in range(cols):
            cell = table.getCellByPosition(c, r)
            text = cell.getText().getString()
            row_cells.append({"text": text})
        spec_rows.append({"cells": row_cells})
    return {"rows": spec_rows}


def _safe_table_name(table) -> Optional[str]:
    try:
        name = table.getName()
        if isinstance(name, str) and name:
            return name
    except Exception:
        pass
    return None


def _apply_run_edit_to_spec(spec: Dict[str, Any], action: Dict[str, Any]) -> None:
    runs = spec.get("runs") or []
    run_index = action.get("run_index")
    run = action.get("run") or {}
    if (
        action.get("kind") == "replace_run"
        and runs
        and run_index is not None
        and run_index < len(runs)
    ):
        runs[run_index] = run
    elif action.get("kind") == "insert_run":
        if run_index is None or run_index >= len(runs):
            runs.append(run)
        else:
            runs.insert(run_index, run)
    spec["runs"] = runs


def _get_element_by_bookmark(
    doc, bookmark_name: Optional[str], prefer_table: bool = False
) -> Optional[Tuple[str, Any]]:
    if not bookmark_name:
        return None
    try:
        bookmarks = doc.getBookmarks()
        if not bookmarks.hasByName(bookmark_name):
            _log(
                f"_get_element_by_bookmark: bookmark '{bookmark_name}' not found by name. Total bookmarks: {bookmarks.getCount()}"
            )
            return None
        bookmark = bookmarks.getByName(bookmark_name)
        anchor = bookmark.getAnchor()
        doc_text = doc.getText()
        anchor_start = _safe_get_start(anchor)
        anchor_text = None
        try:
            anchor_text = anchor.getText()
        except Exception:
            anchor_text = None

        if prefer_table:
            # Fast path for table-targeting actions.
            try:
                owner_table = anchor.getTextTable()
                if owner_table is not None:
                    return ("table", owner_table)
            except Exception:
                pass

        # 1) Fast path: top-level enumeration.
        try:
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
                            if (
                                doc_text.compareRegionStarts(
                                    table_anchor.getStart(), anchor.getStart()
                                )
                                == 0
                            ):
                                return ("table", element)
                    elif element.supportsService("com.sun.star.text.Paragraph"):
                        par_start = element.getStart()
                        try:
                            if doc_text.compareRegionStarts(par_start, anchor) == 0:
                                return ("paragraph", element)
                        except Exception:
                            if (
                                doc_text.compareRegionStarts(
                                    par_start, anchor.getStart()
                                )
                                == 0
                            ):
                                return ("paragraph", element)
                except UNOException:
                    continue
        except Exception:
            pass

        # 2) Fallback: paragraphs nested inside table cells.
        try:
            tables = doc.getTextTables()
            table_names = list(tables.getElementNames())
        except Exception:
            table_names = []

        for table_name in table_names:
            try:
                table = tables.getByName(table_name)
                # table anchor check (for cases where top-level enumeration misses table)
                try:
                    table_anchor = table.getAnchor()
                    try:
                        if doc_text.compareRegionStarts(table_anchor, anchor) == 0:
                            return ("table", table)
                    except Exception:
                        if (
                            doc_text.compareRegionStarts(
                                table_anchor.getStart(), anchor.getStart()
                            )
                                == 0
                            ):
                                return ("table", table)
                except Exception:
                    pass

                for cell_name in table.getCellNames():
                    try:
                        cell = table.getCellByName(cell_name)
                        cell_text = cell.getText()

                        # If bookmark anchor belongs to this cell text container,
                        # this table is the owner. This is common for table bookmarks
                        # anchored at first-cell paragraph start.
                        if anchor_text is not None:
                            try:
                                if cell_text == anchor_text:
                                    if prefer_table:
                                        return ("table", table)
                            except Exception:
                                pass

                        # Region-compare directly against cell text context.
                        try:
                            if (
                                anchor_start is not None
                                and cell_text.compareRegionStarts(
                                    anchor_start, anchor_start
                                )
                                == 0
                            ):
                                if prefer_table:
                                    return ("table", table)
                        except Exception:
                            pass

                        cell_enum = cell_text.createEnumeration()
                        while cell_enum.hasMoreElements():
                            para = cell_enum.nextElement()
                            try:
                                if not para.supportsService(
                                    "com.sun.star.text.Paragraph"
                                ):
                                    continue
                            except Exception:
                                continue

                            par_start = para.getStart()
                            try:
                                if doc_text.compareRegionStarts(par_start, anchor) == 0:
                                    if prefer_table:
                                        return ("table", table)
                                    return ("paragraph", para)
                            except Exception:
                                try:
                                    if (
                                        cell_text.compareRegionStarts(
                                            par_start, anchor
                                        )
                                        == 0
                                    ):
                                        if prefer_table:
                                            return ("table", table)
                                        return ("paragraph", para)
                                except Exception:
                                    continue
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception as e:
        _log(f"_get_element_by_bookmark: error during lookup: {e}")
        return None
    return None


def _get_element_by_index(doc, index: int) -> Optional[Tuple[str, Any]]:
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


g_exportedScripts = (previewActionPlan,)
