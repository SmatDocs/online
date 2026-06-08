"""
EnsureSdocBookmarks.py
----------------------

Ensures Writer content has stable SmartDocs bookmarks.
It adds missing `sdoc_*` bookmarks for:
- top-level body paragraphs
- tables
- paragraphs inside table cells
when they do not already have a managed bookmark anchored at their start.
"""

import json
import re
import time
import traceback
import uuid
from datetime import datetime
from typing import Any, Dict, List, Optional, Set, Tuple

try:  # LibreOffice runtime
    from com.sun.star.uno import Exception as UNOException
except ImportError:  # pragma: no cover (local/static analysis fallback)
    UNOException = Exception  # type: ignore


LOG_FILE = "/tmp/sdoc_bookmarks.log"
BOOKMARK_PREFIX = "sdoc_"
BOOKMARK_RE = re.compile(r"^sdoc_[A-Za-z0-9]+$")
TABLE_BOOKMARK_PREFIX = f"{BOOKMARK_PREFIX}t"
PARAGRAPH_BOOKMARK_PREFIX = f"{BOOKMARK_PREFIX}p"
PREVIEW_BOOKMARK_PREFIXES = ("ov_preview_", "sdoc_preview_")
PREVIEW_STATE_KEY = "sdoc_preview_state"
PREVIEW_GREEN_RGB = 0x008A00
SCRIPT_VERSION = "2026-02-20.table-anchor-separation.v1"


def _log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    full_message = f"[{timestamp}] {message}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as handle:
            handle.write(full_message + "\n")
    except Exception:
        pass
    try:
        print(full_message, flush=True)
    except UnicodeEncodeError:
        safe = full_message.encode("ascii", "backslashreplace").decode("ascii")
        print(safe, flush=True)


def _is_managed_bookmark_name(name: str) -> bool:
    return isinstance(name, str) and bool(BOOKMARK_RE.match(name))


def _is_table_bookmark_name(name: Optional[str]) -> bool:
    return isinstance(name, str) and name.startswith(TABLE_BOOKMARK_PREFIX)


def _is_paragraph_bookmark_name(name: Optional[str]) -> bool:
    return isinstance(name, str) and name.startswith(PARAGRAPH_BOOKMARK_PREFIX)


def _is_preview_bookmark_name(name: str) -> bool:
    return isinstance(name, str) and any(
        name.startswith(prefix) for prefix in PREVIEW_BOOKMARK_PREFIXES
    )


def _should_lock_controllers(payload: Dict[str, Any], request_id: Optional[str]) -> bool:
    explicit = payload.get("lock_controllers")
    if isinstance(explicit, bool):
        return explicit
    # Frontend/live bookmark requests carry request_id. Avoid freezing Collabora's
    # visible controller during large bookmark seeding while preserving the old
    # locked behavior for headless upload/project indexing.
    return not bool(request_id)


def _user_props_has_property(props, name: str) -> bool:
    try:
        props.getPropertyValue(name)
        return True
    except Exception:
        return False


def _load_preview_state(doc) -> Dict[str, Dict[str, Any]]:
    try:
        props = doc.getDocumentProperties().getUserDefinedProperties()
    except Exception:
        return {}
    if not _user_props_has_property(props, PREVIEW_STATE_KEY):
        return {}
    try:
        raw = props.getPropertyValue(PREVIEW_STATE_KEY)
        return json.loads(raw) if raw else {}
    except Exception:
        return {}


def _has_unresolved_preview_state(doc) -> Tuple[bool, int]:
    state = _load_preview_state(doc)
    if not state:
        return False, 0

    unresolved = 0
    for plan_state in state.values():
        if not isinstance(plan_state, dict):
            continue
        for op_state in plan_state.values():
            if not isinstance(op_state, dict):
                unresolved += 1
                continue
            # ResolvePreviewDiff marks preview operations that could not be
            # resolved as "failed". Those are terminal for confirmation: they
            # should not keep bookmark refresh blocked forever after the visible
            # preview has already been accepted/rejected.
            if op_state.get("status") not in ("resolved", "failed", "skipped"):
                unresolved += 1

    return unresolved > 0, unresolved


def _purge_preview_bookmarks(bookmarks) -> Tuple[int, int]:
    removed = 0
    failed = 0
    for name in list(bookmarks.getElementNames()):
        if not _is_preview_bookmark_name(name):
            continue
        try:
            if bookmarks.hasByName(name):
                bookmarks.getByName(name).dispose()
                removed += 1
        except Exception:
            failed += 1
    return removed, failed


def _next_bookmark_name(existing_names: Set[str], kind: Optional[str] = None) -> str:
    marker = ""
    if kind == "table":
        marker = "t"
    elif kind in {"paragraph", "cell_paragraph"}:
        marker = "p"

    while True:
        candidate = f"{BOOKMARK_PREFIX}{marker}{uuid.uuid4().hex[:6]}"
        if candidate not in existing_names:
            existing_names.add(candidate)
            return candidate


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


def _iter_text_tables(doc) -> List[Any]:
    tables: List[Any] = []
    try:
        collection = doc.getTextTables()
    except Exception:
        return tables

    try:
        names = list(collection.getElementNames())
        for name in names:
            try:
                tables.append(collection.getByName(name))
            except Exception:
                continue
        return tables
    except Exception:
        pass

    try:
        count = int(collection.getCount())
        for idx in range(count):
            try:
                tables.append(collection.getByIndex(idx))
            except Exception:
                continue
    except Exception:
        pass

    return tables


def _is_preview_green_color(value: Any) -> bool:
    try:
        return int(value) == PREVIEW_GREEN_RGB
    except Exception:
        return False


def _count_preview_green_cursor(cursor) -> int:
    try:
        return 1 if _is_preview_green_color(getattr(cursor, "CharColor", None)) else 0
    except Exception:
        return 0


def _count_preview_green_range(text_obj, start, end) -> int:
    try:
        cursor = text_obj.createTextCursorByRange(start)
        cursor.gotoRange(end, True)
        return _count_preview_green_cursor(cursor)
    except Exception:
        return 0


def _count_preview_green_paragraph(paragraph) -> int:
    count = 0
    scanned_portions = False
    try:
        text_obj = paragraph.getText()
    except Exception:
        return count

    try:
        enum = paragraph.createEnumeration()
        while enum.hasMoreElements():
            portion = enum.nextElement()
            try:
                count += _count_preview_green_range(
                    text_obj,
                    portion.getStart(),
                    portion.getEnd(),
                )
                scanned_portions = True
            except Exception:
                continue
    except Exception:
        pass

    if not scanned_portions:
        try:
            count += _count_preview_green_range(
                text_obj,
                paragraph.getStart(),
                paragraph.getEnd(),
            )
        except Exception:
            pass

    return count


def _count_preview_green_text(text_obj) -> int:
    count = 0
    scanned_elements = False
    try:
        enum = text_obj.createEnumeration()
        while enum.hasMoreElements():
            element = enum.nextElement()
            try:
                if element.supportsService("com.sun.star.text.Paragraph"):
                    count += _count_preview_green_paragraph(element)
                    scanned_elements = True
                elif element.supportsService("com.sun.star.text.TextTable"):
                    count += _count_preview_green_table(element)
                    scanned_elements = True
            except Exception:
                continue
    except Exception:
        pass

    if not scanned_elements:
        try:
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(True)
            count += _count_preview_green_cursor(cursor)
        except Exception:
            pass

    return count


def _count_preview_green_table(table) -> int:
    count = 0
    try:
        rows = table.getRows().getCount()
        cols = table.getColumns().getCount()
    except Exception:
        return count
    for row in range(rows):
        for col in range(cols):
            try:
                count += _count_preview_green_text(table.getCellByPosition(col, row).getText())
            except Exception:
                continue
    return count


def _count_preview_green_formatting(doc) -> int:
    count = 0

    try:
        for kind, element in _iter_body_elements(doc):
            if kind == "paragraph":
                count += _count_preview_green_paragraph(element)
            elif kind == "table":
                count += _count_preview_green_table(element)
    except Exception as exc:
        _log(f"ensureSdocBookmarks: preview green body scan failed: {exc}")

    try:
        for table in _iter_text_tables(doc):
            count += _count_preview_green_table(table)
    except Exception as exc:
        _log(f"ensureSdocBookmarks: preview green table scan failed: {exc}")

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
                    count += _count_preview_green_text(text_obj)
    except Exception as exc:
        _log(f"ensureSdocBookmarks: preview green page-style scan failed: {exc}")

    return count


def _collect_managed_anchors(doc, bookmarks) -> List[Tuple[str, Any]]:
    anchors: List[Tuple[str, Any]] = []
    for name in bookmarks.getElementNames():
        if not _is_managed_bookmark_name(name):
            continue
        try:
            anchor = bookmarks.getByName(name).getAnchor()
            anchors.append((name, anchor))
        except Exception:
            continue
    return anchors


def _anchor_matches_start(doc_text, target_start, anchor) -> bool:
    owner_text = None
    try:
        owner_text = target_start.getText()
    except Exception:
        owner_text = None

    for text_obj in (owner_text, doc_text):
        if text_obj is None:
            continue
        try:
            if text_obj.compareRegionStarts(anchor, target_start) == 0:
                return True
        except Exception:
            continue
    return False


def _has_anchor_at_start(doc_text, target_start, managed_anchors: List[Tuple[str, Any]]) -> bool:
    for _, anchor in managed_anchors:
        if _anchor_matches_start(doc_text, target_start, anchor):
            return True
    return False


def _has_table_anchor_at_start(
    doc_text, target_start, managed_anchors: List[Tuple[str, Any]]
) -> bool:
    for name, anchor in managed_anchors:
        if not _is_table_bookmark_name(name):
            continue
        if _anchor_matches_start(doc_text, target_start, anchor):
            return True
    return False


def _anchor_matches_table_anchor(doc_text, table, anchor) -> bool:
    try:
        table_anchor = table.getAnchor()
    except Exception:
        table_anchor = None
    if table_anchor is None:
        return False

    targets: List[Any] = [table_anchor]
    try:
        targets.append(table_anchor.getStart())
    except Exception:
        pass

    for target in targets:
        if target is not None and _anchor_matches_start(doc_text, target, anchor):
            return True
    return False


def _anchor_belongs_to_table(anchor, table) -> bool:
    try:
        owner = anchor.getTextTable()
    except Exception:
        owner = None
    if owner is None:
        return False

    try:
        if owner == table:
            return True
    except Exception:
        pass

    try:
        owner_name = owner.getName()
        table_name = table.getName()
        if owner_name and table_name and owner_name == table_name:
            return True
    except Exception:
        pass

    return False


def _collect_table_cell_paragraph_starts(table) -> List[Any]:
    starts: List[Any] = []
    cell_names: List[str] = []
    try:
        cell_names = list(table.getCellNames())
    except Exception:
        return starts

    for cell_name in cell_names:
        try:
            cell = table.getCellByName(cell_name)
            cell_text = cell.getText()
            enum = cell_text.createEnumeration()
        except Exception:
            continue

        while enum.hasMoreElements():
            try:
                elem = enum.nextElement()
            except Exception:
                break
            try:
                if not elem.supportsService("com.sun.star.text.Paragraph"):
                    continue
            except Exception:
                continue
            try:
                starts.append(elem.getStart())
            except Exception:
                continue

    return starts


def _anchor_matches_any_start(doc_text, anchor, starts: List[Any]) -> bool:
    for target_start in starts:
        if _anchor_matches_start(doc_text, target_start, anchor):
            return True
    return False


def _has_table_anchor_for_table(
    doc_text, table, managed_anchors: List[Tuple[str, Any]]
) -> bool:
    for name, anchor in managed_anchors:
        if not _is_table_bookmark_name(name):
            continue
        if _anchor_matches_table_anchor(doc_text, table, anchor):
            return True
    return False


def _has_non_table_anchor_at_start(
    doc_text, target_start, managed_anchors: List[Tuple[str, Any]]
) -> bool:
    for name, anchor in managed_anchors:
        if _is_table_bookmark_name(name):
            continue
        if _anchor_matches_start(doc_text, target_start, anchor):
            return True
    return False


def _create_bookmark(doc, text, cursor, name: str) -> None:
    bookmark = doc.createInstance("com.sun.star.text.Bookmark")
    bookmark.setName(name)
    text.insertTextContent(cursor, bookmark, True)


def _create_collapsed_cursor(text_obj, text_range):
    cursor = text_obj.createTextCursorByRange(text_range)
    cursor.collapseToStart()
    return cursor


def ensureSdocBookmarks(payload_json: str = "{}") -> str:
    start_ts = datetime.now()
    perf_start = time.perf_counter()
    perf_last = perf_start
    request_id = None
    reason = "unspecified"
    file_id = None
    timings: Dict[str, int] = {}

    def _mark_timing(stage: str) -> None:
        nonlocal perf_last
        now = time.perf_counter()
        stage_ms = int((now - perf_last) * 1000)
        total_ms = int((now - perf_start) * 1000)
        timings[stage] = total_ms
        perf_last = now
        _log(
            "ensureSdocBookmarks: timing "
            + f"stage={stage} stage_ms={stage_ms} total_ms={total_ms} "
            + f"reason={reason} request_id={request_id} file_id={file_id}"
        )

    try:
        payload = json.loads(payload_json) if payload_json else {}
        request_id = payload.get("request_id")
        reason = payload.get("reason") or reason
        file_id = payload.get("file_id")
    except Exception:
        payload = {}
        _log("ensureSdocBookmarks: failed to parse payload_json; proceeding with defaults")

    def _result(ok: bool, **fields: Any) -> str:
        elapsed_ms = int((datetime.now() - start_ts).total_seconds() * 1000)
        out = {
            "ok": ok,
            "event": "ensure_sdoc_bookmarks",
            "request_id": request_id,
            "reason": reason,
            "file_id": file_id,
            "elapsed_ms": elapsed_ms,
            "timings": timings,
        }
        out.update(fields)
        try:
            return json.dumps(out)
        except Exception:
            return json.dumps(
                {
                    "ok": False,
                    "event": "ensure_sdoc_bookmarks",
                    "request_id": request_id,
                    "reason": reason,
                    "file_id": file_id,
                    "elapsed_ms": elapsed_ms,
                    "error": "failed to serialize result",
                }
            )

    diagnostics: Dict[str, Any] = {
        "script_version": SCRIPT_VERSION,
        "table_scan_started": False,
        "table_scan_completed": False,
        "table_anchor_attempts": 0,
        "table_anchor_added": 0,
        "table_anchor_skipped_fast_path": 0,
        "table_cell_paragraph_attempts": 0,
        "table_cell_paragraph_added": 0,
        "table_cell_empty_skipped_fast_path": 0,
        "top_level_empty_skipped_fast_path": 0,
        "failure_buckets": {
            "top_level_paragraph": 0,
            "table_level": 0,
            "table_cell_paragraph": 0,
        },
        "sample_errors": [],
    }
    lock_controllers = _should_lock_controllers(payload, request_id)
    diagnostics["controllers_locked"] = lock_controllers

    def _record_error(stage: str, exc: Exception) -> None:
        try:
            if stage in diagnostics["failure_buckets"]:
                diagnostics["failure_buckets"][stage] += 1
            samples = diagnostics["sample_errors"]
            if len(samples) < 10:
                samples.append(
                    {
                        "stage": stage,
                        "error_type": exc.__class__.__name__,
                        "message": str(exc),
                    }
                )
            _log(
                "ensureSdocBookmarks: stage_error "
                + f"stage={stage} type={exc.__class__.__name__} message={exc}"
            )
        except Exception:
            pass

    try:
        doc = XSCRIPTCONTEXT.getDocument()  # type: ignore  # NOQA
        if doc is None:
            _log(
                "ensureSdocBookmarks: no active document "
                + f"reason={reason} request_id={request_id} file_id={file_id}"
            )
            return _result(False, error="no active document")
    except Exception as exc:
        _log(
            "ensureSdocBookmarks: failed to get document "
            + f"reason={reason} request_id={request_id} file_id={file_id} error={exc}"
        )
        return _result(False, error=f"failed to get document: {exc}", diagnostics=diagnostics)

    _log(
        "ensureSdocBookmarks: start "
        + f"reason={reason} request_id={request_id} file_id={file_id}"
    )

    if not doc.supportsService("com.sun.star.text.TextDocument"):
        _log(
            "ensureSdocBookmarks: skipped non-writer document "
            + f"reason={reason} request_id={request_id} file_id={file_id}"
        )
        return _result(
            True,
            skipped=True,
            skip_reason="not_writer_document",
            diagnostics=diagnostics,
        )

    try:
        bookmarks = doc.getBookmarks()
    except Exception as exc:
        _log(
            "ensureSdocBookmarks: failed to access bookmarks "
            + f"reason={reason} request_id={request_id} file_id={file_id} error={exc}"
        )
        return _result(
            False,
            error=f"failed to access bookmarks: {exc}",
            diagnostics=diagnostics,
        )

    preview_removed = 0
    preview_failed = 0
    preview_green_remaining = 0
    unresolved_preview_ops = 0
    skip_preview_purge = False
    preview_state: Dict[str, Dict[str, Any]] = {}
    should_check_preview_state = "preview" in str(reason) or "resolve" in str(reason)
    if should_check_preview_state:
        try:
            preview_state = _load_preview_state(doc)
            if preview_state:
                unresolved = 0
                for plan_state in preview_state.values():
                    if not isinstance(plan_state, dict):
                        continue
                    for op_state in plan_state.values():
                        if not isinstance(op_state, dict):
                            unresolved += 1
                            continue
                        if op_state.get("status") not in ("resolved", "failed", "skipped"):
                            unresolved += 1
                skip_preview_purge = unresolved > 0
                unresolved_preview_ops = unresolved
        except Exception as exc:
            _log(f"ensureSdocBookmarks: failed preview-state check: {exc}")
    else:
        diagnostics["preview_state_check_skipped"] = True
    _mark_timing("preview_state_check")

    if should_check_preview_state:
        try:
            if skip_preview_purge:
                _log(
                    "ensureSdocBookmarks: preview bookmark purge skipped "
                    + f"reason={reason} unresolved_preview_ops={unresolved_preview_ops}"
                )
            else:
                preview_removed, preview_failed = _purge_preview_bookmarks(bookmarks)
                if preview_removed or preview_failed:
                    _log(
                        "ensureSdocBookmarks: preview bookmark purge "
                        + f"removed={preview_removed} failed={preview_failed}"
                    )
        except Exception as exc:
            preview_failed += 1
            _log(f"ensureSdocBookmarks: preview purge failed: {exc}")
    else:
        diagnostics["preview_bookmark_purge_skipped_fast_path"] = True
    _mark_timing("preview_bookmark_purge")

    should_scan_preview_green = bool(preview_state) or "preview" in str(reason) or "resolve" in str(reason)
    if should_scan_preview_green:
        try:
            preview_green_remaining = _count_preview_green_formatting(doc)
            if preview_green_remaining:
                _log(
                    "ensureSdocBookmarks: preview green formatting remains "
                    + f"count={preview_green_remaining} reason={reason}"
                )
        except Exception as exc:
            _log(f"ensureSdocBookmarks: preview green scan failed: {exc}")
    else:
        diagnostics["preview_green_scan_skipped"] = True
    _mark_timing("preview_green_scan")

    existing_names = set(bookmarks.getElementNames())
    managed_count_before = sum(
        1 for name in existing_names if _is_managed_bookmark_name(name)
    )
    doc_text = doc.getText()
    elements = _iter_body_elements(doc)
    managed_anchors = _collect_managed_anchors(doc, bookmarks)

    added = 0
    failed = 0
    skipped_existing = 0
    was_recording = getattr(doc, "RecordChanges", False)
    paragraph_count = 0
    table_count = 0
    table_cell_paragraph_count = 0
    created_bookmarks: List[str] = []  # track names of all bookmarks we create
    fresh_managed_document = managed_count_before == 0
    bookmark_map: List[Dict[str, Any]] = []

    def _paragraph_preview(para_elem: Any) -> str:
        try:
            try:
                raw_text = para_elem.getString() or ""
            except Exception:
                text_obj = para_elem.getText()
                cursor = text_obj.createTextCursorByRange(para_elem.getStart())
                cursor.gotoStartOfParagraph(False)
                cursor.gotoEndOfParagraph(True)
                raw_text = cursor.getString() or ""
            raw_text = raw_text.strip()
            words = raw_text.split()
            if len(words) > 8:
                preview = " ".join(words[:8]) + "..."
            elif len(raw_text) > 80:
                preview = raw_text[:80] + "..."
            else:
                preview = raw_text
            return preview.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
        except Exception:
            return ""

    def _append_bookmark_map_entry(name: str, kind: str, para_elem: Any = None) -> None:
        entry: Dict[str, Any] = {"id": name, "kind": kind}
        if para_elem is not None:
            preview = _paragraph_preview(para_elem)
            if preview:
                entry["preview"] = preview
        bookmark_map.append(entry)

    if fresh_managed_document:
        diagnostics["fresh_document_fast_path"] = True

    controllers_locked = False
    try:
        try:
            doc.RecordChanges = False
        except Exception:
            pass
        if lock_controllers:
            doc.lockControllers()
            controllers_locked = True

        for kind, element in elements:
            try:
                if kind == "paragraph":
                    paragraph_count += 1
                    if fresh_managed_document:
                        try:
                            if not (element.getString() or "").strip():
                                diagnostics["top_level_empty_skipped_fast_path"] += 1
                                continue
                        except Exception:
                            pass
                    text = element.getText()
                    start = element.getStart()
                    if (
                        not fresh_managed_document
                        and _has_anchor_at_start(doc_text, start, managed_anchors)
                    ):
                        skipped_existing += 1
                        continue
                    cursor = text.createTextCursorByRange(start)
                    cursor.collapseToStart()
                else:
                    continue

                bookmark_name = _next_bookmark_name(existing_names, kind="paragraph")
                _create_bookmark(doc, text, cursor, bookmark_name)
                added += 1
                created_bookmarks.append(bookmark_name)
                if fresh_managed_document:
                    _append_bookmark_map_entry(bookmark_name, "paragraph", element)

                if not fresh_managed_document:
                    try:
                        managed_anchors.append(
                            (bookmark_name, bookmarks.getByName(bookmark_name).getAnchor())
                        )
                    except Exception:
                        pass
            except Exception as exc:
                failed += 1
                _record_error("top_level_paragraph", exc)
                continue

        _mark_timing("top_level_paragraphs")

        # Ensure table-level anchors and paragraph anchors inside each table cell.
        diagnostics["table_scan_started"] = True
        for table in _iter_text_tables(doc):
            table_count += 1

            # Table-level anchor (kept isolated so cell processing still runs on error)
            try:
                created = False
                if fresh_managed_document:
                    diagnostics["table_anchor_skipped_fast_path"] += 1
                    created = True

                if (
                    not fresh_managed_document
                    and _has_table_anchor_for_table(doc_text, table, managed_anchors)
                ):
                    skipped_existing += 1
                    created = True

                table_anchor = None
                table_anchor_error = None
                if not created:
                    try:
                        table_anchor = table.getAnchor()
                    except Exception as exc:
                        table_anchor_error = exc

                    if (
                        not fresh_managed_document
                        and table_anchor is not None
                        and _has_table_anchor_at_start(
                            doc_text, table_anchor, managed_anchors
                        )
                    ):
                        skipped_existing += 1
                        created = True

                if not created:
                    diagnostics["table_anchor_attempts"] += 1

                    # Pass 1: anchor at table start using any compatible
                    # text owner returned by UNO for this table anchor.
                    range_candidates = []
                    text_candidates = []
                    if table_anchor is not None:
                        try:
                            range_candidates.append(table_anchor)
                        except Exception:
                            pass
                        try:
                            range_candidates.append(table_anchor.getStart())
                        except Exception:
                            pass
                        try:
                            maybe_text = table_anchor.getText()
                            if maybe_text is not None:
                                text_candidates.append(maybe_text)
                        except Exception:
                            pass
                        try:
                            maybe_text = table_anchor.getStart().getText()
                            if maybe_text is not None:
                                text_candidates.append(maybe_text)
                        except Exception:
                            pass
                    text_candidates.append(doc_text)

                    seen_text_ids = set()
                    unique_text_candidates = []
                    for text_obj in text_candidates:
                        key = id(text_obj)
                        if key in seen_text_ids:
                            continue
                        seen_text_ids.add(key)
                        unique_text_candidates.append(text_obj)

                    for range_obj in range_candidates:
                        if created:
                            break
                        for text_obj in unique_text_candidates:
                            try:
                                cursor = _create_collapsed_cursor(text_obj, range_obj)
                                bookmark_name = _next_bookmark_name(
                                    existing_names, kind="table"
                                )
                                _create_bookmark(doc, text_obj, cursor, bookmark_name)
                                added += 1
                                created_bookmarks.append(bookmark_name)
                                diagnostics["table_anchor_added"] += 1
                                if fresh_managed_document:
                                    _append_bookmark_map_entry(bookmark_name, "table")
                                else:
                                    try:
                                        managed_anchors.append(
                                            (
                                                bookmark_name,
                                                bookmarks.getByName(bookmark_name).getAnchor(),
                                            )
                                        )
                                    except Exception:
                                        pass
                                created = True
                                break
                            except Exception:
                                continue

                    if not created:
                        anchor_hint = (
                            f" (table.getAnchor() failed: {table_anchor_error})"
                            if table_anchor_error is not None
                            else ""
                        )
                        raise RuntimeError(
                            "unable to create table-level bookmark" + anchor_hint
                        )
            except Exception as exc:
                failed += 1
                _record_error("table_level", exc)

            # Cell paragraph anchors
            cell_names = []
            try:
                cell_names = list(table.getCellNames())
            except Exception as exc:
                failed += 1
                _record_error("table_level", exc)

            for cell_name in cell_names:
                try:
                    cell = table.getCellByName(cell_name)
                except Exception as exc:
                    failed += 1
                    _record_error("table_cell_paragraph", exc)
                    continue

                try:
                    cell_text = cell.getText()
                    enum = cell_text.createEnumeration()
                except Exception as exc:
                    failed += 1
                    _record_error("table_cell_paragraph", exc)
                    continue

                try:
                    while enum.hasMoreElements():
                        elem = enum.nextElement()
                        try:
                            if not elem.supportsService("com.sun.star.text.Paragraph"):
                                continue
                        except Exception:
                            continue

                        table_cell_paragraph_count += 1
                        if fresh_managed_document:
                            try:
                                if not (elem.getString() or "").strip():
                                    diagnostics["table_cell_empty_skipped_fast_path"] += 1
                                    continue
                            except Exception:
                                pass
                        start = elem.getStart()
                        if (
                            not fresh_managed_document
                            and _has_non_table_anchor_at_start(
                                doc_text, start, managed_anchors
                            )
                        ):
                            skipped_existing += 1
                            continue

                        diagnostics["table_cell_paragraph_attempts"] += 1
                        p_text = elem.getText() or cell_text
                        p_cursor = p_text.createTextCursorByRange(start)
                        p_cursor.collapseToStart()

                        bookmark_name = _next_bookmark_name(
                            existing_names, kind="cell_paragraph"
                        )
                        _create_bookmark(doc, p_text, p_cursor, bookmark_name)
                        added += 1
                        created_bookmarks.append(bookmark_name)
                        diagnostics["table_cell_paragraph_added"] += 1
                        if fresh_managed_document:
                            _append_bookmark_map_entry(bookmark_name, "paragraph", elem)
                        else:
                            try:
                                managed_anchors.append(
                                    (
                                        bookmark_name,
                                        bookmarks.getByName(bookmark_name).getAnchor(),
                                    )
                                )
                            except Exception:
                                pass
                except Exception as exc:
                    failed += 1
                    _record_error("table_cell_paragraph", exc)
                    continue
        diagnostics["table_scan_completed"] = True
        _mark_timing("table_bookmarks")
    except Exception as exc:
        _log(f"ensureSdocBookmarks: failed with exception: {exc}\n{traceback.format_exc()}")
        return _result(
            False,
            error=str(exc),
            added=added,
            failed=failed,
            diagnostics=diagnostics,
        )
    finally:
        try:
            doc.RecordChanges = was_recording
        except Exception:
            pass
        if controllers_locked:
            try:
                doc.unlockControllers()
            except Exception:
                pass
    _mark_timing("bookmark_creation")

    # Build bookmark_map in document reading order.
    #
    # Why: the Bookmarks collection's getByIndex() order is unrelated to
    # the rendered document. If we iterate it directly, agents receive a
    # numbered list whose neighbors are NOT neighbors on the page, and any
    # follow-up "edit elements 3–5" reasoning targets the wrong spans.
    # Walking the body via createEnumeration() gives paragraphs + tables
    # in reading order, and for each element we look up the sdoc_* anchor
    # that points at its start.
    if fresh_managed_document:
        diagnostics["bookmark_map_fast_path"] = True
        _log(
            "ensureSdocBookmarks: bookmark_map collected "
            + f"{len(bookmark_map)} entries (fresh creation order)"
        )
    else:
        bookmark_map = []
        try:
            bookmarks = doc.getBookmarks()
            managed_anchors = _collect_managed_anchors(doc, bookmarks)

            def _paragraph_preview(para_elem: Any) -> str:
                try:
                    text_obj = para_elem.getText()
                    cursor = text_obj.createTextCursorByRange(para_elem.getStart())
                    cursor.gotoStartOfParagraph(False)
                    cursor.gotoEndOfParagraph(True)
                    raw_text = (cursor.getString() or "").strip()
                    words = raw_text.split()
                    if len(words) > 8:
                        preview = " ".join(words[:8]) + "..."
                    elif len(raw_text) > 80:
                        preview = raw_text[:80] + "..."
                    else:
                        preview = raw_text
                    return preview.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
                except Exception:
                    return ""

            def _paragraph_anchor_name(start: Any) -> Optional[str]:
                for nm, anchor in managed_anchors:
                    if _is_table_bookmark_name(nm):
                        continue
                    if _anchor_matches_start(doc_text, start, anchor):
                        return nm
                return None

            def _table_anchor_name(tbl: Any) -> Optional[str]:
                for nm, anchor in managed_anchors:
                    if not _is_table_bookmark_name(nm):
                        continue
                    if _anchor_matches_table_anchor(doc_text, tbl, anchor):
                        return nm
                return None

            def _append_paragraph_entry(para_elem: Any) -> None:
                try:
                    start = para_elem.getStart()
                except Exception:
                    return
                nm = _paragraph_anchor_name(start)
                if not nm:
                    return
                entry: Dict[str, Any] = {"id": nm, "kind": "paragraph"}
                preview = _paragraph_preview(para_elem)
                if preview:
                    entry["preview"] = preview
                bookmark_map.append(entry)

            for kind, element in _iter_body_elements(doc):
                if kind == "paragraph":
                    _append_paragraph_entry(element)
                elif kind == "table":
                    tnm = _table_anchor_name(element)
                    if tnm:
                        bookmark_map.append({"id": tnm, "kind": "table"})
                    try:
                        cell_names = sorted(list(element.getCellNames()))
                    except Exception:
                        cell_names = []
                    for cell_name in cell_names:
                        try:
                            cell = element.getCellByName(cell_name)
                            cell_enum = cell.getText().createEnumeration()
                        except Exception:
                            continue
                        while cell_enum.hasMoreElements():
                            try:
                                cp = cell_enum.nextElement()
                            except Exception:
                                break
                            try:
                                if not cp.supportsService("com.sun.star.text.Paragraph"):
                                    continue
                            except Exception:
                                continue
                            _append_paragraph_entry(cp)

            _log(
                "ensureSdocBookmarks: bookmark_map collected "
                + f"{len(bookmark_map)} entries (document order)"
            )
        except Exception as bm_exc:
            _log(f"ensureSdocBookmarks: failed to collect bookmark_map: {bm_exc}")
    _mark_timing("bookmark_map")

    _log(
        "ensureSdocBookmarks: "
        + f"reason={reason} request_id={request_id} elements={len(elements)} "
        + f"file_id={file_id} paragraphs={paragraph_count} tables={table_count} "
        + f"table_cell_paragraphs={table_cell_paragraph_count} "
        + f"managed_before={managed_count_before} skipped_existing={skipped_existing} "
        + f"added={added} failed={failed} "
        + f"preview_removed={preview_removed} preview_failed={preview_failed} "
        + f"preview_purge_skipped={skip_preview_purge} unresolved_preview_ops={unresolved_preview_ops} "
        + f"preview_green_remaining={preview_green_remaining} "
        + f"bookmark_map_count={len(bookmark_map)} "
        + f"diag_table_scan_started={diagnostics.get('table_scan_started')} "
        + f"diag_table_scan_completed={diagnostics.get('table_scan_completed')} "
        + f"diag_table_anchor_attempts={diagnostics.get('table_anchor_attempts')} "
        + f"diag_table_anchor_added={diagnostics.get('table_anchor_added')} "
        + f"diag_cell_attempts={diagnostics.get('table_cell_paragraph_attempts')} "
        + f"diag_cell_added={diagnostics.get('table_cell_paragraph_added')} "
        + f"diag_failure_buckets={diagnostics.get('failure_buckets')} "
        + f"timings={json.dumps(timings, sort_keys=True)}"
    )
    _log(
        "ensureSdocBookmarks: bookmark_map_sample="
        + json.dumps(bookmark_map[:25], ensure_ascii=False)
    )

    return _result(
        True,
        added=added,
        failed=failed,
        skipped_existing=skipped_existing,
        total_elements=len(elements),
        paragraph_elements=paragraph_count,
        table_elements=table_count,
        table_cell_paragraph_elements=table_cell_paragraph_count,
        managed_before=managed_count_before,
        managed_after=managed_count_before + added,
        preview_removed=preview_removed,
        preview_failed=preview_failed,
        preview_purge_skipped=skip_preview_purge,
        unresolved_preview_ops=unresolved_preview_ops,
        preview_green_remaining=preview_green_remaining,
        created_bookmarks=created_bookmarks,
        bookmark_map=bookmark_map,
        diagnostics=diagnostics,
    )


g_exportedScripts = (ensureSdocBookmarks,)
