"""
ApplyExcelActionPlan.py
-----------------------

Applies SmartDocs Excel action plans directly to the active Calc document.
Each operation mutates the live workbook (no preview layer) and positions
the cursor on the last affected cell/range to mirror in-editor focus.
"""

import json
import hashlib
import os
import re
import sys
import traceback
import math
import time
import ssl
import tempfile
import urllib.request
from urllib.parse import urlparse
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

import uno

from com.sun.star.awt.FontWeight import BOLD as FONT_WEIGHT_BOLD
from com.sun.star.awt.FontSlant import ITALIC as FONT_SLANT_ITALIC
from com.sun.star.awt.FontUnderline import SINGLE as FONT_UNDERLINE_SINGLE
from com.sun.star.table.CellHoriJustify import LEFT as HORI_LEFT
from com.sun.star.table.CellHoriJustify import CENTER as HORI_CENTER
from com.sun.star.table.CellHoriJustify import RIGHT as HORI_RIGHT
from com.sun.star.table.CellHoriJustify import STANDARD as HORI_STANDARD
from com.sun.star.table.CellVertJustify import TOP as VERT_TOP
from com.sun.star.table.CellVertJustify import CENTER as VERT_CENTER
from com.sun.star.table.CellVertJustify import BOTTOM as VERT_BOTTOM
from com.sun.star.table.CellContentType import EMPTY as CELL_CONTENT_EMPTY
from com.sun.star.table.CellContentType import TEXT as CELL_CONTENT_TEXT
from com.sun.star.table.CellContentType import VALUE as CELL_CONTENT_VALUE
from com.sun.star.table.CellContentType import FORMULA as CELL_CONTENT_FORMULA
from com.sun.star.sheet.CellInsertMode import ROWS as INSERT_ROWS
from com.sun.star.sheet.CellInsertMode import DOWN as INSERT_DOWN

_CHART_LAYOUT_STATE: Optional[Dict[str, Any]] = None
_DEFAULT_CHART_START_CELL = "H2"
_CHART_GRID_COLUMNS = 2
_CHART_COLUMN_STRIDE = 7  # number of columns each chart slot spans (wider charts)
_CHART_ROW_STRIDE = 18  # number of rows each chart slot spans (taller charts)
_CHART_ROW_PADDING = 3  # extra empty rows between data and first slot
_CHART_COLUMN_PADDING = 2  # extra empty columns between data and first slot
_CHART_ROW_GAP = 2  # blank rows between chart slots (minimum 2 for readability)
_CHART_COLUMN_GAP = 2  # blank columns between chart slots (minimum 2 for readability)
_ACTIVE_CHART_DOC_ID: Optional[str] = None
_IMAGE_MAX_COLUMN_SCAN = 24
_IMAGE_MAX_ROW_SCAN = 80
_IMAGE_ROW_SLOT_STRIDE = 8
# Dispatch context for .uno:Color (set by applyExcelActionPlan at entry).
# Using .uno:Color dispatch instead of setPropertyValue("CharColor", ...) is
# critical: it properly clears the internal ComplexColor theme reference that
# the XLSX exporter reads.  setPropertyValue does NOT clear it.
_DISPATCH_CONTROLLER: Any = None
_DISPATCH_FRAME: Any = None
# When False, formulas may reference empty cells (treated as zero in Calc).
# This prevents plan execution from failing during staged model builds.
_STRICT_FORMULA_REFERENCE_CHECK = False
# Calc's `OptimalWidth` is often tighter than Excel's AutoFit; add a small fixed
# width after autofit so text doesn't visually touch borders.
_DEFAULT_AUTOFIT_PADDING_MM100 = 200  # 2mm in 1/100th-mm units
_ACCOUNTING_AUTOFIT_EXTRA_CHARS = 3

# Autofit ceiling: at 11pt Calibri, ~1.85mm/char. 9500 in 1/100mm ≈ 51 chars.
# This prevents a single long-description cell from stretching a column to
# half a page wide. Long cells get a partial preview; the user sees the first
# ~50 chars and clicks the cell to read the rest. Applied in both the UNO
# autofit path and the openpyxl path (services/excel/tools.py).
_MAX_AUTOFIT_WIDTH_MM100 = 9500
# Important: keep script-managed store disabled and rely on host-managed Action_Save/WOPI.
# This avoids lock/permission conflicts when both script and host try to persist at once.
_ENABLE_SCRIPT_SIDE_STORE = False
# Excel styling convention defaults (kept aligned with backend/prompts/agents/excel_agent.md)
_DEFAULT_TABLE_HEADER_BG_COLOR = 0x203864  # Dark Navy
_DEFAULT_TABLE_HEADER_FONT_COLOR = 0xFFFFFF  # White
_DEFAULT_SECTION_HEADER_BG_COLOR = 0xD9D9D9  # Light Gray
_DEFAULT_TABLE_BORDER_COLOR = 0xC0C0C0  # Light gray grid
_PLAIN_NUMERIC_TEXT_RE = re.compile(r"^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$")
_DUPLICATE_HEADER_WRITE_SKIPS: List[Dict[str, Any]] = []

_UNVERIFIED_SSL_CONTEXT = ssl.create_default_context()
try:
    _UNVERIFIED_SSL_CONTEXT.check_hostname = False
    _UNVERIFIED_SSL_CONTEXT.verify_mode = ssl.CERT_NONE
except Exception:
    pass


class ActionApplicationError(RuntimeError):
    """Action-level failure with structured details for agent replanning."""

    def __init__(
        self,
        message: str,
        *,
        error_code: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None,
    ):
        super().__init__(message)
        self.error_code = error_code
        self.details = details or {}


def _env_flag(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return bool(default)
    return str(raw).strip().lower() in {"1", "true", "yes", "on"}


_SNAPSHOT_MAX_CELLS = max(64, int(os.environ.get("SMARTDOCS_EXCEL_SNAPSHOT_MAX_CELLS", "4000")))
_SNAPSHOT_FORMULA_MAX_CELLS = max(
    16, int(os.environ.get("SMARTDOCS_EXCEL_SNAPSHOT_FORMULA_MAX_CELLS", "1200"))
)
_OPTIMIZE_SNAPSHOT_CAPTURE = _env_flag(
    "SMARTDOCS_EXCEL_OPTIMIZE_SNAPSHOT_CAPTURE", True
)
_OPTIMIZE_POSTPROCESS = _env_flag("SMARTDOCS_EXCEL_OPTIMIZE_POSTPROCESS", True)

# Action kinds that only affect display/layout — they do not change cell
# values or formulas. When a plan contains ONLY these kinds, we can skip
# the value-dependent post-processing passes (calculateAll, thousands /
# currency number-format pass, autofit) which are O(touched_sheets ×
# used_cells) and dominate macro runtime as the workbook grows.
_FORMAT_ONLY_ACTION_KINDS = frozenset({
    "apply_formatting",
    "apply_conditional_formatting",
    "set_column_width",
    "set_row_height",
    "autofit_columns",
    "set_freeze_panes",
    "hide_rows",
    "show_rows",
    "hide_columns",
    "show_columns",
    "protect_sheet",
    "apply_filter",
})


def _bootstrap_extra_python_paths() -> None:
    """
    Allow optional third-party libs for UNO scripts (e.g. httpx).
    Priority:
    1) SMARTDOCS_LO_PYTHONPATH (os.pathsep-delimited)
    2) backend/services/libreoffice/vendor under repo root (best-effort)
    """
    candidates: List[str] = []

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


def _log(message: str) -> None:
    try:
        print(message, flush=True)
    except Exception:
        pass


def _set_user_profile_for_author(author_name: str = "SmartDocs AI") -> bool:
    """
    Set the LibreOffice user profile name, which is used as the Author for annotations.
    This modifies /org.openoffice.UserProfile/Data configuration.
    
    Returns True if successful, False otherwise.
    """
    try:
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        
        # Get ConfigurationProvider
        config_provider = smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx
        )
        
        # Get writable access to user profile data
        props = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        props.Name = "nodepath"
        props.Value = "/org.openoffice.UserProfile/Data"
        
        config_access = config_provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", (props,)
        )
        
        # Set the user name fields
        config_access.setPropertyValue("givenname", author_name)
        config_access.setPropertyValue("sn", "")  # surname/last name
        config_access.setPropertyValue("initials", "AI")
        
        # Commit the changes
        config_access.commitChanges()
        
        _log(f"_set_user_profile_for_author: set user profile to '{author_name}'")
        return True
        
    except Exception as e:
        _log(f"_set_user_profile_for_author: failed - {e}")
        return False


def _document_identifier(document) -> str:
    """Return a stable identifier for the current UNO document."""
    for attr in ("getURL", "getTitle"):
        getter = getattr(document, attr, None)
        if callable(getter):
            try:
                value = getter()
            except Exception:
                continue
            if isinstance(value, str) and value:
                return value
    return f"doc-{id(document)}"


def _has_prop(obj, name: str) -> bool:
    """Return True if obj exposes a property by name via XPropertySet info."""
    if obj is None or not name:
        return False
    try:
        info = getattr(obj, "getPropertySetInfo", None)
        psi = info() if callable(info) else None
        if psi is None:
            return False
        has_by_name = getattr(psi, "hasPropertyByName", None)
        if callable(has_by_name):
            return bool(has_by_name(name))
    except Exception:
        return False
    return False


def _range_address_cell_count(address) -> int:
    try:
        cols = int(address.EndColumn) - int(address.StartColumn) + 1
        rows = int(address.EndRow) - int(address.StartRow) + 1
    except Exception:
        return 0
    if cols <= 0 or rows <= 0:
        return 0
    return cols * rows


def _clip_range_address_for_snapshot(sheet, address, max_cells: int):
    """
    Cap very large snapshot ranges to keep UNO script latency bounded.
    Preference order:
      1) Used-area intersection for the target range.
      2) Top-left crop to max_cells.
    """
    original_count = _range_address_cell_count(address)
    if original_count <= max_cells:
        return address, False, original_count

    start_col = int(address.StartColumn)
    end_col = int(address.EndColumn)
    start_row = int(address.StartRow)
    end_row = int(address.EndRow)
    cols = max(1, end_col - start_col + 1)

    used = _sheet_used_area(sheet)
    if used is not None:
        try:
            used_start_col = int(used.StartColumn)
            used_end_col = int(used.EndColumn)
            used_start_row = int(used.StartRow)
            used_end_row = int(used.EndRow)
            start_col = max(start_col, used_start_col)
            end_col = min(end_col, used_end_col)
            start_row = max(start_row, used_start_row)
            end_row = min(end_row, used_end_row)
            if end_col >= start_col and end_row >= start_row:
                cols = max(1, end_col - start_col + 1)
            else:
                start_col = int(address.StartColumn)
                end_col = int(address.EndColumn)
                start_row = int(address.StartRow)
                end_row = int(address.EndRow)
                cols = max(1, end_col - start_col + 1)
        except Exception:
            start_col = int(address.StartColumn)
            end_col = int(address.EndColumn)
            start_row = int(address.StartRow)
            end_row = int(address.EndRow)
            cols = max(1, end_col - start_col + 1)

    max_rows = max(1, max_cells // cols)
    end_row = min(end_row, start_row + max_rows - 1)
    if end_row < start_row:
        end_row = start_row
    if end_col < start_col:
        end_col = start_col

    clipped = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    clipped.Sheet = int(getattr(address, "Sheet", 0))
    clipped.StartColumn = start_col
    clipped.EndColumn = end_col
    clipped.StartRow = start_row
    clipped.EndRow = end_row

    clipped_count = _range_address_cell_count(clipped)
    return clipped, True, clipped_count


def _range_address_to_ref(address) -> Optional[str]:
    try:
        start = f"{_column_index_to_name(int(address.StartColumn))}{int(address.StartRow) + 1}"
        end = f"{_column_index_to_name(int(address.EndColumn))}{int(address.EndRow) + 1}"
    except Exception:
        return None
    if start == end:
        return start
    return f"{start}:{end}"


def _action_payload_contains_formula(action: Dict[str, Any], max_items: int = 5000) -> bool:
    """
    Fast heuristic: detect if action payload includes formula-like strings.
    Used to gate expensive post-write formula diagnostics.
    """
    payload = action.get("payload")
    if payload is None:
        return False

    stack: List[Any] = [payload]
    visited = 0
    while stack and visited < max_items:
        current = stack.pop()
        visited += 1
        if isinstance(current, str):
            if current.strip().startswith("="):
                return True
            continue
        if isinstance(current, dict):
            stack.extend(current.values())
            continue
        if isinstance(current, (list, tuple)):
            stack.extend(current)
            continue
    return False


def _capture_range_snapshot(
    document, action: Dict[str, Any]
) -> Optional[Dict[str, Any]]:
    """
    Capture the current state of a cell range before mutation.
    Used for undo/revert functionality.

    Returns:
        Dict containing original values, formulas, and range info
    """
    target = action.get("target") or {}
    sheet_name = target.get("sheet_name")
    range_name = target.get("range")

    if not range_name:
        # Check payload for range info
        payload = action.get("payload") or {}
        range_name = payload.get("start_cell") or payload.get("range") or payload.get("target_range")
        sheet_name = sheet_name or payload.get("sheet_name")

    if not range_name:
        return None

    try:
        sheet, range_name = _resolve_sheet_and_range_name(
            document, sheet_name, range_name, "_capture_range_snapshot"
        )
        if sheet is None:
            return None

        cell_range = sheet.getCellRangeByName(range_name)
        if cell_range is None:
            return None

        # Capture current state (capped for very large ranges to avoid timeouts).
        address = cell_range.getRangeAddress()
        original_cell_count = _range_address_cell_count(address)
        if _OPTIMIZE_SNAPSHOT_CAPTURE:
            capture_address, truncated, captured_cell_count = _clip_range_address_for_snapshot(
                sheet, address, _SNAPSHOT_MAX_CELLS
            )
        else:
            capture_address = address
            truncated = False
            captured_cell_count = original_cell_count

        capture_range = cell_range
        if truncated:
            capture_range = sheet.getCellRangeByPosition(
                int(capture_address.StartColumn),
                int(capture_address.StartRow),
                int(capture_address.EndColumn),
                int(capture_address.EndRow),
            )

        values = list(map(list, capture_range.getDataArray()))
        formulas = []

        # Formula snapshots are expensive; cap them separately in optimized mode.
        if not _OPTIMIZE_SNAPSHOT_CAPTURE or captured_cell_count <= _SNAPSHOT_FORMULA_MAX_CELLS:
            try:
                formulas = list(map(list, capture_range.getFormulaArray()))
            except Exception:
                pass

        snapshot = {
            "kind": action.get("kind"),
            "target": {
                "sheet_name": sheet_name or sheet.getName(),
                "range": range_name,
                "start_row": address.StartRow,
                "start_col": address.StartColumn,
                "end_row": address.EndRow,
                "end_col": address.EndColumn,
            },
            "payload": {
                "original_values": values,
                "original_formulas": formulas if formulas else None,
                "snapshot_cell_count": captured_cell_count,
            },
        }
        if truncated:
            snapshot["payload"]["snapshot_truncated"] = True
            snapshot["payload"]["snapshot_original_cell_count"] = original_cell_count
            snapshot["payload"]["snapshot_captured_range"] = _range_address_to_ref(
                capture_address
            )
            _log(
                "_capture_range_snapshot: truncated %s from %s cells to %s cells "
                "captured_range=%s"
                % (
                    range_name,
                    original_cell_count,
                    captured_cell_count,
                    snapshot["payload"].get("snapshot_captured_range"),
                )
            )
        _log(f"_capture_range_snapshot: captured {range_name}")
        return snapshot

    except Exception as e:
        _log(f"_capture_range_snapshot: failed - {e}")
        return None


def _add_range_comment(
    sheet,
    range_ref: str,
    text: str,
    author: str = "AI Assistant",
) -> bool:
    """
    Add a comment (annotation) to a cell explaining AI reasoning.

    This allows users to see why the AI made specific changes.

    Args:
        sheet: The sheet object
        range_ref: Cell reference (e.g., "A1")
        text: Comment text explaining the reasoning
        author: Author name for the comment

    Returns:
        True if comment was added successfully
    """
    try:
        # Set user profile so annotation Author shows our brand name
        _set_user_profile_for_author(author)
        
        cell = sheet.getCellRangeByName(range_ref)
        if cell is None:
            return False

        # For ranges, get top-left cell
        try:
            address = cell.getRangeAddress()
            target_cell = sheet.getCellByPosition(address.StartColumn, address.StartRow)
        except Exception:
            target_cell = cell

        # Get annotation - LibreOffice returns an annotation object even if empty
        annotation = target_cell.getAnnotation()
        current_text = ""
        try:
            current_text = annotation.getString() if annotation else ""
        except Exception:
            current_text = ""

        # Build comment text (author is set via user profile, not in text)
        full_text = text
        if current_text:
            full_text = f"{current_text}\n---\n{text}"

        if current_text:
            # Update existing annotation
            annotation.setString(full_text)
        else:
            # No existing content - create new annotation
            # Author is set via _set_user_profile_for_author() called earlier
            annotations = sheet.getAnnotations()
            cell_address = target_cell.getCellAddress()
            annotations.insertNew(cell_address, full_text)

        _log(f"_add_range_comment: added comment to {range_ref}")
        return True

    except Exception as e:
        _log(f"_add_range_comment: failed - {e}")
        return False


def applyExcelActionPlan(plan_json: str) -> str:
    """
    Entry point executed by Collabora Calc. Expects a JSON encoded ExcelActionPlan.
    """
    script_received_ts_ms = int(time.time() * 1000)
    final_store_start_ts_ms: Optional[int] = None
    final_store_end_ts_ms: Optional[int] = None
    # High-resolution timing accumulators for diagnostic logging.
    # Lines tagged "[TIME]" are easy to grep for after a slow run; the final
    # "[TIME-SUMMARY]" line gives an at-a-glance breakdown without grep.
    _plan_perf_t0 = time.perf_counter()
    phase_timings: Dict[str, int] = {}
    action_timings: List[Dict[str, Any]] = []
    _log("ApplyExcelActionPlan: received payload")
    try:
        plan = json.loads(plan_json)
        _log(
            "ApplyExcelActionPlan: parsed plan plan_id=%s actions=%s"
            % (plan.get("plan_id"), len(plan.get("actions", [])))
        )
    except Exception as exc:
        _log(
            f"ApplyExcelActionPlan: failed to parse JSON: {exc}\n{traceback.format_exc()}"
        )
        # Return a structured error so host can parse it
        return json.dumps(
            {
                "ok": False,
                "error": "invalid_json",
                "message": str(exc),
            }
        )

    try:
        document = XSCRIPTCONTEXT.getDocument()  # type: ignore  # NOQA
    except Exception as exc:
        _log(
            f"ApplyExcelActionPlan: unable to obtain document: {exc}\n{traceback.format_exc()}"
        )
        return json.dumps(
            {"ok": False, "plan_id": plan.get("plan_id"), "error": "no_document"}
        )

    if document is None:
        _log("ApplyExcelActionPlan: XSCRIPTCONTEXT returned None document")
        return json.dumps(
            {"ok": False, "plan_id": plan.get("plan_id"), "error": "no_controller"}
        )
    doc_identifier = _document_identifier(document)

    controller = None
    try:
        controller = document.getCurrentController()
    except Exception as exc:
        _log(
            f"ApplyExcelActionPlan: getCurrentController failed: {exc}\n{traceback.format_exc()}"
        )
        return json.dumps(
            {
                "ok": False,
                "plan_id": plan.get("plan_id"),
                "error": "controller_exception",
                "message": str(exc),
            }
        )

    if controller is None:
        _log("ApplyExcelActionPlan: controller is None")
        return json.dumps(
            {"ok": False, "plan_id": plan.get("plan_id"), "error": "no_controller"}
        )

    # Stash dispatch context so _set_explicit_char_color can use .uno:Color.
    global _DISPATCH_CONTROLLER, _DISPATCH_FRAME
    _DISPATCH_CONTROLLER = controller
    try:
        _DISPATCH_FRAME = controller.getFrame()
    except Exception:
        _DISPATCH_FRAME = None

    actions = plan.get("actions") or []
    if not isinstance(actions, Iterable):
        _log("ApplyExcelActionPlan: plan.actions is not iterable")
        return json.dumps(
            {"ok": False, "plan_id": plan.get("plan_id"), "error": "invalid_actions"}
        )
    global _CHART_LAYOUT_STATE, _ACTIVE_CHART_DOC_ID
    _DUPLICATE_HEADER_WRITE_SKIPS.clear()
    if doc_identifier != _ACTIVE_CHART_DOC_ID:
        _CHART_LAYOUT_STATE = None
        _ACTIVE_CHART_DOC_ID = doc_identifier
    if _CHART_LAYOUT_STATE is None:
        _CHART_LAYOUT_STATE = {"plan_id": plan.get("plan_id"), "per_sheet": {}}
    else:
        _CHART_LAYOUT_STATE["plan_id"] = plan.get("plan_id")
        _CHART_LAYOUT_STATE.setdefault("per_sheet", {})

    # Settings: Forced to False to prevent Document Override warnings in Collabora.
    # We now handle a single final save at the end of applyExcelActionPlan.
    save_after_each = False

    # Minimal set of actions that mutate the document and should trigger a save when enabled
    WRITE_KINDS = {
        "create_sheet",
        "delete_sheet",
        "write_range",
        "set_cell_formula",
        "append_rows",
        "create_column",
        "create_conditional_column",
        "delete_rows",
        "delete_columns",
        "find_and_replace",
        "manipulate_text_column",
        "update_cells",
        "apply_formatting",
        "apply_conditional_formatting",
        "autofit_columns",
        "sort_range",
        "create_chart",
        "delete_chart",
        "add_title",
        "insert_comment",
        # New operations
        "set_freeze_panes",
        "add_data_validation",
        "protect_sheet",
        "insert_image",
        "add_hyperlink",
        "apply_filter",
        "hide_rows",
        "show_rows",
        "hide_columns",
        "show_columns",
        "set_row_height",
        "set_column_width",
        "merge_cells",
    }

    results = []
    postprocess_sheets = set()
    postprocess_ranges_by_sheet: Dict[str, List[tuple[int, int, int, int]]] = {}
    postprocess_required = False
    # Columns explicitly sized by a set_column_width action. Keyed by
    # sheet_name → set of 0-based column indexes. Post-plan autofit skips
    # these specific columns so the agent's intent is respected, while the
    # rest of the sheet still gets autofit. (Previously a global boolean
    # skipped autofit for the ENTIRE sheet if any set_column_width ran —
    # making most columns miss their autofit pass.)
    explicitly_sized_columns: Dict[str, set] = {}
    needs_formula_diagnostics = False
    # True once any non-format-only action has been applied. When False at
    # the end of the loop, every value-dependent post-processing pass can
    # be skipped because no cell value or formula has changed.
    content_change_detected = False
    deferred_deletes = []  # delete_sheet actions deferred to end of plan

    controllers_locked = False
    try:
        document.lockControllers()
        controllers_locked = True
        _log("ApplyExcelActionPlan: controllers locked")
    except Exception as exc:
        _log(
            f"ApplyExcelActionPlan: failed to lock controllers: {exc}\n{traceback.format_exc()}"
        )
        return json.dumps(
            {
                "ok": False,
                "plan_id": plan.get("plan_id"),
                "error": "lock_failed",
                "message": str(exc),
            }
        )

    try:
        for index, action in enumerate(actions):
            if not isinstance(action, dict):
                _log(f"ApplyExcelActionPlan: skipping action index {index}; not a dict")
                results.append(
                    {
                        "index": index,
                        "action_name": action.get("action_name"),
                        "status": "skipped",
                        "reason": "not_a_dict",
                    }
                )
                continue
            kind = action.get("kind")
            # Defer delete_sheet to after all other actions so that actions
            # referencing the default sheet (e.g. Sheet1) can still succeed.
            if kind == "delete_sheet":
                _log(f"ApplyExcelActionPlan: deferring action[{index}] kind=delete_sheet to end of plan")
                deferred_deletes.append((index, action))
                continue
            action.pop("_result_meta", None)
            handler = _ACTION_HANDLERS.get(kind)
            _log(f"ApplyExcelActionPlan: applying action[{index}] kind={kind}")
            if handler is None:
                _log(f"ApplyExcelActionPlan: unsupported action kind '{kind}'")
                results.append(
                    {
                        "index": index,
                        "operation_id": action.get("operation_id"),
                        "action_name": action.get("action_name"),
                        "kind": kind,
                        "status": "error",
                        "error": "unsupported_action",
                    }
                )
                continue
            _action_perf_t0 = time.perf_counter()
            _snapshot_ms = 0
            _handler_ms = 0
            try:
                # Capture snapshot BEFORE mutation for undo support
                _snap_t0 = time.perf_counter()
                snapshot = None
                if kind in WRITE_KINDS:
                    snapshot = _capture_range_snapshot(document, action)
                    if snapshot:
                        _log(
                            f"ApplyExcelActionPlan: captured snapshot for op "
                            f"{action.get('operation_id')}"
                        )
                _snapshot_ms = int((time.perf_counter() - _snap_t0) * 1000)

                # Add reasoning comment if provided
                reasoning = (action.get("metadata") or {}).get("reasoning")
                if reasoning:
                    target = action.get("target") or {}
                    range_ref = target.get("range")
                    sheet_name = target.get("sheet_name")
                    if range_ref:
                        sheet = _get_sheet(document, sheet_name)
                        if sheet:
                            _add_range_comment(sheet, range_ref, reasoning)

                _handler_t0 = time.perf_counter()
                target_range = handler(document, controller, action)
                _handler_ms = int((time.perf_counter() - _handler_t0) * 1000)
                # Postprocess after every successful mutating action.
                # This keeps thousands-format enforcement resilient when new action
                # kinds are introduced and WRITE_KINDS is not updated in lockstep.
                if _OPTIMIZE_POSTPROCESS:
                    if isinstance(kind, str) and kind != "set_active_sheet":
                        postprocess_required = True
                        if kind not in _FORMAT_ONLY_ACTION_KINDS:
                            content_change_detected = True
                        sheet_name = _resolve_postprocess_sheet_name(
                            document, controller, action, target_range
                        )
                        if sheet_name:
                            postprocess_sheets.add(sheet_name)
                            _record_postprocess_range(
                                postprocess_ranges_by_sheet, sheet_name, target_range
                            )
                        if kind == "set_column_width":
                            # Remember the specific columns that were explicitly
                            # sized so post-plan autofit can skip just those.
                            try:
                                cols = _columns_from_action_range(
                                    document, action, default_sheet=sheet_name
                                )
                                for (sname, col_idx) in cols:
                                    explicitly_sized_columns.setdefault(
                                        sname, set()
                                    ).add(col_idx)
                            except Exception as exc:
                                _log(
                                    f"ApplyExcelActionPlan: could not record explicit "
                                    f"column widths from set_column_width: {exc}"
                                )
                        if kind in {"set_cell_formula", "create_column", "create_conditional_column"}:
                            needs_formula_diagnostics = True
                        elif kind in {
                            "write_range",
                            "append_rows",
                            "update_cells",
                            "manipulate_text_column",
                        } and _action_payload_contains_formula(action):
                            needs_formula_diagnostics = True
                else:
                    if isinstance(kind, str) and (kind in WRITE_KINDS or kind == "set_active_sheet"):
                        postprocess_required = True
                        if kind not in _FORMAT_ONLY_ACTION_KINDS and kind != "set_active_sheet":
                            content_change_detected = True
                        sheet_name = _resolve_postprocess_sheet_name(
                            document, controller, action, target_range
                        )
                        if sheet_name:
                            postprocess_sheets.add(sheet_name)
                            _record_postprocess_range(
                                postprocess_ranges_by_sheet, sheet_name, target_range
                            )
                cursor_ref = (
                    _resolve_range(document, action.get("cursor_target"))
                    or target_range
                )
                if cursor_ref is not None:
                    try:
                        controller.select(cursor_ref)
                    except Exception as select_exc:
                        _log(
                            f"ApplyExcelActionPlan: controller.select failed: {select_exc}"
                        )
                # Save (store) after each mutating action if enabled
                if save_after_each and isinstance(kind, str) and kind in WRITE_KINDS:
                    try:
                        if _ENABLE_SCRIPT_SIDE_STORE:
                            # Most documents implement com.sun.star.frame.XStorable
                            document.store()
                            _log(f"ApplyExcelActionPlan: document.store() after '{kind}'")
                        else:
                            _log(
                                f"ApplyExcelActionPlan: skipped script-side document.store() after '{kind}' (host-managed save)"
                            )
                    except Exception as store_exc:
                        _log(
                            f"ApplyExcelActionPlan: document.store failed: {store_exc}"
                        )
                # Record success with snapshot for undo
                result_entry = {
                    "index": index,
                    "operation_id": action.get("operation_id"),
                    "action_name": action.get("action_name"),
                    "kind": kind,
                    "status": "ok",
                }
                result_entry.update(
                    _describe_applied_target(document, controller, action, target_range)
                )
                action_meta = action.get("_result_meta")
                if isinstance(action_meta, dict) and action_meta:
                    result_entry["applied_details"] = action_meta
                if snapshot:
                    result_entry["snapshot"] = snapshot
                results.append(result_entry)
            except Exception as handler_exc:
                _log(
                    f"ApplyExcelActionPlan: handler failed for {kind}: {handler_exc}\n"
                    f"{traceback.format_exc()}"
                )
                error_entry = {
                    "index": index,
                    "operation_id": action.get("operation_id"),
                    "action_name": action.get("action_name"),
                    "kind": kind,
                    "status": "error",
                    "error": str(handler_exc),
                }
                error_code = getattr(handler_exc, "error_code", None)
                if error_code:
                    error_entry["error_code"] = error_code
                error_details = getattr(handler_exc, "details", None)
                if isinstance(error_details, dict) and error_details:
                    error_entry["error_details"] = error_details
                results.append(
                    error_entry
                )
            _action_total_ms = int(
                (time.perf_counter() - _action_perf_t0) * 1000
            )
            action_timings.append(
                {
                    "index": index,
                    "kind": str(kind),
                    "ms": _action_total_ms,
                    "handler_ms": _handler_ms,
                    "snapshot_ms": _snapshot_ms,
                }
            )
            _log(
                f"ApplyExcelActionPlan[TIME]: action[{index}] kind={kind} "
                f"ms={_action_total_ms} handler_ms={_handler_ms} "
                f"snapshot_ms={_snapshot_ms}"
            )

        # Execute deferred delete_sheet actions now that all other actions
        # have finished. Only delete sheets that are completely empty.
        for del_index, del_action in deferred_deletes:
            del_target = del_action.get("target") or del_action.get("payload") or {}
            del_sheet_name = del_target.get("sheet_name")
            if not del_sheet_name:
                _log(f"ApplyExcelActionPlan: skipping deferred delete[{del_index}] — no sheet_name")
                continue
            del_sheet = _get_sheet(document, del_sheet_name)
            if del_sheet is None:
                _log(f"ApplyExcelActionPlan: skipping deferred delete[{del_index}] — sheet '{del_sheet_name}' already gone")
                results.append({"index": del_index, "operation_id": del_action.get("operation_id"), "action_name": del_action.get("action_name"), "kind": "delete_sheet", "status": "skipped", "reason": "already_deleted"})
                continue
            # Only delete if the sheet is empty (no content in used area)
            try:
                cursor = del_sheet.createCursor()
                cursor.gotoEndOfUsedArea(True)
                addr = cursor.getRangeAddress()
                has_data = addr.EndColumn > 0 or addr.EndRow > 0
                if not has_data:
                    # Double-check the single cell A1
                    cell_a1 = del_sheet.getCellByPosition(0, 0)
                    has_data = not _cell_is_effectively_empty(cell_a1)
            except Exception:
                has_data = True
            if has_data:
                _log(f"ApplyExcelActionPlan: skipping deferred delete[{del_index}] — sheet '{del_sheet_name}' has data")
                results.append({"index": del_index, "operation_id": del_action.get("operation_id"), "action_name": del_action.get("action_name"), "kind": "delete_sheet", "status": "skipped", "reason": "sheet_has_data"})
                continue
            try:
                handler = _ACTION_HANDLERS.get("delete_sheet")
                if handler:
                    _log(f"ApplyExcelActionPlan: executing deferred delete[{del_index}] kind=delete_sheet for '{del_sheet_name}'")
                    handler(document, controller, del_action)
                    results.append({"index": del_index, "operation_id": del_action.get("operation_id"), "action_name": del_action.get("action_name"), "kind": "delete_sheet", "status": "ok"})
            except Exception as del_exc:
                _log(f"ApplyExcelActionPlan: deferred delete[{del_index}] failed: {del_exc}")
                results.append({"index": del_index, "operation_id": del_action.get("operation_id"), "action_name": del_action.get("action_name"), "kind": "delete_sheet", "status": "error", "error": str(del_exc)})

        if _OPTIMIZE_POSTPROCESS:
            formula_diagnostics_requested = bool(
                needs_formula_diagnostics or _FORMULA_RECALC_PENDING
            )
        else:
            formula_diagnostics_requested = bool(postprocess_required)

        # When the plan contained ONLY format-only actions, skip the
        # value-dependent post-processing passes. Those passes iterate the
        # used area of every touched sheet and dominate runtime once the
        # workbook grows; they only matter when cell values or formulas
        # actually changed.
        skip_value_dependent_passes = (
            postprocess_required
            and not content_change_detected
            and not _FORMULA_RECALC_PENDING
        )

        style_postprocess_sheet_targets = _resolve_postprocess_sheet_targets(
            document,
            postprocess_sheets,
            include_all_when_empty=postprocess_required and not postprocess_sheets,
        )

        if skip_value_dependent_passes:
            _log(
                "ApplyExcelActionPlan: skipping value-dependent post-processing "
                "(formatting-only plan, no content changes)"
            )
            postprocess_sheet_targets: Tuple[str, ...] = ()
        else:
            _phase_t0 = time.perf_counter()
            _maybe_calculate_formulas(
                document, controller, force=postprocess_required
            )
            phase_timings["calculate_formulas"] = int(
                (time.perf_counter() - _phase_t0) * 1000
            )
            _log(
                f"ApplyExcelActionPlan[TIME]: phase=calculate_formulas "
                f"ms={phase_timings['calculate_formulas']}"
            )

            postprocess_sheet_targets = style_postprocess_sheet_targets
            _sheet_count = (
                len(postprocess_sheet_targets) if postprocess_sheet_targets else 0
            )

            _phase_t0 = time.perf_counter()
            _maybe_apply_thousands_separator_format(document, postprocess_sheet_targets)
            phase_timings["thousands_format"] = int(
                (time.perf_counter() - _phase_t0) * 1000
            )
            _log(
                f"ApplyExcelActionPlan[TIME]: phase=thousands_format "
                f"ms={phase_timings['thousands_format']} sheets={_sheet_count}"
            )

            _phase_t0 = time.perf_counter()
            _maybe_apply_currency_separator_format(document, postprocess_sheet_targets)
            phase_timings["currency_format"] = int(
                (time.perf_counter() - _phase_t0) * 1000
            )
            _log(
                f"ApplyExcelActionPlan[TIME]: phase=currency_format "
                f"ms={phase_timings['currency_format']} sheets={_sheet_count}"
            )

            _phase_t0 = time.perf_counter()
            _maybe_autofit_after_calculate(
                document,
                controller,
                postprocess_sheet_targets,
                excluded_columns_by_sheet=explicitly_sized_columns,
                ranges_by_sheet=postprocess_ranges_by_sheet,
            )
            phase_timings["autofit"] = int(
                (time.perf_counter() - _phase_t0) * 1000
            )
            _log(
                f"ApplyExcelActionPlan[TIME]: phase=autofit "
                f"ms={phase_timings['autofit']} sheets={_sheet_count}"
            )

        _phase_t0 = time.perf_counter()
        _maybe_repair_dark_fill_font_contrast(
            document,
            style_postprocess_sheet_targets,
            ranges_by_sheet=postprocess_ranges_by_sheet,
        )
        phase_timings["dark_fill_font_contrast"] = int(
            (time.perf_counter() - _phase_t0) * 1000
        )
        _log(
            f"ApplyExcelActionPlan[TIME]: phase=dark_fill_font_contrast "
            f"ms={phase_timings['dark_fill_font_contrast']} "
            f"sheets={len(style_postprocess_sheet_targets) if style_postprocess_sheet_targets else 0}"
        )
        if formula_diagnostics_requested:
            _phase_t0 = time.perf_counter()
            if _OPTIMIZE_POSTPROCESS:
                formula_errors = _scan_formula_errors(document, postprocess_sheet_targets)
            else:
                formula_errors = _scan_formula_errors(document)
            phase_timings["scan_formula_errors"] = int(
                (time.perf_counter() - _phase_t0) * 1000
            )
            _log(
                f"ApplyExcelActionPlan[TIME]: phase=scan_formula_errors "
                f"ms={phase_timings['scan_formula_errors']} "
                f"errors={int(formula_errors.get('total') or 0)}"
            )
            if formula_errors.get("total", 0):
                total_errors = int(formula_errors.get("total") or 0)
                by_type = formula_errors.get("by_type") or {}
                sample_items = formula_errors.get("sample") or []
                sample_refs: List[str] = []
                for sample_entry in sample_items[:5]:
                    if not isinstance(sample_entry, dict):
                        continue
                    sheet = sample_entry.get("sheet_name") or ""
                    cell = sample_entry.get("cell") or ""
                    err = sample_entry.get("error") or "ERROR"
                    if cell:
                        sample_refs.append(
                            f"{sheet}!{cell}={err}" if sheet else f"{cell}={err}"
                        )
                for item in results:
                    if (
                        isinstance(item, dict)
                        and item.get("status") == "ok"
                        and str(item.get("kind") or "").lower() in WRITE_KINDS
                    ):
                        item["post_apply_formula_error_total"] = total_errors
                        if by_type:
                            item["post_apply_formula_error_by_type"] = by_type
                        if sample_refs:
                            item["post_apply_formula_error_cells"] = sample_refs
                results.append(
                    {
                        "index": len(actions),
                        "kind": "formula_error_scan",
                        "status": "error",
                        "errors": formula_errors,
                    }
                )
                _log(
                    "ApplyExcelActionPlan: formula errors detected after recalc "
                    f"total={formula_errors.get('total')}"
                )

        # Final save if any write actions were performed and not already saved after each
        if not save_after_each and any(r.get("kind") in WRITE_KINDS for r in results if r.get("status") == "ok"):
            try:
                if _ENABLE_SCRIPT_SIDE_STORE:
                    final_store_start_ts_ms = int(time.time() * 1000)
                    _store_perf_t0 = time.perf_counter()
                    document.store()
                    phase_timings["final_store"] = int(
                        (time.perf_counter() - _store_perf_t0) * 1000
                    )
                    final_store_end_ts_ms = int(time.time() * 1000)
                    _log(
                        f"ApplyExcelActionPlan[TIME]: phase=final_store "
                        f"ms={phase_timings['final_store']}"
                    )
                    _log("ApplyExcelActionPlan: final document.store() completed")
                else:
                    _log(
                        "ApplyExcelActionPlan: skipped final script-side document.store() (host-managed save)"
                    )
            except Exception as store_exc:
                final_store_end_ts_ms = int(time.time() * 1000)
                _log(f"ApplyExcelActionPlan: final document.store failed: {store_exc}")
    finally:
        if controllers_locked:
            try:
                document.unlockControllers()
                _log("ApplyExcelActionPlan: controllers unlocked")
            except Exception as unlock_exc:
                _log(
                    f"ApplyExcelActionPlan: unlockControllers failed: {unlock_exc}\n"
                    f"{traceback.format_exc()}"
                )
        # Force UI refresh so changes are visible without tab switching
        try:
            _force_ui_refresh(controller)
        except Exception as refresh_exc:
            _log(f"ApplyExcelActionPlan: post-unlock refresh failed: {refresh_exc}")

    # Return a compact JSON summary for the host frame to consume
    try:
        script_return_ts_ms = int(time.time() * 1000)
        total_plan_ms = int((time.perf_counter() - _plan_perf_t0) * 1000)
        actions_total_ms = sum(int(t.get("ms") or 0) for t in action_timings)
        phase_total_ms = sum(int(v or 0) for v in phase_timings.values())
        # Top-5 slowest actions by total ms (handler + snapshot + bookkeeping).
        top_actions = sorted(
            action_timings, key=lambda t: int(t.get("ms") or 0), reverse=True
        )[:5]
        top_actions_brief = [
            f"{t.get('kind')}@{t.get('index')}={t.get('ms')}ms"
            for t in top_actions
        ]
        timing = {
            "script_received_ts_ms": script_received_ts_ms,
            "final_store_start_ts_ms": final_store_start_ts_ms,
            "final_store_end_ts_ms": final_store_end_ts_ms,
            "script_return_ts_ms": script_return_ts_ms,
            # Extended fields — perf_counter()-based, monotonic, ms.
            "total_plan_ms": total_plan_ms,
            "actions_total_ms": actions_total_ms,
            "phase_total_ms": phase_total_ms,
            "phase_timings_ms": dict(phase_timings),
            "action_timings_ms": list(action_timings),
            "top_actions": top_actions_brief,
        }
        _log(
            "ApplyExcelActionPlan[TIMELINE]: plan_id=%s script_received_ts_ms=%s "
            "final_store_start_ts_ms=%s final_store_end_ts_ms=%s script_return_ts_ms=%s"
            % (
                plan.get("plan_id"),
                timing.get("script_received_ts_ms"),
                timing.get("final_store_start_ts_ms"),
                timing.get("final_store_end_ts_ms"),
                timing.get("script_return_ts_ms"),
            )
        )
        _log(
            "ApplyExcelActionPlan[TIME-SUMMARY]: plan_id=%s actions=%s "
            "total_ms=%s actions_total_ms=%s phase_total_ms=%s "
            "phases=%s top_actions=%s"
            % (
                plan.get("plan_id"),
                len(action_timings),
                total_plan_ms,
                actions_total_ms,
                phase_total_ms,
                phase_timings,
                top_actions_brief,
            )
        )
        return json.dumps(
            {
                "ok": True,
                "plan_id": plan.get("plan_id"),
                "file_name": plan.get("file_name"),
                "save_after_each_action": save_after_each,
                "timing": timing,
                "results": results,
            },
            default=str,
        )
    except Exception as exc:
        return json.dumps(
            {
                "ok": False,
                "plan_id": plan.get("plan_id"),
                "error": "response_encode_failed",
                "message": str(exc),
            }
        )


# ---------------------------------------------------------------------------
# Cell Border Helpers
# ---------------------------------------------------------------------------


def _apply_cell_borders(
    sheet,
    start_col: int,
    start_row: int,
    end_col: int,
    end_row: int,
    color: int = 0x000000,
    width: int = 26,
    outline: bool = True,
    inside: bool = True,
):
    """
    Apply visible borders to a cell range.

    Args:
        sheet: The sheet object
        start_col, start_row: Top-left cell position (0-indexed)
        end_col, end_row: Bottom-right cell position (0-indexed)
        color: Border color as integer (default: black 0x000000)
        width: Border width in 1/100mm (default: 26 = ~0.75pt, thin border)
    """
    try:
        cell_range = sheet.getCellRangeByPosition(
            start_col, start_row, end_col, end_row
        )

        # Create border line struct
        border_line = uno.createUnoStruct("com.sun.star.table.BorderLine2")
        border_line.Color = color
        border_line.LineWidth = width
        border_line.OuterLineWidth = width
        border_line.InnerLineWidth = 0
        border_line.LineDistance = 0
        border_line.LineStyle = 0  # SOLID

        # Create table border struct for outer/inner borders
        table_border = uno.createUnoStruct("com.sun.star.table.TableBorder2")
        if outline:
            table_border.TopLine = border_line
            table_border.BottomLine = border_line
            table_border.LeftLine = border_line
            table_border.RightLine = border_line
        if inside:
            table_border.HorizontalLine = border_line
            table_border.VerticalLine = border_line
        table_border.IsTopLineValid = bool(outline)
        table_border.IsBottomLineValid = bool(outline)
        table_border.IsLeftLineValid = bool(outline)
        table_border.IsRightLineValid = bool(outline)
        table_border.IsHorizontalLineValid = bool(inside)
        table_border.IsVerticalLineValid = bool(inside)

        cell_range.TableBorder2 = table_border
        _log(
            f"_apply_cell_borders: Applied borders to range ({start_col},{start_row})-({end_col},{end_row})"
        )
    except Exception as e:
        _log(f"_apply_cell_borders: Error applying borders: {e}")


def _apply_header_style(
    sheet,
    start_col: int,
    row: int,
    end_col: int,
    bg_color: int = _DEFAULT_TABLE_HEADER_BG_COLOR,
    font_color: int = _DEFAULT_TABLE_HEADER_FONT_COLOR,
):
    """
    Apply header styling (background color, bold, font color) to a row.

    Args:
        sheet: The sheet object
        start_col: Starting column (0-indexed)
        row: Row index (0-indexed)
        end_col: Ending column (0-indexed)
        bg_color: Background color (default: dark navy)
        font_color: Font color (default: white)
    """
    try:
        for col in range(start_col, end_col + 1):
            try:
                cell = sheet.getCellByPosition(col, row)
                _set_uno_property(cell, "IsCellBackgroundTransparent", False)
                _set_uno_property(cell, "CellBackColor", bg_color)
                _set_uno_property(cell, "CharWeight", FONT_WEIGHT_BOLD)
                _set_uno_property(cell, "HoriJustify", HORI_CENTER)
                # Headers must never wrap: a wrapped header lets autofit
                # collapse the column to the data width and forces multi-line
                # rendering ("Emplo / yee ID").  Forcing IsTextWrapped=False
                # keeps OptimalWidth honest about the full header text.
                _set_uno_property(cell, "IsTextWrapped", False)
                # Font colour MUST be set last.  Changing CharWeight or
                # other font properties can cause LO to re-derive the
                # ComplexColor from the cell style (theme=1).  Setting the
                # explicit RGB + ComplexColor struct as the final step
                # ensures the export path sees our value, not the theme.
                _set_explicit_char_color(cell, font_color)
            except Exception:
                # Best-effort formatting: fail silently per-cell.
                continue
    except Exception:
        # Best-effort formatting: fail silently for header styling.
        return


def _reset_row_optimal_height(sheet, row_idx: int) -> None:
    """Let Calc recalculate row height after changing wrap/formatting."""
    try:
        row = sheet.getRows().getByIndex(row_idx)
    except Exception:
        return
    try:
        row.OptimalHeight = True
    except Exception:
        pass


def _is_header_like_format(format_spec: Dict[str, Any]) -> bool:
    """Detect single-row header/total/card styling that must not wrap."""
    if not format_spec:
        return False
    font_spec = format_spec.get("font") or {}
    fill_spec = format_spec.get("fill") or {}
    return bool(fill_spec.get("color")) or font_spec.get("bold") is True


def _cell_has_custom_header_style(cell) -> bool:
    if cell is None:
        return False
    try:
        back_color = int(getattr(cell, "CellBackColor", -1))
        if back_color not in (-1, 0xFFFFFF):
            return True
    except Exception:
        pass
    try:
        char_weight = float(getattr(cell, "CharWeight", 0.0))
        if char_weight >= float(FONT_WEIGHT_BOLD):
            return True
    except Exception:
        pass
    try:
        char_color = int(getattr(cell, "CharColor", -1))
        if char_color not in (-1, 0x000000):
            return True
    except Exception:
        pass
    try:
        justify = int(getattr(cell, "HoriJustify"))
        if justify in (HORI_CENTER, HORI_RIGHT):
            return True
    except Exception:
        pass
    return False


def _ensure_default_header_style_if_missing(
    sheet, start_col: int, row: int, end_col: int
) -> bool:
    """
    Apply default header style only when the header row appears unstyled.
    This preserves explicit or inherited workbook styling when present.
    """
    header_cells = []
    for col in range(start_col, end_col + 1):
        try:
            cell = sheet.getCellByPosition(col, row)
        except Exception:
            continue
        if _cell_has_visible_value(cell):
            header_cells.append(cell)
    if not header_cells:
        return False
    if any(_cell_has_custom_header_style(cell) for cell in header_cells):
        return False
    _apply_header_style(
        sheet,
        start_col,
        row,
        end_col,
        bg_color=_DEFAULT_TABLE_HEADER_BG_COLOR,
        font_color=_DEFAULT_TABLE_HEADER_FONT_COLOR,
    )
    return True


def _border_width_from_style(style: Optional[str]) -> int:
    if not style:
        return 26
    style_key = str(style).strip().lower()
    if style_key == "thin":
        return 26
    if style_key == "medium":
        return 50
    if style_key == "thick":
        return 100
    if style_key in ("dashed", "dotted"):
        return 26
    return 26


def _points_to_mm100(points: float) -> int:
    # 1 point = 0.352777 mm -> 1/100 mm units
    return int(round(points * 35.2777))


def _chars_to_mm100(chars: float) -> int:
    # Rough conversion: 1 character ~= 2.6mm at 11pt default font.
    return int(round(chars * 260))


def _autofit_column_with_padding(column, padding_mm100: int = _DEFAULT_AUTOFIT_PADDING_MM100) -> None:
    """
    Apply Calc "optimal width" then add a small extra width for visual padding.
    Caps the final width at _MAX_AUTOFIT_WIDTH_MM100 so a single long cell
    (e.g. a 400-char description) can't stretch the column to half a page.
    Skips hidden columns to avoid unhiding them.
    """
    try:
        if not column.IsVisible:
            return
    except Exception:
        pass
    try:
        column.OptimalWidth = True
    except Exception:
        return
    try:
        padding = int(padding_mm100 or 0)
    except Exception:
        padding = 0
    try:
        current = int(getattr(column, "Width", 0) or 0)
    except Exception:
        current = 0
    if current <= 0:
        return
    target = current + padding if padding > 0 else current
    if target > _MAX_AUTOFIT_WIDTH_MM100:
        target = _MAX_AUTOFIT_WIDTH_MM100
    if target != current:
        try:
            column.Width = target
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Action handlers
# ---------------------------------------------------------------------------


def _handle_set_active_sheet(document, controller, action: Dict[str, Any]):
    ref = action.get("target")
    sheet_name = None
    if isinstance(ref, dict):
        sheet_name = ref.get("sheet_name")
    payload = action.get("payload") or {}
    sheet_name = payload.get("sheet_name") or sheet_name
    sheet = _get_sheet(document, sheet_name)
    if sheet is None:
        raise RuntimeError(f"Sheet '{sheet_name}' not found for set_active_sheet")
    controller.setActiveSheet(sheet)
    return sheet.getCellRangeByName("A1")


def _handle_create_sheet(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    sheet_name = payload.get("sheet_name") or payload.get("name")
    if not sheet_name:
        raise RuntimeError("create_sheet payload missing sheet_name")
    sheets = _get_sheet_container(document)
    if sheets is None:
        raise RuntimeError("create_sheet unable to access document sheets collection")
    if sheets.hasByName(sheet_name):
        _log(
            f"_handle_create_sheet: sheet '{sheet_name}' already exists; skipping insert"
        )
    else:
        position = sheets.getCount()
        sheets.insertNewByName(sheet_name, position)
        _log(f"_handle_create_sheet: inserted sheet '{sheet_name}' at index {position}")
    return _get_sheet(document, sheet_name).getCellRangeByName("A1")


def _handle_rename_sheet(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target.get("sheet_name")
    new_sheet_name = payload.get("new_sheet_name") or payload.get("name")
    if not sheet_name:
        raise RuntimeError("rename_sheet payload missing sheet_name")
    if not new_sheet_name:
        raise RuntimeError("rename_sheet payload missing new_sheet_name")

    sheets = _get_sheet_container(document)
    if sheets is None:
        raise RuntimeError("rename_sheet unable to access document sheets collection")
    if not sheets.hasByName(sheet_name):
        raise RuntimeError(f"rename_sheet: sheet '{sheet_name}' not found")
    if sheet_name != new_sheet_name and sheets.hasByName(new_sheet_name):
        raise RuntimeError(
            f"rename_sheet: destination sheet '{new_sheet_name}' already exists"
        )

    sheet = sheets.getByName(sheet_name)
    if sheet_name != new_sheet_name:
        sheet.setName(str(new_sheet_name))
        _log(f"_handle_rename_sheet: renamed sheet '{sheet_name}' -> '{new_sheet_name}'")
    else:
        _log(f"_handle_rename_sheet: sheet '{sheet_name}' already has requested name")
    return _get_sheet(document, new_sheet_name).getCellRangeByName("A1")


def _handle_delete_sheet(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target.get("sheet_name")
    if not sheet_name:
        raise RuntimeError("delete_sheet payload missing sheet_name")

    sheets = _get_sheet_container(document)
    if sheets is None:
        raise RuntimeError("delete_sheet unable to access document sheets collection")

    names = []
    try:
        names = list(sheets.getElementNames())
    except Exception:
        names = []

    if not names or not sheets.hasByName(sheet_name):
        raise RuntimeError(f"delete_sheet: sheet '{sheet_name}' not found")
    if len(names) <= 1:
        raise RuntimeError(f"delete_sheet: cannot delete the last sheet '{sheet_name}'")

    fallback_name = None
    for candidate in names:
        if candidate != sheet_name:
            fallback_name = candidate
            break

    try:
        active_sheet = controller.getActiveSheet()
        active_name = active_sheet.getName() if active_sheet is not None else None
    except Exception:
        active_name = None

    if active_name == sheet_name and fallback_name and sheets.hasByName(fallback_name):
        controller.setActiveSheet(sheets.getByName(fallback_name))

    sheets.removeByName(sheet_name)
    _log(f"_handle_delete_sheet: removed sheet '{sheet_name}'")

    target_sheet = _get_sheet(document, fallback_name) if fallback_name else None
    if target_sheet is None:
        try:
            active_sheet = controller.getActiveSheet()
            if active_sheet is not None:
                return active_sheet.getCellRangeByName("A1")
        except Exception:
            pass
        raise RuntimeError("delete_sheet: no fallback sheet after delete")
    return target_sheet.getCellRangeByName("A1")


_TYPED_COLUMN_NUMBER_FORMATS = {
    "integer": "#,##0",
    "decimal": "#,##0.####;(#,##0.####);-",
    "currency": '"$"* #,##0.00;"$"* (#,##0.00);"$"* -',  # default USD accounting style
    "percentage": "0.0%",
    "date": "YYYY-MM-DD",
}


# Optional ISO-code → format-string hints. Mirrored from
# backend/services/excel/schema.py CURRENCY_CODE_HINTS. Codes not in this
# dict are accepted as free text and embedded verbatim — see
# _resolve_currency_format. This keeps the pipeline permissive: a user can
# request any currency (e.g. "TZS", "TSh", "Tanzanian Shilling", "万円")
# without a backend change.
_CURRENCY_CODE_HINTS = {
    "USD": '"$"* #,##0.00;"$"* (#,##0.00);"$"* -',
    "EUR": '"€"* #,##0.00;"€"* (#,##0.00);"€"* -',
    "GBP": '"£"* #,##0.00;"£"* (#,##0.00);"£"* -',
    "JPY": '"¥"* #,##0;"¥"* (#,##0);"¥"* -',
    "CNY": '"¥"* #,##0.00;"¥"* (#,##0.00);"¥"* -',
    "INR": '"₹"* #,##0.00;"₹"* (#,##0.00);"₹"* -',
    "KES": '"KSh"* #,##0.00;"KSh"* (#,##0.00);"KSh"* -',
    "UGX": '"USh"* #,##0;"USh"* (#,##0);"USh"* -',
    "TZS": '"TSh"* #,##0.00;"TSh"* (#,##0.00);"TSh"* -',
    "NGN": '"₦"* #,##0.00;"₦"* (#,##0.00);"₦"* -',
    "ZAR": '"R"* #,##0.00;"R"* (#,##0.00);"R"* -',
    "AUD": '"A$"* #,##0.00;"A$"* (#,##0.00);"A$"* -',
    "CAD": '"C$"* #,##0.00;"C$"* (#,##0.00);"C$"* -',
    "CHF": '"CHF"* #,##0.00;"CHF"* (#,##0.00);"CHF"* -',
    "RWF": '"RF"* #,##0;"RF"* (#,##0);"RF"* -',
    "ETB": '"Br"* #,##0.00;"Br"* (#,##0.00);"Br"* -',
    "GHS": '"GH₵"* #,##0.00;"GH₵"* (#,##0.00);"GH₵"* -',
}


def _resolve_currency_format(currency_code) -> str:
    """Permissive currency format resolver. Mirrors
    backend/services/excel/schema.py:resolve_currency_format. Known ISO codes
    render with a Unicode symbol; any other value is embedded verbatim.
    Falls back to the legacy USD format when no code is given.
    """
    if not currency_code:
        return _TYPED_COLUMN_NUMBER_FORMATS["currency"]
    text = str(currency_code).strip()
    if not text:
        return _TYPED_COLUMN_NUMBER_FORMATS["currency"]
    upper = text.upper()
    hinted = _CURRENCY_CODE_HINTS.get(upper)
    if hinted:
        return hinted
    safe = text.replace('"', "")
    return f'"{safe}"* #,##0.00;"{safe}"* (#,##0.00);"{safe}"* -'


def _column_currency_code(col_def):
    """Extract optional ``currency_code`` (free text) from a ColumnDef payload."""
    if isinstance(col_def, dict):
        code = col_def.get("currency_code")
        return str(code) if code else None
    return None

# LO Calc date epoch.  setValue(serial) on a cell with a date number format
# renders as the formatted date.  Using setString() leaves the cell as TEXT
# and the date format has no effect — that's the bug this constant fixes.
_LO_DATE_EPOCH = None  # initialized lazily inside _coerce_value_for_data_type

_DATE_PARSE_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M:%SZ",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%m/%d/%Y",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%d %b %Y",
    "%d %B %Y",
)


def _coerce_value_for_data_type(value, data_type):
    """Coerce a row value to match the column's declared data_type so that
    LibreOffice stores it as a number/date/bool, not as text.

    Why this exists: agents (and JSON in general) frequently encode numbers
    and dates as strings.  ``cell.setString("1800000")`` makes the cell TEXT;
    a currency number_format on a TEXT cell does nothing.  This helper turns
    such strings into the correct native value before _write_cell runs.

    Returns the coerced value, or the original value if coercion is not
    applicable, not needed, or fails.  No exception is raised — fall back to
    the original value and let _write_cell render it as text.
    """
    if value is None or value == "" or not data_type:
        return value
    if data_type in ("integer", "decimal", "currency"):
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            cleaned = value.strip().replace(",", "").replace(" ", "")
            if cleaned.startswith("$"):
                cleaned = cleaned[1:]
            if cleaned.startswith("(") and cleaned.endswith(")"):
                cleaned = "-" + cleaned[1:-1]
            try:
                return float(cleaned)
            except (ValueError, TypeError):
                return value
        return value
    if data_type == "percentage":
        if isinstance(value, bool):
            return float(int(value))
        if isinstance(value, (int, float)):
            # Heuristic: a percentage column receiving 50 (not 0.5) — assume
            # the agent meant 50% rather than 5000%.  Anything already in
            # 0..1 we treat as a fraction.
            return float(value) / 100.0 if abs(value) > 1 else float(value)
        if isinstance(value, str):
            cleaned = value.strip()
            had_pct = cleaned.endswith("%")
            cleaned = cleaned.rstrip("%").replace(",", "").strip()
            try:
                num = float(cleaned)
                return num / 100.0 if had_pct else num
            except (ValueError, TypeError):
                return value
        return value
    if data_type == "date":
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return value  # already a serial date
        if isinstance(value, str):
            from datetime import datetime
            text = value.strip()
            for fmt in _DATE_PARSE_FORMATS:
                try:
                    dt = datetime.strptime(text, fmt)
                    epoch = datetime(1899, 12, 30)
                    delta = dt - epoch
                    return delta.days + (delta.seconds / 86400.0)
                except ValueError:
                    continue
        return value
    if data_type == "boolean":
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(value)
        if isinstance(value, str):
            v = value.strip().lower()
            if v in ("true", "yes", "y", "1"):
                return True
            if v in ("false", "no", "n", "0"):
                return False
        return value
    return value


def _infer_data_type_from_format(number_formats, fmt_id) -> Optional[str]:
    """Reverse-map a cell's NumberFormat back to a data_type so APPEND_ROWS
    knows how to coerce string values.  Best-effort heuristic on the format
    code's substring; ambiguous formats return None and the caller writes
    the value verbatim."""
    if number_formats is None or fmt_id is None:
        return None
    try:
        fmt_string = _number_format_string(number_formats, fmt_id)
    except Exception:
        fmt_string = None
    if not fmt_string:
        return None
    s = fmt_string.lower()
    # Date precedes everything because "yyyy", "mm", "dd" tokens are unique
    # to date/time formats — they don't appear in numeric format codes.
    if any(tok in s for tok in ("yyyy", "yy", "mmm", "dd", "hh:mm")):
        return "date"
    if "%" in s:
        return "percentage"
    if "$" in s or "€" in s or "£" in s or "¥" in s:
        return "currency"
    if "0.00" in s or "#.##0.00" in s or "0.0" in s:
        return "decimal"
    if "0" in s or "#" in s:
        return "integer"
    return None


def _column_header_name(col_def):
    """Return a display string for a column definition, accepting either a
    plain string or a {"name": ..., "data_type": ...} dict.  Defensive: a
    bug elsewhere that leaks raw ColumnDef dicts will not produce a dict-repr
    cell value."""
    if isinstance(col_def, dict):
        return str(col_def.get("name") or "")
    return str(col_def or "")


def _column_data_type(col_def):
    if isinstance(col_def, dict):
        dt = col_def.get("data_type")
        return str(dt) if dt else None
    return None


def _row_strings_at(sheet, row_idx: int, start_col: int, end_col: int):
    """Return the display strings of cells in a single row, A→… ."""
    out = []
    for col in range(start_col, end_col + 1):
        try:
            cell = sheet.getCellByPosition(col, row_idx)
            out.append((cell.getString() or "").strip())
        except Exception:
            out.append("")
    return out


def _sheet_name_for_state(sheet, fallback: Optional[str] = None) -> str:
    try:
        get_name = getattr(sheet, "getName", None)
        if callable(get_name):
            name = get_name()
            if isinstance(name, str) and name:
                return name
    except Exception:
        pass
    try:
        name = getattr(sheet, "Name", None)
        if isinstance(name, str) and name:
            return name
    except Exception:
        pass
    return str(fallback or "")


def _same_sheet_for_state(left: Optional[str], right: Optional[str]) -> bool:
    if str(left or "") == str(right or ""):
        return True
    try:
        return _normalize_sheet_name(str(left or "")) == _normalize_sheet_name(
            str(right or "")
        )
    except Exception:
        return False


def _range_bounds_for_state(cell_range) -> Optional[Tuple[int, int, int, int]]:
    try:
        address = cell_range.getRangeAddress()
        return (
            int(address.StartColumn),
            int(address.StartRow),
            int(address.EndColumn),
            int(address.EndRow),
        )
    except Exception:
        return None


def _bounds_to_a1(bounds: Tuple[int, int, int, int]) -> str:
    start_col, start_row, end_col, end_row = bounds
    start = f"{_column_index_to_name(start_col)}{start_row + 1}"
    end = f"{_column_index_to_name(end_col)}{end_row + 1}"
    return start if start == end else f"{start}:{end}"


def _record_duplicate_header_write_skip(
    sheet,
    start_col: int,
    skipped_header_row: int,
    end_col: int,
    existing_header_row: int,
) -> Dict[str, Any]:
    sheet_name = _sheet_name_for_state(sheet)
    skipped_bounds = (start_col, skipped_header_row, end_col, skipped_header_row)
    existing_bounds = (start_col, existing_header_row, end_col, existing_header_row)
    entry = {
        "sheet_name": sheet_name,
        "skipped_header_bounds": skipped_bounds,
        "existing_header_bounds": existing_bounds,
        "data_start_cell": f"{_column_index_to_name(start_col)}{skipped_header_row + 2}",
        "skipped_header_range_a1": _bounds_to_a1(skipped_bounds),
        "existing_header_range_a1": _bounds_to_a1(existing_bounds),
    }
    _DUPLICATE_HEADER_WRITE_SKIPS.append(entry)
    return entry


def _find_duplicate_header_write_skip(sheet, cell_range) -> Optional[Dict[str, Any]]:
    bounds = _range_bounds_for_state(cell_range)
    if bounds is None:
        return None
    sheet_name = _sheet_name_for_state(sheet)
    for entry in reversed(_DUPLICATE_HEADER_WRITE_SKIPS):
        if not _same_sheet_for_state(sheet_name, entry.get("sheet_name")):
            continue
        if tuple(entry.get("skipped_header_bounds") or ()) == tuple(bounds):
            return entry
    return None


def _find_duplicate_header_freeze_skip(sheet, freeze_cell) -> Optional[Dict[str, Any]]:
    try:
        address = freeze_cell.getCellAddress()
        freeze_ref = f"{_column_index_to_name(int(address.Column))}{int(address.Row) + 1}"
    except Exception:
        return None
    sheet_name = _sheet_name_for_state(sheet)
    for entry in reversed(_DUPLICATE_HEADER_WRITE_SKIPS):
        if not _same_sheet_for_state(sheet_name, entry.get("sheet_name")):
            continue
        if str(entry.get("data_start_cell") or "") == freeze_ref:
            return entry
    return None


def _handle_write_range(document, controller, action: Dict[str, Any]):
    target = action.get("target") or {}
    payload = action.get("payload") or {}
    sheet_name = target.get("sheet_name") or payload.get("sheet_name")

    range_name = target.get("range") or payload.get("start_cell") or "A1"
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "write_range"
    )
    if sheet is None:
        raise RuntimeError(f"write_range target sheet '{sheet_name}' not found")
    start_range = sheet.getCellRangeByName(range_name)
    address = start_range.getRangeAddress()
    start_row = address.StartRow
    start_col = address.StartColumn

    table = payload.get("data") or {}
    columns = table.get("columns") or []
    rows = table.get("rows") or []

    # Check for styling options in payload
    apply_borders = payload.get("apply_borders", True)  # Default to True
    style_header = payload.get("style_header", True)  # Default to True
    border_color = _parse_hex_color_to_int(payload.get("border_color"))
    if border_color is None:
        border_color = _DEFAULT_TABLE_BORDER_COLOR
    header_bg = _parse_hex_color_to_int(payload.get("header_bg_color"))
    if header_bg is None:
        header_bg = _DEFAULT_TABLE_HEADER_BG_COLOR
    header_fg = _parse_hex_color_to_int(payload.get("header_font_color"))
    if header_fg is None:
        header_fg = _DEFAULT_TABLE_HEADER_FONT_COLOR

    column_names = [_column_header_name(c) for c in columns]
    num_cols = max(len(columns), len(rows[0]) if rows else 1)

    # Header-collision detection: agents using WRITE_DATA to add a totals
    # row, append more rows below a table, or write below an existing table
    # re-emit the header because the schema makes columns mandatory.  Two
    # distinct collision shapes:
    #
    #   (a) start_row IS the row of an existing header → agent is REWRITING
    #       the table from the top.  Re-write the header (no-op visually),
    #       data rows go at start_row+1 as usual.  Treated as no-collision.
    #
    #   (b) start_row is BELOW a row that already holds the same header →
    #       agent is adding rows or a totals line.  Skip the header write,
    #       AND advance the data-row pointer by one so the agent's intended
    #       data row goes at start_row+1 (not start_row, which would
    #       overwrite an existing data row).  This matches the agent's
    #       mental model: "header at start_cell, data at start_cell+1".
    #
    # Bounded scan: 200 rows is enough for any plausible table without
    # making the lookup O(sheet).
    skip_header = False
    duplicate_header_existing_row = None
    if columns and num_cols > 0:
        end_col_check = start_col + num_cols - 1
        normalized_new = [n.strip() for n in column_names]
        if any(normalized_new):
            scan_top = max(0, start_row - 200)
            # Scan rows STRICTLY ABOVE start_row.  If start_row itself holds
            # the matching header, that's the rewrite case (a) and we leave
            # skip_header False.
            for probe_row in range(start_row - 1, scan_top - 1, -1):
                existing = _row_strings_at(
                    sheet, probe_row, start_col, end_col_check
                )
                if existing == normalized_new:
                    _log(
                        f"_handle_write_range: header already exists at row "
                        f"{probe_row + 1}; skipping duplicate header at row "
                        f"{start_row + 1} and writing data from row "
                        f"{start_row + 2}"
                    )
                    skip_header = True
                    duplicate_header_existing_row = probe_row
                    break

    current_row = start_row
    header_row = None
    first_data_row = start_row

    if columns and not skip_header:
        header_row = current_row
        for offset, header in enumerate(columns):
            cell = sheet.getCellByPosition(start_col + offset, current_row)
            _write_cell(cell, _column_header_name(header))
        current_row += 1
        first_data_row = current_row
    elif skip_header:
        # Skip the (would-have-been-duplicate) header AND advance the data
        # pointer so the agent's data lands where they expected — at
        # start_row+1 — not on top of the existing row at start_row.
        current_row += 1
        first_data_row = current_row
        if duplicate_header_existing_row is not None:
            try:
                skip_entry = _record_duplicate_header_write_skip(
                    sheet,
                    start_col,
                    start_row,
                    start_col + num_cols - 1,
                    duplicate_header_existing_row,
                )
                meta = action.setdefault("_result_meta", {})
                meta["skipped_duplicate_header"] = True
                meta["skipped_header_range_a1"] = skip_entry.get(
                    "skipped_header_range_a1"
                )
                meta["existing_header_range_a1"] = skip_entry.get(
                    "existing_header_range_a1"
                )
                meta["data_start_cell_after_skipped_header"] = skip_entry.get(
                    "data_start_cell"
                )
            except Exception as exc:
                _log(
                    "_handle_write_range: failed to record duplicate header "
                    f"skip state: {exc}"
                )

    column_data_types = [_column_data_type(c) for c in columns] if columns else []
    for row_data in rows:
        for offset, raw in enumerate(row_data):
            cell = sheet.getCellByPosition(start_col + offset, current_row)
            # CellObject form: {"value": ..., "data_type": ..., ...}.
            # Cell-level data_type overrides the column-level fallback.
            if isinstance(raw, dict) and "value" in raw:
                value = raw.get("value")
                cell_dt = raw.get("data_type")
                col_dt = column_data_types[offset] if offset < len(column_data_types) else None
                dt = cell_dt or col_dt
            else:
                value = raw
                dt = column_data_types[offset] if offset < len(column_data_types) else None
            _write_cell(cell, _coerce_value_for_data_type(value, dt))
        current_row += 1

    # Calculate the data range dimensions
    end_col = start_col + num_cols - 1
    end_row = max(current_row - 1, start_row)
    applied_range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)

    # Apply cell borders to make the data look like a proper Excel table
    if apply_borders and (columns or rows):
        _apply_cell_borders(
            sheet, start_col, start_row, end_col, end_row, color=border_color
        )

    # Apply header styling if there are column headers.
    # If explicit styling is not applied, enforce default header style only when
    # the header currently appears unstyled.
    header_style_applied = False
    if style_header and header_row is not None and columns:
        _apply_header_style(
            sheet,
            start_col,
            header_row,
            end_col,
            bg_color=header_bg,
            font_color=header_fg,
        )
        header_style_applied = True
    if header_row is not None and columns and not header_style_applied:
        if _ensure_default_header_style_if_missing(
            sheet, start_col, header_row, end_col
        ):
            _log(
                "_handle_write_range: applied default header style "
                f"row={header_row} cols={start_col}-{end_col}"
            )

    # Apply per-column number formats from data_type metadata.  This makes
    # the LO macro self-sufficient: the agent's data_type declarations always
    # produce the correct number format on the data cells, regardless of
    # whether preflight injected a deferred APPLY_FORMATTING.
    last_data_row = end_row
    if columns and rows and first_data_row <= last_data_row:
        try:
            number_formats = document.getNumberFormats()
            locale = _resolve_locale(number_formats)
        except Exception:
            number_formats = None
            locale = None
        if number_formats is not None:
            for c_idx, col_def in enumerate(columns):
                dt = _column_data_type(col_def)
                if not dt:
                    continue
                if dt == "currency":
                    # Free-text currency_code wins over the hardcoded USD default.
                    fmt_str = _resolve_currency_format(_column_currency_code(col_def))
                else:
                    fmt_str = _TYPED_COLUMN_NUMBER_FORMATS.get(dt)
                if not fmt_str:
                    continue
                try:
                    fmt_id = number_formats.queryKey(fmt_str, locale, False)
                    if fmt_id == -1:
                        fmt_id = number_formats.addNew(fmt_str, locale)
                except Exception:
                    continue
                col_position = start_col + c_idx
                for row_idx in range(first_data_row, last_data_row + 1):
                    try:
                        cell = sheet.getCellByPosition(col_position, row_idx)
                        cell.NumberFormat = int(fmt_id)
                    except Exception:
                        continue

    meta = action.setdefault("_result_meta", {})
    applied_abs = _range_to_representation(applied_range)
    if applied_abs:
        meta["applied_table_range"] = applied_abs
    applied_local = _range_to_a1_relative(applied_range)
    if applied_local:
        meta["applied_table_range_a1"] = applied_local

    last_cell = sheet.getCellByPosition(end_col, end_row)
    return _expand_to_single_cell(last_cell)


def _handle_set_cell_formula(document, controller, action: Dict[str, Any]):
    target = action.get("target") or {}
    payload = action.get("payload") or {}
    sheet_name = target.get("sheet_name") or payload.get("sheet_name")
    cell_name = target.get("range") or payload.get("target_cell")
    if not cell_name:
        raise RuntimeError("set_cell_formula missing target cell range")
    sheet, cell_name = _resolve_sheet_and_range_name(
        document, sheet_name, cell_name, "set_cell_formula"
    )
    if sheet is None:
        raise RuntimeError(f"set_cell_formula sheet '{sheet_name}' not found")
    cell = sheet.getCellRangeByName(cell_name)
    formula = payload.get("formula_string") or payload.get("formula")
    if formula is None:
        raise RuntimeError("set_cell_formula missing formula string")
    formula_text = str(formula).strip()
    if not formula_text:
        raise RuntimeError("set_cell_formula missing formula string")
    if not formula_text.startswith("="):
        formula_text = "=" + formula_text
    _validate_formula_references(document, sheet.getName(), formula_text)
    _set_formula_with_locale(cell, formula_text)
    return _expand_to_single_cell(cell)


def _handle_append_rows(document, controller, action: Dict[str, Any]):
    target = action.get("target") or {}
    payload = action.get("payload") or {}
    sheet_name = target.get("sheet_name") or payload.get("sheet_name")
    range_name = target.get("range")
    if not range_name:
        raise RuntimeError("append_rows requires a target range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "append_rows"
    )
    if sheet is None:
        raise RuntimeError(f"append_rows sheet '{sheet_name}' not found")
    base_range = sheet.getCellRangeByName(range_name)
    address = base_range.getRangeAddress()
    start_col = address.StartColumn
    insert_rows = payload.get("data") or []
    row_count = len(insert_rows)
    if not insert_rows:
        return _expand_to_single_cell(base_range)
    _log("append_rows: sheet=%s range=%s rows=%d" % (sheet_name, range_name, row_count))

    max_columns = max(
        (len(row_values) for row_values in insert_rows if row_values),
        default=0,
    )
    if max_columns == 0:
        max_columns = 1

    initial_end_col = address.EndColumn
    effective_end_col = max(initial_end_col, start_col + max_columns - 1)

    last_data_row = address.StartRow - 1
    try:
        data_rows = list(map(list, base_range.getDataArray()))
    except Exception:
        data_rows = []
    if data_rows:
        first_row_index = address.StartRow
        for offset, row in enumerate(data_rows):
            if _row_from_data_has_content(row):
                last_data_row = first_row_index + offset
        if last_data_row < address.StartRow:
            last_data_row = address.StartRow - 1
    else:
        last_data_row = address.StartRow - 1

    next_row = last_data_row + 1

    needs_shift = False
    for offset in range(row_count):
        if _row_contains_content(
            sheet, next_row + offset, start_col, effective_end_col
        ):
            needs_shift = True
            break

    if needs_shift:
        try:
            rows = sheet.getRows()
            total_rows = rows.getCount() if hasattr(rows, "getCount") else 0
            insert_index = min(max(next_row, 0), max(total_rows - 1, 0))
            rows.insertByIndex(insert_index, row_count)
        except Exception as exc:
            _log(
                f"append_rows: insertByIndex failed ({exc}); trying sheet.insertCells fallback"
            )
            if not _insert_rows_via_sheet(
                sheet, next_row, row_count, start_col, effective_end_col
            ):
                if not _shift_rows_down_manual(
                    sheet, start_col, effective_end_col, next_row, row_count
                ):
                    raise RuntimeError(
                        "append_rows failed to insert blank rows before writing"
                    ) from exc

    # Ensure the target write rows exist even when there is no content to shift below.
    try:
        rows = sheet.getRows()
        total_rows = rows.getCount() if hasattr(rows, "getCount") else 0
    except Exception:
        total_rows = 0
    required_rows = next_row + row_count
    if required_rows > max(total_rows, 0):
        missing = required_rows - max(total_rows, 0)
        _log(
            f"append_rows: growing sheet rows; total={total_rows} need={required_rows} add={missing}"
        )
        grew = False
        try:
            grow_index = max(total_rows, 0)
            sheet.getRows().insertByIndex(grow_index, missing)
            grew = True
        except Exception as exc:
            _log(
                f"append_rows: grow insertByIndex failed ({exc}); trying sheet.insertCells"
            )
        if not grew:
            if not _insert_rows_via_sheet(
                sheet, max(total_rows - 1, 0), missing, start_col, effective_end_col
            ):
                # As a last resort, try inserting at column 0 across sheet
                if not _insert_rows_via_sheet(
                    sheet, max(total_rows - 1, 0), missing, 0, effective_end_col
                ):
                    _log(
                        "append_rows: unable to grow rows; proceeding but writes may fail"
                    )

    # Use the last existing data row (before insertion) for template styling.
    template_row = max(address.StartRow, last_data_row)
    current_row = next_row

    # Pre-compute each template column's inferred data_type from its
    # NumberFormat string.  Without this the appended value would be written
    # as TEXT (because it's a JSON string) and the inherited number format
    # would have no effect ("1988-03-12" stays a text cell).
    try:
        number_formats = document.getNumberFormats()
    except Exception:
        number_formats = None
    template_data_types: List[Optional[str]] = []
    for idx in range(max_columns):
        template_col = min(start_col + idx, initial_end_col)
        try:
            tcell = sheet.getCellByPosition(template_col, template_row)
            fmt_id = getattr(tcell, "NumberFormat", None)
        except Exception:
            fmt_id = None
        template_data_types.append(
            _infer_data_type_from_format(number_formats, fmt_id)
        )

    for row_values in insert_rows:
        row_length = len(row_values)
        _log("append_rows: writing row=%s len=%s" % (row_values, row_length))
        for idx in range(max_columns):
            col_position = start_col + idx
            cell = sheet.getCellByPosition(col_position, current_row)
            template_col = min(col_position, initial_end_col)
            template_cell = sheet.getCellByPosition(template_col, template_row)
            _copy_cell_style(template_cell, cell)
            if idx < row_length:
                value = row_values[idx]
                dt = template_data_types[idx] if idx < len(template_data_types) else None
                _write_cell_with_validation(cell, _coerce_value_for_data_type(value, dt))
            else:
                cell.setString("")
        current_row += 1

    applied_range = sheet.getCellRangeByPosition(
        start_col,
        next_row,
        effective_end_col,
        current_row - 1,
    )
    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "rows_appended": int(row_count),
            "appended_start_row": int(next_row) + 1,
            "appended_end_row": int(current_row),
            "appended_start_column": int(start_col) + 1,
            "appended_end_column": int(effective_end_col) + 1,
            "shifted_existing_rows_down": bool(needs_shift),
        }
    )
    return applied_range


def _handle_create_column(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("create_column requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "create_column"
    )
    if sheet is None:
        raise RuntimeError(f"create_column sheet '{sheet_name}' not found")
    data_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, data_range.getDataArray()))
    if not data:
        raise RuntimeError("create_column target range is empty")

    header = data[0]
    header_map = {str(name): idx for idx, name in enumerate(header)}

    column_name = payload.get("column_name")
    if not column_name:
        raise RuntimeError("create_column missing column_name")

    calc_spec = payload.get("calculation") or {}
    calc_type = (calc_spec.get("type") or "").upper()
    operation = (calc_spec.get("operation") or "").upper()
    source_columns = calc_spec.get("source_columns") or []
    blank_column = False
    if not source_columns:
        blank_column = True
    else:
        if calc_type != "ROW_WISE_OPERATION":
            raise RuntimeError(
                "create_column currently supports ROW_WISE_OPERATION only"
            )
        if not operation:
            raise RuntimeError(
                "create_column missing operation for row-wise calculation"
            )

    source_indices = []
    for column in source_columns:
        if column not in header_map:
            raise RuntimeError(
                f"create_column source column '{column}' not found in header"
            )
        source_indices.append(header_map[column])

    address = data_range.getRangeAddress()
    position, new_col_index, template_col_index = _resolve_column_insert_location(
        header, address, payload, "create_column"
    )
    conditional_format_snapshot = None
    if position != "END":
        conditional_format_snapshot = _capture_column_property_matrix(
            sheet,
            "ConditionalFormat",
            new_col_index,
            address.EndColumn,
            address.StartRow,
            address.EndRow,
        )
    if position != "END":
        inserted = _insert_columns_via_sheet(
            sheet, new_col_index, 1, address.StartRow, address.EndRow
        )
        if not inserted:
            inserted = _shift_columns_right_manual(
                sheet,
                new_col_index,
                address.EndColumn,
                address.StartRow,
                address.EndRow,
                1,
            )
        if not inserted:
            try:
                sheet.getColumns().insertByIndex(new_col_index, 1)
                inserted = True
            except Exception as exc:
                _log(f"create_column: insertByIndex fallback failed: {exc}")
        if not inserted:
            raise RuntimeError("create_column failed to insert a blank column slot")
        _shift_column_property_right_from_snapshot(
            sheet,
            "ConditionalFormat",
            conditional_format_snapshot,
            new_col_index,
            address.EndColumn,
            address.StartRow,
            address.EndRow,
        )
    header_row = address.StartRow
    if position == "END":
        template_header = sheet.getCellByPosition(template_col_index, header_row)
        if not _cell_has_visible_value(template_header):
            template_col_index = max(address.StartColumn, template_col_index - 1)
            template_header = sheet.getCellByPosition(template_col_index, header_row)

    try:
        _clone_column_format(
            sheet, template_col_index, new_col_index, address.StartRow, address.EndRow
        )
    except Exception:
        _copy_column_styles(
            sheet, template_col_index, new_col_index, address.StartRow, address.EndRow
        )

    header_cell = sheet.getCellByPosition(new_col_index, header_row)
    header_cell.setString(column_name)
    table_end_col = max(address.EndColumn + 1, new_col_index)
    if _ensure_default_header_style_if_missing(
        sheet, address.StartColumn, header_row, table_end_col
    ):
        _log(
            "_handle_create_column: applied default header style "
            f"row={header_row} cols={address.StartColumn}-{table_end_col}"
        )

    try:
        columns = sheet.getColumns()
        source_column_obj = columns.getByIndex(template_col_index)
        target_column_obj = columns.getByIndex(new_col_index)
        try:
            target_column_obj.Width = source_column_obj.Width
        except Exception:
            pass
        try:
            _autofit_column_with_padding(target_column_obj)
        except Exception:
            pass
    except Exception:
        pass

    for row_offset, row_data in enumerate(data[1:], start=1):
        values = []
        for idx in source_indices:
            if idx < len(row_data):
                values.append(_coerce_number(row_data[idx]))
            else:
                values.append(None)
        if blank_column:
            result = None
        else:
            result = _aggregate_row(values, operation)
        cell = sheet.getCellByPosition(new_col_index, address.StartRow + row_offset)
        template_cell = sheet.getCellByPosition(
            template_col_index, address.StartRow + row_offset
        )
        _copy_cell_style(template_cell, cell)
        if result is None:
            cell.setString("")
        else:
            cell.setValue(result)

    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "created_column_name": column_name,
            "created_column_index_1based": int(new_col_index) + 1,
            "created_data_row_count": max(0, len(data) - 1),
        }
    )
    return sheet.getCellRangeByPosition(
        new_col_index,
        header_row,
        new_col_index,
        address.EndRow,
    )


def _handle_create_conditional_column(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("create_conditional_column requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "create_conditional_column"
    )
    if sheet is None:
        raise RuntimeError(f"create_conditional_column sheet '{sheet_name}' not found")

    data_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, data_range.getDataArray()))
    if not data or len(data) < 2:
        raise RuntimeError(
            "create_conditional_column target range must include header and data rows"
        )

    header = data[0]
    header_map = {str(name): idx for idx, name in enumerate(header)}

    new_column_name = payload.get("new_column_name")
    if not new_column_name:
        raise RuntimeError("create_conditional_column missing new_column_name")

    conditions = payload.get("conditions") or []
    if not conditions:
        raise RuntimeError("create_conditional_column requires conditions")

    normalized_conditions: List[Dict[str, Any]] = []
    for cond in conditions:
        column_name = cond.get("column")
        operator = (cond.get("operator") or "").upper()
        value = cond.get("value")
        output = cond.get("output")
        if not column_name or column_name not in header_map:
            raise RuntimeError(
                f"create_conditional_column condition references unknown column '{column_name}'"
            )
        if not operator:
            raise RuntimeError("create_conditional_column condition missing operator")
        normalized_conditions.append(
            {
                "index": header_map[column_name],
                "operator": operator,
                "value": value,
                "output": output,
            }
        )

    default_value = payload.get("default_value")

    address = data_range.getRangeAddress()
    position, new_col_index, template_col_index = _resolve_column_insert_location(
        header, address, payload, "create_conditional_column"
    )
    conditional_format_snapshot = None
    if position != "END":
        conditional_format_snapshot = _capture_column_property_matrix(
            sheet,
            "ConditionalFormat",
            new_col_index,
            address.EndColumn,
            address.StartRow,
            address.EndRow,
        )
    if position != "END":
        inserted = _insert_columns_via_sheet(
            sheet, new_col_index, 1, address.StartRow, address.EndRow
        )
        if not inserted:
            inserted = _shift_columns_right_manual(
                sheet,
                new_col_index,
                address.EndColumn,
                address.StartRow,
                address.EndRow,
                1,
            )
        if not inserted:
            try:
                sheet.getColumns().insertByIndex(new_col_index, 1)
                inserted = True
            except Exception as exc:
                _log(f"create_conditional_column: insertByIndex fallback failed: {exc}")
        if not inserted:
            raise RuntimeError(
                "create_conditional_column failed to insert a blank column slot"
            )
        _shift_column_property_right_from_snapshot(
            sheet,
            "ConditionalFormat",
            conditional_format_snapshot,
            new_col_index,
            address.EndColumn,
            address.StartRow,
            address.EndRow,
        )
    header_row = address.StartRow
    if position == "END":
        template_header = sheet.getCellByPosition(template_col_index, header_row)
        if not _cell_has_visible_value(template_header):
            template_col_index = max(address.StartColumn, template_col_index - 1)
            template_header = sheet.getCellByPosition(template_col_index, header_row)

    try:
        _clone_column_format(
            sheet, template_col_index, new_col_index, address.StartRow, address.EndRow
        )
    except Exception:
        _copy_column_styles(
            sheet, template_col_index, new_col_index, address.StartRow, address.EndRow
        )

    header_cell = sheet.getCellByPosition(new_col_index, header_row)
    header_cell.setString(new_column_name)
    table_end_col = max(address.EndColumn + 1, new_col_index)
    if _ensure_default_header_style_if_missing(
        sheet, address.StartColumn, header_row, table_end_col
    ):
        _log(
            "_handle_create_conditional_column: applied default header style "
            f"row={header_row} cols={address.StartColumn}-{table_end_col}"
        )

    try:
        columns = sheet.getColumns()
        source_column_obj = columns.getByIndex(template_col_index)
        target_column_obj = columns.getByIndex(new_col_index)
        try:
            target_column_obj.Width = source_column_obj.Width
        except Exception:
            pass
        try:
            _autofit_column_with_padding(target_column_obj)
        except Exception:
            pass
    except Exception:
        pass

    for row_offset, row_values in enumerate(data[1:], start=1):
        resolved_value = default_value
        for cond in normalized_conditions:
            column_index = cond["index"]
            cell_value = (
                row_values[column_index] if column_index < len(row_values) else None
            )
            if _compare_values(cell_value, cond["operator"], cond["value"]):
                resolved_value = cond.get("output")
                break
        abs_row = address.StartRow + row_offset
        cell = sheet.getCellByPosition(new_col_index, abs_row)
        template_cell = sheet.getCellByPosition(template_col_index, abs_row)
        _copy_cell_style(template_cell, cell)
        _write_cell_with_validation(cell, resolved_value)

    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "created_column_name": new_column_name,
            "created_column_index_1based": int(new_col_index) + 1,
            "created_data_row_count": max(0, len(data) - 1),
        }
    )
    return sheet.getCellRangeByPosition(
        new_col_index,
        header_row,
        new_col_index,
        address.EndRow,
    )


def _handle_delete_columns(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("delete_columns requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "delete_columns"
    )
    if sheet is None:
        raise RuntimeError(f"delete_columns sheet '{sheet_name}' not found")

    column_names = payload.get("column_names") or []
    if not column_names:
        return sheet.getCellRangeByName(range_name)

    cell_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, cell_range.getDataArray()))
    if not data or len(data) < 1:
        return cell_range

    header = data[0]
    header_map: Dict[str, int] = {}
    for idx, name in enumerate(header):
        normalized = _normalize_header_label(name)
        if normalized and normalized not in header_map:
            header_map[normalized] = idx

    address = cell_range.getRangeAddress()
    absolute_columns: list[int] = []
    for name in column_names:
        normalized = _normalize_header_label(name)
        if not normalized or normalized not in header_map:
            _log(
                f"delete_columns: column '{name}' not found in range header {list(header_map.keys())}"
            )
            continue
        absolute_columns.append(address.StartColumn + header_map[normalized])

    if not absolute_columns:
        _log("delete_columns: no matching columns to delete; skipping")
        return cell_range

    mode_left = _get_cell_delete_mode_left()
    sheet_index = address.Sheet
    for abs_col in sorted(set(absolute_columns), reverse=True):
        range_address = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
        range_address.Sheet = sheet_index
        range_address.StartColumn = abs_col
        range_address.EndColumn = abs_col
        range_address.StartRow = address.StartRow
        range_address.EndRow = address.EndRow
        removed = False
        if mode_left:
            try:
                sheet.removeRange(range_address, mode_left)
                removed = True
            except Exception as exc:
                _log(f"delete_columns: removeRange failed for column {abs_col}: {exc}")
        if not removed:
            try:
                sheet.getColumns().removeByIndex(abs_col, 1)
                removed = True
            except Exception as exc:
                _log(
                    f"delete_columns: removeByIndex failed for column {abs_col}: {exc}"
                )
        if not removed:
            for row in range(range_address.StartRow, range_address.EndRow + 1):
                cell = sheet.getCellByPosition(abs_col, row)
                _write_cell(cell, "")

    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "deleted_columns_count": len(sorted(set(absolute_columns))),
            "deleted_columns_requested": [str(name) for name in column_names],
            "structure_shift": "left",
        }
    )
    return sheet.getCellRangeByName(range_name)


def _handle_delete_rows(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("delete_rows requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "delete_rows"
    )
    if sheet is None:
        raise RuntimeError(f"delete_rows sheet '{sheet_name}' not found")

    cell_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, cell_range.getDataArray()))
    if not data or len(data) <= 1:
        return cell_range

    header = data[0]
    header_map = {str(name): idx for idx, name in enumerate(header)}
    conditions = payload.get("find_conditions") or []
    if not conditions:
        return cell_range

    rows_to_delete: list[int] = []
    for idx, row_values in enumerate(data[1:], start=1):
        if _row_matches_conditions(row_values, header_map, conditions):
            rows_to_delete.append(idx)

    if not rows_to_delete:
        return cell_range

    address = cell_range.getRangeAddress()
    sheet_index = address.Sheet
    for relative_row in reversed(rows_to_delete):
        abs_row = address.StartRow + relative_row
        range_address = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
        range_address.Sheet = sheet_index
        range_address.StartColumn = address.StartColumn
        range_address.EndColumn = address.EndColumn
        range_address.StartRow = abs_row
        range_address.EndRow = abs_row
        removed = False
        try:
            sheet.getRows().removeByIndex(abs_row, 1)
            removed = True
        except Exception as exc:
            _log(f"_handle_delete_rows: removeByIndex failed at row {abs_row}: {exc}")
        if not removed:
            mode_up = _get_cell_delete_mode_up()
            if mode_up:
                try:
                    sheet.removeRange(range_address, mode_up)
                    removed = True
                except Exception as exc:
                    _log(
                        f"_handle_delete_rows: removeRange failed at row {abs_row}: {exc}; clearing cells instead"
                    )
        if not removed:
            for col in range(address.StartColumn, address.EndColumn + 1):
                cell = sheet.getCellByPosition(col, abs_row)
                _write_cell(cell, "")

    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "deleted_rows_count": len(rows_to_delete),
            "deleted_rows_relative_to_target": [int(row) for row in rows_to_delete],
            "structure_shift": "up",
        }
    )
    return sheet.getCellRangeByName(range_name)


_UPDATE_CELLS_DATA_TYPE_FORMATS = {
    "integer": "#,##0",
    "decimal": "#,##0.####;(#,##0.####);-",
    "currency": '"$"* #,##0.00;"$"* (#,##0.00);"$"* -',
    "percentage": "0.0%",
    "date": "YYYY-MM-DD",
}

_POST_WRITE_RECEIPT_LIMIT = 20
_POST_WRITE_VALUE_TEXT_LIMIT = 120


def _build_table_header_map(header: Iterable[Any]) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    for idx, name in enumerate(header):
        text = str(name).strip()
        if text and text not in header_map:
            header_map[text] = idx
        normalized = _normalize_header_label(name)
        if normalized and normalized not in header_map:
            header_map[normalized] = idx
    return header_map


def _resolve_table_header_index(header_map: Dict[str, int], column_name: Any) -> Optional[int]:
    text = str(column_name).strip() if column_name is not None else ""
    if text in header_map:
        return header_map[text]
    normalized = _normalize_header_label(column_name)
    if normalized and normalized in header_map:
        return header_map[normalized]
    return None


def _column_letters_to_index(column_letters: str) -> Optional[int]:
    text = str(column_letters or "").strip().replace("$", "")
    if not text or not re.fullmatch(r"[A-Za-z]{1,3}", text):
        return None
    col_idx = 0
    for ch in text.upper():
        col_idx = col_idx * 26 + (ord(ch) - 64)
    return col_idx - 1


def _resolve_update_cells_column_index(
    header_map: Dict[str, int],
    column_name: Any,
    width: int,
    start_column: int,
) -> Optional[int]:
    header_index = _resolve_table_header_index(header_map, column_name)
    if header_index is not None:
        return header_index
    if column_name is None:
        return None
    text = str(column_name).strip()
    if not text:
        return None
    if re.fullmatch(r"\d+", text):
        index = int(text) - 1
        return index if 0 <= index < width else None
    match = re.fullmatch(r"(?:col(?:umn)?[_ -]?)?(\d+)", text, re.IGNORECASE)
    if match:
        index = int(match.group(1)) - 1
        return index if 0 <= index < width else None
    absolute_column = _column_letters_to_index(text)
    if absolute_column is not None:
        index = absolute_column - start_column
        return index if 0 <= index < width else None
    return None


def _required_update_cells_columns(
    conditions: Iterable[Dict[str, Any]],
    updates: Iterable[Dict[str, Any]],
) -> List[Any]:
    columns: List[Any] = []
    for condition in conditions:
        if isinstance(condition, dict):
            columns.append(condition.get("column"))
    for update in updates:
        if isinstance(update, dict):
            columns.append(update.get("column"))
    return [column for column in columns if str(column or "").strip()]


def _all_update_cells_columns_resolve(
    header_map: Dict[str, int],
    columns: Iterable[Any],
    width: int,
    start_column: int,
) -> bool:
    return all(
        _resolve_update_cells_column_index(header_map, column, width, start_column)
        is not None
        for column in columns
    )


def _all_update_cells_columns_resolve_as_headers(
    header_map: Dict[str, int],
    columns: Iterable[Any],
) -> bool:
    return all(
        _resolve_table_header_index(header_map, column) is not None
        for column in columns
    )


def _find_update_cells_header_context(
    sheet,
    data_range,
    data: List[List[Any]],
    conditions: Iterable[Dict[str, Any]],
    updates: Iterable[Dict[str, Any]],
) -> Dict[str, Any]:
    address = data_range.getRangeAddress()
    width = max(0, int(address.EndColumn) - int(address.StartColumn) + 1)
    required_columns = _required_update_cells_columns(conditions, updates)

    first_row_header = data[0] if data else []
    first_row_header_map = _build_table_header_map(first_row_header)
    if not required_columns or _all_update_cells_columns_resolve_as_headers(
        first_row_header_map, required_columns
    ):
        return {
            "header": first_row_header,
            "header_map": first_row_header_map,
            "rows": data[1:],
            "row_start_offset": 1,
            "header_source": "target_first_row",
            "header_row": int(address.StartRow),
        }

    if _all_update_cells_columns_resolve(
        first_row_header_map, required_columns, width, int(address.StartColumn)
    ):
        return {
            "header": first_row_header,
            "header_map": first_row_header_map,
            "rows": data,
            "row_start_offset": 0,
            "header_source": "target_positional",
            "header_row": int(address.StartRow),
        }

    # Fixed-layout ranges often target only the row to update, while the real
    # header row is above the target in the same columns. Scan nearby rows so
    # an action like range=A10:F10 can still resolve Check/FY2026/FY2027...
    # from A4:F4 and update row 10 in place.
    scan_floor = max(0, int(address.StartRow) - 25)
    for abs_row in range(int(address.StartRow) - 1, scan_floor - 1, -1):
        try:
            candidate_range = sheet.getCellRangeByPosition(
                int(address.StartColumn),
                abs_row,
                int(address.EndColumn),
                abs_row,
            )
            candidate_data = list(map(list, candidate_range.getDataArray()))
        except Exception:
            continue
        if not candidate_data:
            continue
        candidate_header = candidate_data[0]
        candidate_map = _build_table_header_map(candidate_header)
        if _all_update_cells_columns_resolve(
            candidate_map, required_columns, width, int(address.StartColumn)
        ):
            return {
                "header": candidate_header,
                "header_map": candidate_map,
                "rows": data,
                "row_start_offset": 0,
                "header_source": "nearest_header_above",
                "header_row": int(abs_row),
            }

    return {
        "header": first_row_header,
        "header_map": first_row_header_map,
        "rows": data[1:],
        "row_start_offset": 1,
        "header_source": "unresolved",
        "header_row": int(address.StartRow),
    }


def _format_missing_update_cells_columns(kind: str, columns: Iterable[Any], header: Iterable[Any]) -> str:
    missing = [str(item) for item in columns if str(item).strip()]
    header_values = [str(item) for item in header if str(item).strip()]
    return (
        f"update_cells {kind} column(s) not found in the target range header, "
        f"a nearby header row above it, or as positional column references: "
        f"{', '.join(missing)}. Found target first-row values: {header_values[:12]}."
    )


def _cell_ref_from_position(col_index: int, row_index: int) -> str:
    return f"{_column_index_to_name(int(col_index))}{int(row_index) + 1}"


def _compact_post_write_value(value: Any) -> Any:
    if isinstance(value, float):
        if math.isfinite(value) and value.is_integer():
            return int(value)
        return round(value, 6)
    if isinstance(value, str) and len(value) > _POST_WRITE_VALUE_TEXT_LIMIT:
        return value[: _POST_WRITE_VALUE_TEXT_LIMIT - 1] + "…"
    return value


def _cell_type_name(cell) -> str:
    for probe in ("CellContentType", "Type"):
        try:
            value = getattr(cell, probe)
            if value is not None:
                return str(value).upper()
        except Exception:
            pass
    getter = getattr(cell, "getType", None)
    if callable(getter):
        try:
            return str(getter()).upper()
        except Exception:
            pass
    return ""


def _cell_error_code(cell) -> Optional[int]:
    getter = getattr(cell, "getError", None)
    if callable(getter):
        try:
            code = int(getter())
            return code if code else None
        except Exception:
            return None
    return None


def _snapshot_written_cell(cell, wrote_formula: bool = False) -> Dict[str, Any]:
    type_name = _cell_type_name(cell)
    is_formula = bool(wrote_formula or "FORMULA" in type_name)
    entry: Dict[str, Any] = {}

    error_code = _cell_error_code(cell)
    if error_code is not None:
        entry["error"] = error_code

    if is_formula:
        try:
            formula = cell.getFormula()
            if isinstance(formula, str) and formula:
                entry["formula"] = _compact_post_write_value(formula)
        except Exception:
            pass

    value: Any = None
    got_value = False
    if "VALUE" in type_name or is_formula:
        try:
            value = cell.getValue()
            got_value = True
        except Exception:
            got_value = False

    try:
        text = cell.getString()
    except Exception:
        text = ""

    if not got_value:
        value = text

    entry["value"] = _compact_post_write_value(value)
    return entry


def _build_updated_cells_range(sheet, updated_positions: List[Dict[str, Any]]):
    if not updated_positions:
        return None
    try:
        min_col = min(int(item["col"]) for item in updated_positions)
        max_col = max(int(item["col"]) for item in updated_positions)
        min_row = min(int(item["row"]) for item in updated_positions)
        max_row = max(int(item["row"]) for item in updated_positions)
        return sheet.getCellRangeByPosition(min_col, min_row, max_col, max_row)
    except Exception as exc:
        _log(f"_build_updated_cells_range: failed: {exc}")
        return None


def _build_post_write_summary(sheet, updated_positions: List[Dict[str, Any]]) -> Dict[str, Any]:
    groups: List[Dict[str, Any]] = []
    group_by_column: Dict[str, Dict[str, Any]] = {}
    sampled_cells = 0
    formula_error_cells: List[str] = []

    for item in updated_positions:
        column_name = str(item.get("column") or "")
        group = group_by_column.get(column_name)
        if group is None:
            if len(groups) >= _POST_WRITE_RECEIPT_LIMIT:
                continue
            group = {"column": column_name, "cells": []}
            group_by_column[column_name] = group
            groups.append(group)
        if len(group["cells"]) >= _POST_WRITE_RECEIPT_LIMIT:
            continue
        try:
            col = int(item["col"])
            row = int(item["row"])
            cell = sheet.getCellByPosition(col, row)
            ref = _cell_ref_from_position(col, row)
            cell_entry = {"ref": ref}
            cell_entry.update(
                _snapshot_written_cell(cell, bool(item.get("wrote_formula")))
            )
            if cell_entry.get("error") is not None:
                formula_error_cells.append(ref)
            group["cells"].append(cell_entry)
            sampled_cells += 1
        except Exception as exc:
            _log(f"_build_post_write_summary: failed for {item}: {exc}")

    summary: Dict[str, Any] = {
        "total_changed": int(len(updated_positions)),
        "sampled": int(sampled_cells),
        "groups": groups,
    }
    if len(updated_positions) > sampled_cells:
        summary["truncated"] = int(len(updated_positions) - sampled_cells)
    if formula_error_cells:
        summary["formula_error_cells"] = formula_error_cells[:_POST_WRITE_RECEIPT_LIMIT]
    return summary


def _handle_update_cells(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("update_cells requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "update_cells"
    )
    if sheet is None:
        raise RuntimeError(f"update_cells sheet '{sheet_name}' not found")

    data_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, data_range.getDataArray()))
    if not data:
        return None

    conditions = payload.get("find_conditions") or []
    updates = payload.get("updates") or []
    if not updates:
        return _expand_to_single_cell(data_range)

    address = data_range.getRangeAddress()
    width = int(address.EndColumn) - int(address.StartColumn) + 1
    header_context = _find_update_cells_header_context(
        sheet, data_range, data, conditions, updates
    )
    header = header_context["header"]
    header_map = header_context["header_map"]
    rows_to_scan = header_context["rows"]
    row_start_offset = int(header_context["row_start_offset"])

    missing_condition_columns = [
        condition.get("column")
        for condition in conditions
        if _resolve_update_cells_column_index(
            header_map, condition.get("column"), width, int(address.StartColumn)
        )
        is None
    ]
    if missing_condition_columns:
        raise RuntimeError(
            _format_missing_update_cells_columns("find condition", missing_condition_columns, header)
        )

    missing_update_columns_preflight = [
        update.get("column")
        for update in updates
        if _resolve_update_cells_column_index(
            header_map, update.get("column"), width, int(address.StartColumn)
        )
        is None
    ]
    if missing_update_columns_preflight:
        raise RuntimeError(
            _format_missing_update_cells_columns("update", missing_update_columns_preflight, header)
        )

    matched_rows = 0
    updated_cells = 0
    updated_positions: List[Dict[str, Any]] = []
    missing_update_columns: List[str] = []
    # Track affected row offsets per updated column, for post-write formatting below.
    affected_row_offsets_by_column: Dict[str, List[int]] = {}
    for row_offset, row_values in enumerate(rows_to_scan, start=row_start_offset):
        if _row_matches_conditions(
            row_values, header_map, conditions, width, int(address.StartColumn)
        ):
            matched_rows += 1
            for update in updates:
                column_name = update.get("column")
                col_index = _resolve_update_cells_column_index(
                    header_map, column_name, width, int(address.StartColumn)
                )
                if col_index is None:
                    if column_name and column_name not in missing_update_columns:
                        missing_update_columns.append(str(column_name))
                    continue
                abs_col = address.StartColumn + col_index
                abs_row = address.StartRow + row_offset
                cell = sheet.getCellByPosition(
                    abs_col, abs_row
                )
                # Coerce so '50%', '1800000', '1988-03-12' become real numbers/
                # dates instead of TEXT cells where number_format is a no-op.
                coerced = _coerce_value_for_data_type(
                    update.get("new_value"), update.get("data_type")
                )
                _write_cell(cell, coerced)
                updated_cells += 1
                affected_row_offsets_by_column.setdefault(column_name, []).append(row_offset)
                updated_positions.append(
                    {
                        "column": str(column_name),
                        "col": int(abs_col),
                        "row": int(abs_row),
                        "wrote_formula": isinstance(coerced, str)
                        and coerced.strip().startswith("="),
                    }
                )

    # Auto-apply number_format for updates that declared a data_type.
    # Mirrors the openpyxl UpdateCellsHandler path; the live header row is required
    # to resolve column name → column position, so this happens at execution time
    # (not in preflight) and identically in both backends.
    number_formats = None
    locale = None
    for update in updates:
        dt = update.get("data_type")
        if dt == "currency":
            # Free-text currency_code wins over the hardcoded USD default.
            fmt_string = _resolve_currency_format(update.get("currency_code"))
        elif dt:
            fmt_string = _UPDATE_CELLS_DATA_TYPE_FORMATS.get(dt)
        else:
            fmt_string = None
        column_name = update.get("column")
        offsets = affected_row_offsets_by_column.get(column_name) or []
        if not fmt_string or not offsets:
            continue
        if number_formats is None:
            try:
                number_formats = document.getNumberFormats()
                locale = _resolve_locale(number_formats)
            except Exception as exc:
                _log(f"_handle_update_cells: number format init failed: {exc}")
                break
        try:
            fmt_id = number_formats.queryKey(fmt_string, locale, False)
            if fmt_id == -1:
                fmt_id = number_formats.addNew(fmt_string, locale)
        except Exception as exc:
            _log(f"_handle_update_cells: could not register format '{fmt_string}': {exc}")
            continue
        col_index = _resolve_update_cells_column_index(
            header_map, column_name, width, int(address.StartColumn)
        )
        if col_index is None:
            continue
        for row_offset in offsets:
            try:
                cell = sheet.getCellByPosition(
                    address.StartColumn + col_index, address.StartRow + row_offset
                )
                cell.NumberFormat = fmt_id
            except Exception as exc:
                _log(
                    f"_handle_update_cells: could not apply {dt} format at "
                    f"col={column_name} row_offset={row_offset}: {exc}"
                )

    if any(bool(item.get("wrote_formula")) for item in updated_positions):
        _maybe_calculate_formulas(document, controller, force=False)

    applied_range = _build_updated_cells_range(sheet, updated_positions)
    warnings: List[str] = []
    if missing_update_columns:
        warnings.append(
            "update_columns_not_found="
            + ",".join(missing_update_columns[:_POST_WRITE_RECEIPT_LIMIT])
        )
    if matched_rows == 0 and conditions:
        warnings.append("no_rows_matched_find_conditions")

    meta = action.setdefault("_result_meta", {})
    meta_payload = {
        "matched_rows_count": int(matched_rows),
        "updated_cells_count": int(updated_cells),
        "header_source": str(header_context.get("header_source") or ""),
    }
    if applied_range is not None:
        meta_payload["applied_table_range"] = _range_to_representation(applied_range)
        meta_payload["applied_table_range_a1"] = _range_to_a1_relative(applied_range)
        meta_payload["applied_cells_a1"] = [
            _cell_ref_from_position(item["col"], item["row"])
            for item in updated_positions[:_POST_WRITE_RECEIPT_LIMIT]
        ]
    if updated_positions:
        meta_payload["post_write_summary"] = _build_post_write_summary(
            sheet, updated_positions
        )
    if warnings:
        meta_payload["warnings"] = warnings
    meta.update(meta_payload)

    if applied_range is not None:
        return applied_range
    return _expand_to_single_cell(data_range)


def _handle_find_and_replace(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("find_and_replace requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "find_and_replace"
    )
    if sheet is None:
        raise RuntimeError(f"find_and_replace sheet '{sheet_name}' not found")

    data_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, data_range.getDataArray()))
    if not data or len(data) < 2:
        return data_range

    header = data[0]
    column_name = payload.get("column")
    if column_name not in header:
        raise RuntimeError(
            f"find_and_replace column '{column_name}' not in range header"
        )
    column_index = header.index(column_name)
    find_text = payload.get("find_text")
    replace_with = payload.get("replace_with") or ""
    if find_text is None:
        raise RuntimeError("find_and_replace missing find_text value")

    address = data_range.getRangeAddress()
    col_position = address.StartColumn + column_index
    replacements_count = 0
    for row_offset in range(1, len(data)):
        row_idx = address.StartRow + row_offset
        cell = sheet.getCellByPosition(col_position, row_idx)
        try:
            current_value = cell.getString()
        except Exception:
            current_value = ""
        if not current_value or find_text not in current_value:
            continue
        new_value = current_value.replace(find_text, replace_with)
        if new_value != current_value:
            cell.setString(new_value)
            replacements_count += 1

    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "replacements_count": int(replacements_count),
            "replaced_column": str(column_name),
        }
    )

    return data_range


def _handle_manipulate_text_column(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("manipulate_text_column requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "manipulate_text_column"
    )
    if sheet is None:
        raise RuntimeError(f"manipulate_text_column sheet '{sheet_name}' not found")

    data_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, data_range.getDataArray()))
    if not data:
        raise RuntimeError("manipulate_text_column target range is empty")

    header = data[0]
    header_map = {str(name): idx for idx, name in enumerate(header)}

    new_col_name = payload.get("new_column_name")
    source_columns = payload.get("source_columns") or []
    operation = (payload.get("operation") or "").upper()
    separator = payload.get("separator") or ""

    if not new_col_name:
        raise RuntimeError("manipulate_text_column missing new_column_name")
    if not source_columns:
        raise RuntimeError("manipulate_text_column requires source_columns")

    source_indices: list[int] = []
    for column in source_columns:
        if column not in header_map:
            raise RuntimeError(
                f"manipulate_text_column source column '{column}' not found"
            )
        source_indices.append(header_map[column])

    address = data_range.getRangeAddress()
    position, new_col_index, template_col_index = _resolve_column_insert_location(
        header, address, payload, "manipulate_text_column"
    )
    if position != "END":
        inserted = _insert_columns_via_sheet(
            sheet, new_col_index, 1, address.StartRow, address.EndRow
        )
        if not inserted:
            inserted = _shift_columns_right_manual(
                sheet,
                new_col_index,
                address.EndColumn,
                address.StartRow,
                address.EndRow,
                1,
            )
        if not inserted:
            try:
                sheet.getColumns().insertByIndex(new_col_index, 1)
                inserted = True
            except Exception as exc:
                _log(f"manipulate_text_column: insertByIndex fallback failed: {exc}")
        if not inserted:
            raise RuntimeError(
                "manipulate_text_column failed to insert a blank column slot"
            )
    header_row = address.StartRow
    if position == "END":
        template_header = sheet.getCellByPosition(template_col_index, header_row)
        if not _cell_has_visible_value(template_header):
            template_col_index = max(address.StartColumn, template_col_index - 1)

    try:
        _clone_column_format(
            sheet, template_col_index, new_col_index, address.StartRow, address.EndRow
        )
    except Exception:
        _copy_column_styles(
            sheet, template_col_index, new_col_index, address.StartRow, address.EndRow
        )

    header_cell = sheet.getCellByPosition(new_col_index, header_row)
    header_cell.setString(new_col_name)
    table_end_col = max(address.EndColumn + 1, new_col_index)
    if _ensure_default_header_style_if_missing(
        sheet, address.StartColumn, header_row, table_end_col
    ):
        _log(
            "_handle_manipulate_text_column: applied default header style "
            f"row={header_row} cols={address.StartColumn}-{table_end_col}"
        )

    def _to_text(value: Any) -> str:
        if value is None:
            return ""
        return str(value)

    results: list[str] = []
    for row_values in data[1:]:
        values = []
        for idx in source_indices:
            if idx < len(row_values):
                values.append(_to_text(row_values[idx]))
            else:
                values.append("")
        if operation == "CONCAT":
            results.append(separator.join(values))
        elif operation == "LOWER":
            if len(values) != 1:
                raise RuntimeError("LOWER operation requires exactly one source column")
            results.append(values[0].lower())
        elif operation == "UPPER":
            if len(values) != 1:
                raise RuntimeError("UPPER operation requires exactly one source column")
            results.append(values[0].upper())
        elif operation == "TRIM":
            if len(values) != 1:
                raise RuntimeError("TRIM operation requires exactly one source column")
            results.append(values[0].strip())
        elif operation == "PROPER":
            if len(values) != 1:
                raise RuntimeError("PROPER operation requires exactly one source column")
            results.append(values[0].title())
        else:
            raise RuntimeError(
                f"manipulate_text_column unsupported operation '{operation}'"
            )

    for row_offset, result in enumerate(results, start=1):
        row_index = header_row + row_offset
        cell = sheet.getCellByPosition(new_col_index, row_index)
        template_cell = sheet.getCellByPosition(template_col_index, row_index)
        _copy_cell_style(template_cell, cell)
        cell.setString(result)

    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "created_column_name": new_col_name,
            "created_column_index_1based": int(new_col_index) + 1,
            "created_data_row_count": int(len(results)),
        }
    )
    end_row = header_row + max(0, len(results))
    return sheet.getCellRangeByPosition(
        new_col_index,
        header_row,
        new_col_index,
        end_row,
    )


def _iterate_cells(sheet, cell_range):
    """Yield ACTUAL cell objects (ScCellObj) for every cell in *cell_range*.

    ``cell_range.getCells().createEnumeration()`` returns proper cell objects
    when the range is a multi-cell range.  Single-cell ranges sometimes do
    NOT expose ``getCells()`` — and the previously-used fallback
    ``cell_range.createEnumeration()`` is dangerous: it yields
    ``SvxUnoTextContent`` (text portions inside the cell), not cells.
    Setting CharColor / CharWeight / etc. on a text portion of a NUMERIC
    cell silently commits the cell as TEXT (e.g. 85000 → text "85000",
    losing the numeric value and breaking currency formats).

    Position-based iteration via ``sheet.getCellByPosition`` always yields
    real cells, so it is the safe fallback.
    """
    try:
        enumerator = cell_range.getCells().createEnumeration()
        while enumerator.hasMoreElements():
            yield enumerator.nextElement()
        return
    except Exception:
        pass

    try:
        address = cell_range.getRangeAddress()
    except Exception:
        return
    for col in range(address.StartColumn, address.EndColumn + 1):
        for row in range(address.StartRow, address.EndRow + 1):
            try:
                yield sheet.getCellByPosition(col, row)
            except Exception:
                continue


def _handle_apply_formatting(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("apply_formatting requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "apply_formatting"
    )
    if sheet is None:
        raise RuntimeError(f"apply_formatting sheet '{sheet_name}' not found")

    # Try to access as a single range first
    try:
        cell_range = sheet.getCellRangeByName(range_name)
        skip_entry = _find_duplicate_header_write_skip(sheet, cell_range)
        action_name = str(action.get("action_name") or "")
        if skip_entry and action_name == "Header Styling Fallback":
            meta = action.setdefault("_result_meta", {})
            meta["skipped_due_to_duplicate_header_write"] = True
            meta["skipped_header_range_a1"] = skip_entry.get(
                "skipped_header_range_a1"
            )
            meta["existing_header_range_a1"] = skip_entry.get(
                "existing_header_range_a1"
            )
            _log(
                "_handle_apply_formatting: skipped Header Styling Fallback "
                f"on duplicate WRITE_DATA header row {range_name}; existing "
                f"header is {skip_entry.get('existing_header_range_a1')}"
            )
            return cell_range
        _apply_format_to_range_cells(sheet, cell_range, payload, document)
        return cell_range
    except Exception:
        # Fallback for comma-separated ranges (e.g., "A1:B2,D5")
        if "," in range_name:
            sub_ranges = [r.strip() for r in range_name.split(",")]
            for sub_range in sub_ranges:
                if not sub_range:
                    continue
                try:
                    rng = sheet.getCellRangeByName(sub_range)
                    _apply_format_to_range_cells(sheet, rng, payload, document)
                except Exception as e:
                    _log(f"_handle_apply_formatting: skipping invalid sub-range '{sub_range}': {e}")
            return None  # Can't return a single range object for multi-selection
        else:
            raise

def _apply_format_to_range_cells(sheet, cell_range, payload, document):
    format_spec = dict(payload.get("format") or {})
    document_formats = document.getNumberFormats()
    locale = _resolve_locale(document_formats)
    force_no_wrap = False
    range_address = None

    try:
        range_address = cell_range.getRangeAddress()
        force_no_wrap = (
            range_address.StartRow == range_address.EndRow
            and _is_header_like_format(format_spec)
        )
        if force_no_wrap:
            format_spec["_force_no_wrap"] = True
    except Exception:
        range_address = None

    border_spec = format_spec.get("border") or {}
    if border_spec:
        try:
            address = range_address or cell_range.getRangeAddress()
            raw_color = border_spec.get("color")
            color_int = 0xC0C0C0
            if raw_color:
                value = str(raw_color).strip().lstrip("#")
                color_int = int(value, 16)
            width = _border_width_from_style(border_spec.get("style"))
            outline = border_spec.get("outline", True)
            inside = border_spec.get("inside", True)
            _apply_cell_borders(
                sheet,
                address.StartColumn,
                address.StartRow,
                address.EndColumn,
                address.EndRow,
                color=color_int,
                width=width,
                outline=outline,
                inside=inside,
            )
        except Exception as exc:
            _log(f"_apply_format_to_range_cells: border error - {exc}")

    for cell in _iterate_cells(sheet, cell_range):
        try:
            _apply_format_to_cell(cell, format_spec, document_formats, locale)
        except Exception:
            # Best-effort formatting: fail silently per-cell.
            continue
    if force_no_wrap and range_address is not None:
        _reset_row_optimal_height(sheet, range_address.StartRow)


def _handle_insert_comment(document, controller, action: Dict[str, Any]):
    """Handle insert_comment action - adds AI commentary as a cell comment."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = target_info.get("sheet_name") or "Sheet1"
    # Default sheet logic if not explicitly provided or found
    sheet = None
    cell_ref = target_info.get("range") or "A1"
    sheet, cell_ref = _resolve_sheet_and_range_name(
        document, sheet_name, cell_ref, "insert_comment"
    )
    if sheet is None and sheet_name:
        sheet = _get_sheet(document, sheet_name)

    if sheet is None:
        # Fallback: use active sheet or first sheet
        try:
            sheet = document.getCurrentController().getActiveSheet()
            _log(f"_handle_insert_comment: sheet '{sheet_name}' not found, using active sheet '{sheet.getName()}'")
        except Exception:
            _log(f"_handle_insert_comment: sheet '{sheet_name}' not found and no active sheet")
            return None

    author = payload.get("author") or "SmartDocs AI"
    text = payload.get("text") or ""

    if not text:
        _log("_handle_insert_comment: no text provided, skipping")
        return None

    # Handle multi-range references (e.g. "A1:B2,C5") by picking the first cell
    effective_cell_ref = cell_ref
    if "," in cell_ref:
        effective_cell_ref = cell_ref.split(",")[0].strip()

    # If it's a range (A1:B2), LibreOffice usually wants a single cell for annotation.
    # We try to get the range, then pick the top-left cell.
    try:
        rng = sheet.getCellRangeByName(effective_cell_ref)
        address = rng.getRangeAddress()
        # Top-left cell
        col = address.StartColumn
        row = address.StartRow
        # Convert back to A1 for logging/debugging or use getCellByPosition
        # But _add_range_comment expects a range ref string or we modify it to take cell object
        # Let's modify usage of _add_range_comment or just fetch the cell object there.
        # Simpler: just use the top-left cell object for annotation add.
        target_cell = sheet.getCellByPosition(col, row)
        _add_comment_to_cell_obj(sheet, target_cell, text, author)
    except Exception as e:
         _log(f"_handle_insert_comment: failed to resolve target '{effective_cell_ref}': {e}")
         return None

    _log(f"_handle_insert_comment: added comment at {sheet.getName()}!{effective_cell_ref}")
    return None

def _add_comment_to_cell_obj(sheet, cell, text, author):
    """Add comment to a specific cell object."""
    try:
        # Set user profile so annotation Author shows our brand name
        _set_user_profile_for_author(author)
        
        # Get annotation - LibreOffice returns an annotation object even if empty
        annotation = cell.getAnnotation()
        current_text = ""
        try:
            current_text = annotation.getString() if annotation else ""
        except Exception:
            current_text = ""

        # Build comment text (author is set via user profile, not in text)
        full_text = text
        if current_text:
            full_text = f"{current_text}\n---\n{text}"

        if current_text:
            # Update existing annotation
            annotation.setString(full_text)
        else:
            # No existing content - create new annotation
            # Author is set via _set_user_profile_for_author() called earlier
            annotations = sheet.getAnnotations()
            cell_address = cell.getCellAddress()
            annotations.insertNew(cell_address, full_text)

        _log(f"_add_comment_to_cell_obj: successfully added comment")
    except Exception as e:
        _log(f"_add_comment_to_cell_obj failed: {e}")


def _merge_range_unlocked(document, cell_range) -> bool:
    """Merge a cell range with controllers briefly unlocked.

    lockControllers() suppresses the UNO broadcast that makes merge() take
    effect in Collabora Online.  Unlocking for the single merge call and
    immediately re-locking is safe: no other writer can intervene inside the
    same Python frame.  Returns True if the merge call succeeded without an
    exception (the visual effect still depends on the Collabora renderer).
    """
    try:
        document.unlockControllers()
    except Exception:
        pass
    ok = True
    try:
        cell_range.merge(True)
    except Exception as exc:
        _log(f"_merge_range_unlocked: merge raised {type(exc).__name__}: {exc}")
        ok = False
    try:
        document.lockControllers()
    except Exception:
        pass
    return ok


def _handle_add_title(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    cell_ref = payload.get("target_cell") or target_info.get("range")
    if not cell_ref:
        raise RuntimeError("add_title requires target_cell")
    sheet, cell_ref = _resolve_sheet_and_range_name(
        document, sheet_name, cell_ref, "add_title"
    )
    if sheet is None:
        raise RuntimeError(f"add_title sheet '{sheet_name}' not found")

    try:
        target_range = sheet.getCellRangeByName(cell_ref)
    except Exception as exc:
        raise RuntimeError(
            f"add_title invalid target_cell '{cell_ref}': {exc}"
        ) from exc

    text = payload.get("text") or ""
    address = target_range.getRangeAddress()
    top_cell = sheet.getCellByPosition(address.StartColumn, address.StartRow)
    top_cell.setString(text)

    # Determine the target merge range: explicit `range` takes priority,
    # otherwise auto-expand based on sheet data width / text length.
    merge_range_name = payload.get("range")
    final_merge_range = None
    if merge_range_name:
        try:
            _, merge_range_name = _resolve_sheet_and_range_name(
                document, sheet_name, merge_range_name, "add_title.range"
            )
            final_merge_range = sheet.getCellRangeByName(merge_range_name)
        except Exception as exc:
            _log(
                f"_handle_add_title: range path failed for '{merge_range_name}': "
                f"{type(exc).__name__}: {exc}"
            )
    else:
        final_merge_range = _calculate_title_expand_range(sheet, target_range, text)

    # Perform the merge.  lockControllers() suppresses the UNO broadcast
    # required for merge() to take effect in Collabora; unlock briefly.
    if final_merge_range is not None:
        merge_addr = final_merge_range.getRangeAddress()
        ok = _merge_range_unlocked(document, final_merge_range)
        _log(
            f"_handle_add_title: merge {'ok' if ok else 'failed'} "
            f"end_col={merge_addr.EndColumn}"
        )
        target_range = final_merge_range
        top_cell = sheet.getCellByPosition(merge_addr.StartColumn, merge_addr.StartRow)
        top_cell.setString(text)

    format_spec = payload.get("format") or {}
    if format_spec:
        document_formats = document.getNumberFormats()
        locale = _resolve_locale(document_formats)
        for cell in _iterate_cells(sheet, target_range):
            _apply_format_to_cell(cell, format_spec, document_formats, locale)

    return target_range


def _sheet_last_used_column(sheet) -> int:
    """Return the 0-based last column index that holds data on `sheet`.
    Returns -1 if the sheet is empty."""
    try:
        cursor = sheet.createCursor()
        cursor.gotoStartOfUsedArea(False)
        cursor.gotoEndOfUsedArea(True)
        addr = cursor.getRangeAddress()
        if addr.EndColumn >= 0 and addr.EndRow >= 0:
            # If used area is just A1 and A1 is empty, treat as empty sheet.
            top_left = sheet.getCellByPosition(addr.StartColumn, addr.StartRow)
            if (
                addr.StartColumn == 0
                and addr.StartRow == 0
                and addr.EndColumn == 0
                and addr.EndRow == 0
                and not (top_left.getString() or "").strip()
            ):
                return -1
            return addr.EndColumn
    except Exception as exc:
        _log(f"_sheet_last_used_column: failed - {exc}")
    return -1


def _calculate_title_expand_range(sheet, cell_range, text: str):
    """Return the cell range that the title should span (without merging).

    Prefers the sheet's used-data width over a text-length heuristic so that
    titles automatically span a full data table.  Merging is the caller's
    responsibility (see _handle_add_title / _merge_range_unlocked).
    """
    if not text:
        return cell_range
    address = cell_range.getRangeAddress()
    current_cols = address.EndColumn - address.StartColumn + 1
    max_columns = sheet.getColumns().getCount()

    text_based_cols = max(1, math.ceil(len(text) / 12))
    if current_cols <= 1:
        last_used_col = _sheet_last_used_column(sheet)
        if last_used_col >= address.StartColumn:
            data_based_cols = last_used_col - address.StartColumn + 1
            desired_cols = max(text_based_cols, data_based_cols)
        else:
            desired_cols = text_based_cols
    else:
        desired_cols = max(current_cols, text_based_cols)

    new_end_col = min(address.StartColumn + desired_cols - 1, max_columns - 1)
    if new_end_col <= address.EndColumn:
        return cell_range
    return sheet.getCellRangeByPosition(
        address.StartColumn, address.StartRow, new_end_col, address.EndRow
    )


def _handle_apply_conditional_formatting(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")

    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("apply_conditional_formatting requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "apply_conditional_formatting"
    )
    if sheet is None:
        raise RuntimeError(
            f"apply_conditional_formatting sheet '{sheet_name}' not found"
        )

    cell_range = sheet.getCellRangeByName(range_name)
    rule = payload.get("rule") or {}
    rule_type = (rule.get("rule_type") or "CELL_VALUE").upper()

    document_formats = document.getNumberFormats()
    locale = _resolve_locale(document_formats)

    if rule_type == "COLOR_SCALE":
        colors = rule.get("colors") or []
        midpoint_value = rule.get("midpoint_value")
        _apply_color_scale(sheet, cell_range, colors, midpoint_value)
        return cell_range

    format_spec = rule.get("format") or {}
    if not format_spec:
        _log("_handle_apply_conditional_formatting: missing format spec; skipping")
        return cell_range

    if rule_type == "TOP_BOTTOM":
        mode = (rule.get("mode") or "").upper()
        threshold_value = rule.get("value")
        _apply_top_bottom_rule(
            sheet,
            cell_range,
            mode,
            threshold_value,
            format_spec,
            document_formats,
            locale,
        )
        return cell_range

    operator = rule.get("operator")
    value = rule.get("value")
    if not operator:
        _log("_handle_apply_conditional_formatting: missing operator; skipping")
        return cell_range

    durable_ok = _try_apply_durable_conditional_format(
        document,
        cell_range,
        rule_type,
        operator,
        value,
        format_spec,
        document_formats,
        locale,
    )
    if durable_ok:
        try:
            document.setModified(True)
        except Exception:
            pass
        return cell_range

    def _predicate(cell_value: Any) -> bool:
        return _compare_values(cell_value, operator, value)

    _apply_predicate_format(
        sheet,
        cell_range,
        _predicate,
        format_spec,
        document_formats,
        locale,
    )

    return cell_range


def _try_apply_durable_conditional_format(
    document,
    cell_range,
    rule_type: str,
    operator: str,
    value: Any,
    format_spec: Dict[str, Any],
    document_formats,
    locale,
) -> bool:
    """Add a real Calc conditional format that survives XLSX export.

    The previous implementation eagerly styled matching cells, which looked
    correct in-session but exported no ``<conditionalFormatting>`` rule.
    Calc's durable conditional formatting stores a condition entry that points
    at a cell style; the XLSX filter converts that to a differential format.
    """
    if rule_type not in {"CELL_VALUE", "TEXT"}:
        return False
    op = (operator or "").upper()
    operator_value = _condition_operator_for_rule(rule_type, op)
    if operator_value is None:
        _log(
            "_try_apply_durable_conditional_format: unsupported "
            f"rule_type={rule_type} operator={op}"
        )
        return False

    formula1, formula2 = _condition_formula_values(op, value)
    if formula1 is None:
        return False

    try:
        style_name = _ensure_conditional_cell_style(
            document, format_spec, document_formats, locale
        )
        address = cell_range.getRangeAddress()
        source = uno.createUnoStruct("com.sun.star.table.CellAddress")
        source.Sheet = address.Sheet
        source.Column = address.StartColumn
        source.Row = address.StartRow

        entries = cell_range.getPropertyValue("ConditionalFormat")
        props = [
            _property_value("Operator", operator_value),
            _property_value("Formula1", formula1),
            _property_value("SourcePosition", source),
            _property_value("StyleName", style_name),
        ]
        if formula2 is not None:
            props.append(_property_value("Formula2", formula2))
        entries.addNew(tuple(props))
        cell_range.setPropertyValue("ConditionalFormat", entries)
        _log(
            "_try_apply_durable_conditional_format: added durable "
            f"{rule_type}/{op} rule style={style_name}"
        )
        return True
    except Exception as exc:
        _log(f"_try_apply_durable_conditional_format: failed: {exc}")
        return False


def _property_value(name: str, value: Any):
    prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    prop.Name = name
    prop.Value = value
    return prop


def _condition_operator_for_rule(rule_type: str, operator: str):
    try:
        from com.sun.star.sheet.ConditionOperator import (
            BETWEEN,
            EQUAL,
            GREATER,
            GREATER_EQUAL,
            LESS,
            LESS_EQUAL,
            NOT_BETWEEN,
            NOT_EQUAL,
        )
    except Exception as exc:
        _log(f"_condition_operator_for_rule: import failed: {exc}")
        return None

    if rule_type == "TEXT" and operator not in {"EQUALS", "NOT_EQUALS"}:
        return None
    return {
        "EQUALS": EQUAL,
        "NOT_EQUALS": NOT_EQUAL,
        "GREATER_THAN": GREATER,
        "LESS_THAN": LESS,
        "GREATER_OR_EQUAL": GREATER_EQUAL,
        "LESS_OR_EQUAL": LESS_EQUAL,
        "BETWEEN": BETWEEN,
        "NOT_BETWEEN": NOT_BETWEEN,
    }.get(operator)


def _condition_formula_values(operator: str, value: Any) -> Tuple[Optional[str], Optional[str]]:
    if operator in {"BETWEEN", "NOT_BETWEEN"}:
        if not isinstance(value, (list, tuple)) or len(value) < 2:
            return None, None
        return _condition_formula_literal(value[0]), _condition_formula_literal(value[1])
    return _condition_formula_literal(value), None


def _condition_formula_literal(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE()" if value else "FALSE()"
    if isinstance(value, (int, float)):
        return str(value)
    text = str(value)
    if not text:
        return '""'
    if text.startswith("="):
        return text[1:]
    try:
        float(text.replace(",", ""))
        return text.replace(",", "")
    except Exception:
        pass
    escaped = text.replace('"', '""')
    return f'"{escaped}"'


def _ensure_conditional_cell_style(
    document,
    format_spec: Dict[str, Any],
    document_formats,
    locale,
) -> str:
    styles = document.getStyleFamilies().getByName("CellStyles")
    normalized = json.dumps(format_spec or {}, sort_keys=True, default=str)
    digest = hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:12]
    style_name = f"SmartDocs_CF_{digest}"
    try:
        if not styles.hasByName(style_name):
            style = document.createInstance("com.sun.star.style.CellStyle")
            styles.insertByName(style_name, style)
        else:
            style = styles.getByName(style_name)
    except Exception:
        style_name = f"SmartDocs_CF_{int(time.time() * 1000)}"
        style = document.createInstance("com.sun.star.style.CellStyle")
        styles.insertByName(style_name, style)
    _apply_format_to_cell(style, format_spec, document_formats, locale)
    return style_name


def _apply_predicate_format(
    sheet,
    cell_range,
    predicate: Callable[[Any], bool],
    format_spec,
    document_formats,
    locale,
):
    enumerator = None
    try:
        enumerator = cell_range.getCells().createEnumeration()
    except AttributeError:
        enumerator = None
    except Exception as exc:
        _log(f"_apply_predicate_format: enumeration failed: {exc}")
        enumerator = None

    if enumerator is not None:
        while enumerator.hasMoreElements():
            cell = enumerator.nextElement()
            cell_value = _extract_cell_value(cell)
            try:
                if predicate(cell_value):
                    _apply_format_to_cell(cell, format_spec, document_formats, locale)
            except Exception as exc:
                _log(f"_apply_predicate_format: predicate failed: {exc}")
    else:
        address = cell_range.getRangeAddress()
        for col in range(address.StartColumn, address.EndColumn + 1):
            for row in range(address.StartRow, address.EndRow + 1):
                cell = sheet.getCellByPosition(col, row)
                cell_value = _extract_cell_value(cell)
                try:
                    if predicate(cell_value):
                        _apply_format_to_cell(
                            cell, format_spec, document_formats, locale
                        )
                except Exception as exc:
                    _log(f"_apply_predicate_format: predicate iteration failed: {exc}")


def _apply_top_bottom_rule(
    sheet, cell_range, mode: str, threshold_value, format_spec, document_formats, locale
):
    if not mode:
        _log("_apply_top_bottom_rule: missing mode; skipping")
        return

    numeric_cells: List[Tuple[Any, float]] = []
    for cell in _iterate_cells(sheet, cell_range):
        numeric_val = _coerce_number(_extract_cell_value(cell))
        if numeric_val is not None:
            numeric_cells.append((cell, numeric_val))

    if not numeric_cells:
        _log("_apply_top_bottom_rule: no numeric cells found in range; skipping")
        return

    values_only = [val for _, val in numeric_cells]
    total = len(values_only)
    mode_upper = mode.upper()
    selected_cells: List[Any] = []

    if mode_upper == "TOP_N":
        count = min(total, _resolve_rank_count(threshold_value, default=10))
        cutoff = _resolve_cutoff(values_only, count, reverse=True)
        selected_cells = [cell for cell, val in numeric_cells if val >= cutoff]
    elif mode_upper == "BOTTOM_N":
        count = min(total, _resolve_rank_count(threshold_value, default=10))
        cutoff = _resolve_cutoff(values_only, count, reverse=False)
        selected_cells = [cell for cell, val in numeric_cells if val <= cutoff]
    elif mode_upper == "TOP_PERCENT":
        count = min(total, _resolve_percent_count(threshold_value, total, default=10.0))
        cutoff = _resolve_cutoff(values_only, count, reverse=True)
        selected_cells = [cell for cell, val in numeric_cells if val >= cutoff]
    elif mode_upper == "BOTTOM_PERCENT":
        count = min(total, _resolve_percent_count(threshold_value, total, default=10.0))
        cutoff = _resolve_cutoff(values_only, count, reverse=False)
        selected_cells = [cell for cell, val in numeric_cells if val <= cutoff]
    elif mode_upper == "ABOVE_AVERAGE":
        average = sum(values_only) / total
        selected_cells = [cell for cell, val in numeric_cells if val > average]
    elif mode_upper == "BELOW_AVERAGE":
        average = sum(values_only) / total
        selected_cells = [cell for cell, val in numeric_cells if val < average]
    else:
        _log(f"_apply_top_bottom_rule: unsupported mode '{mode_upper}'")
        return

    for cell in selected_cells:
        try:
            _apply_format_to_cell(cell, format_spec, document_formats, locale)
        except Exception:
            # Best-effort formatting: fail silently per-cell.
            continue


def _apply_color_scale(
    sheet, cell_range, colors: List[str], midpoint_value: Optional[float]
):
    if not colors or len(colors) < 2:
        _log("_apply_color_scale: at least two colors are required")
        return

    color_stops: List[Tuple[int, int, int]] = []
    for color in colors:
        rgb = _hex_to_rgb_tuple(color)
        if rgb is not None:
            color_stops.append(rgb)
        else:
            _log(f"_apply_color_scale: invalid color '{color}' ignored")

    if len(color_stops) < 2:
        _log("_apply_color_scale: no valid colors supplied; skipping")
        return

    numeric_cells: List[Tuple[Any, Optional[float]]] = []
    numeric_values: List[float] = []

    for cell in _iterate_cells(sheet, cell_range):
        numeric_val = _coerce_number(_extract_cell_value(cell))
        numeric_cells.append((cell, numeric_val))
        if numeric_val is not None:
            numeric_values.append(numeric_val)

    if not numeric_values:
        _log("_apply_color_scale: no numeric values found; skipping")
        return

    min_val = min(numeric_values)
    max_val = max(numeric_values)
    if math.isclose(min_val, max_val):
        max_val = min_val + 1.0

    midpoint = midpoint_value
    if len(color_stops) == 3 and midpoint is None:
        midpoint = _median(numeric_values)

    for cell, numeric_val in numeric_cells:
        if numeric_val is None:
            continue

        if len(color_stops) == 2:
            ratio = (numeric_val - min_val) / (max_val - min_val)
            ratio = _clamp(ratio)
            rgb = _lerp_color(color_stops[0], color_stops[1], ratio)
        else:
            assert len(color_stops) == 3
            lower_bound = min_val
            upper_bound = max_val
            mid_value = midpoint
            if mid_value is None or mid_value <= lower_bound:
                mid_value = lower_bound
            if mid_value >= upper_bound:
                mid_value = upper_bound

            if numeric_val <= mid_value:
                denom = max(mid_value - lower_bound, 1e-9)
                ratio = (numeric_val - lower_bound) / denom
                ratio = _clamp(ratio)
                rgb = _lerp_color(color_stops[0], color_stops[1], ratio)
            else:
                denom = max(upper_bound - mid_value, 1e-9)
                ratio = (numeric_val - mid_value) / denom
                ratio = _clamp(ratio)
                rgb = _lerp_color(color_stops[1], color_stops[2], ratio)

        try:
            _set_uno_property(cell, "IsCellBackgroundTransparent", False)
            _set_uno_property(cell, "CellBackColor", _rgb_to_int(rgb))
        except Exception:
            # Best-effort formatting: fail silently per-cell.
            continue


def _resolve_rank_count(raw_value, default: int = 10) -> int:
    try:
        count = int(float(raw_value))
    except (TypeError, ValueError):
        count = default
    return max(1, count)


def _resolve_percent_count(raw_value, total: int, default: float = 10.0) -> int:
    try:
        percent = float(raw_value)
    except (TypeError, ValueError):
        percent = default
    percent = _clamp(percent, 0.01, 100.0)
    count = int(math.ceil((percent / 100.0) * total))
    return max(1, count)


def _resolve_cutoff(values: List[float], count: int, reverse: bool) -> float:
    ordered = sorted(values, reverse=reverse)
    index = min(max(count - 1, 0), len(ordered) - 1)
    return ordered[index]


def _hex_to_rgb_tuple(color: str) -> Optional[Tuple[int, int, int]]:
    if not color:
        return None
    value = str(color).strip()
    if value.startswith("#"):
        value = value[1:]
    if len(value) == 8:
        value = value[2:]
    if len(value) != 6:
        return None
    try:
        r = int(value[0:2], 16)
        g = int(value[2:4], 16)
        b = int(value[4:6], 16)
        return (r, g, b)
    except ValueError:
        return None


def _rgb_to_int(rgb: Tuple[int, int, int]) -> int:
    r, g, b = rgb
    return (
        (max(0, min(255, r)) << 16) | (max(0, min(255, g)) << 8) | max(0, min(255, b))
    )


def _lerp_color(
    a: Tuple[int, int, int], b: Tuple[int, int, int], ratio: float
) -> Tuple[int, int, int]:
    ratio = _clamp(ratio)
    ar, ag, ab = a
    br, bg, bb = b
    return (
        int(round(ar + (br - ar) * ratio)),
        int(round(ag + (bg - ag) * ratio)),
        int(round(ab + (bb - ab) * ratio)),
    )


def _median(values: List[float]) -> float:
    ordered = sorted(values)
    length = len(ordered)
    mid = length // 2
    if length % 2 == 0:
        return (ordered[mid - 1] + ordered[mid]) / 2.0
    return ordered[mid]


def _clamp(value: float, minimum: float = 0.0, maximum: float = 1.0) -> float:
    if value is None:
        return minimum
    if value < minimum:
        return minimum
    if value > maximum:
        return maximum
    return value


def _handle_sort_range(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("sort_range requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "sort_range"
    )
    if sheet is None:
        raise RuntimeError(f"sort_range sheet '{sheet_name}' not found")

    cell_range = sheet.getCellRangeByName(range_name)
    data = list(map(list, cell_range.getDataArray()))
    if not data:
        return cell_range

    header = data[0]
    header_map = {str(name): idx for idx, name in enumerate(header)}

    sort_specs = payload.get("sort_by") or []
    if not sort_specs:
        return cell_range

    descriptor = list(cell_range.createSortDescriptor())
    sort_fields = []
    for spec in sort_specs:
        column_name = spec.get("column")
        if column_name not in header_map:
            continue
        field = uno.createUnoStruct("com.sun.star.table.TableSortField")
        field.Field = header_map[column_name]
        field.IsAscending = spec.get("order", "ascending").lower() != "descending"
        col_index = header_map[column_name]
        sample_values = [row[col_index] for row in data[1:] if col_index < len(row)]
        numeric_candidates = [v for v in sample_values if v not in (None, "")]
        is_numeric = bool(numeric_candidates) and all(
            _coerce_number(v) is not None for v in numeric_candidates
        )
        type_name = "NUMERIC" if is_numeric else "ALPHANUMERIC"
        try:
            field.FieldType = uno.getConstantByName(
                f"com.sun.star.table.TableSortFieldType.{type_name}"
            )
        except Exception:
            try:
                field.FieldType = uno.getConstantByName(
                    "com.sun.star.table.TableSortFieldType.AUTOMATIC"
                )
            except Exception:
                pass
        sort_fields.append(field)

    if not sort_fields:
        return cell_range

    for prop in descriptor:
        if prop.Name == "ContainsHeader":
            prop.Value = True
        if prop.Name == "SortFields":
            try:
                prop.Value = uno.Any(
                    "[]com.sun.star.table.TableSortField", tuple(sort_fields)
                )
            except Exception:
                prop.Value = tuple(sort_fields)

    cell_range.sort(tuple(descriptor))
    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "sort_keys": [
                {
                    "column": str(spec.get("column")),
                    "order": str(spec.get("order", "ascending")),
                }
                for spec in sort_specs
                if isinstance(spec, dict) and spec.get("column")
            ],
            "sorted_with_header": True,
        }
    )
    return cell_range


def _handle_create_chart(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}

    destination_sheet_name = (
        payload.get("destination_sheet")
        or target_info.get("sheet_name")
        or (payload.get("target") or {}).get("sheet_name")
    )
    if not destination_sheet_name:
        raise RuntimeError(
            "create_chart requires destination_sheet in payload or target"
        )

    dest_sheet = _get_sheet(document, destination_sheet_name)
    if dest_sheet is None:
        raise RuntimeError(
            f"create_chart destination sheet '{destination_sheet_name}' not found"
        )

    data_range_ref = payload.get("data_target_range")
    category_range_ref = payload.get("category_target_range")
    if not data_range_ref or not category_range_ref:
        raise RuntimeError(
            "create_chart requires data_target_range and category_target_range"
        )

    data_sheet, data_range = _resolve_sheet_and_range(
        document, data_range_ref, dest_sheet
    )
    category_sheet, category_range = _resolve_sheet_and_range(
        document, category_range_ref, dest_sheet
    )

    if data_range is None or data_sheet is None:
        raise RuntimeError(
            f"create_chart could not resolve data range '{data_range_ref}'"
        )
    if category_range is None or category_sheet is None:
        raise RuntimeError(
            f"create_chart could not resolve category range '{category_range_ref}'"
        )

    data_range, data_header_extended = _maybe_extend_with_header(data_range)
    category_range, cat_header_extended = _maybe_extend_with_header(category_range)
    header_was_extended = data_header_extended or cat_header_extended

    ranges_for_chart = []
    ranges_for_chart.append(data_range.getRangeAddress())

    combined_range = None
    if data_sheet == category_sheet:
        combined_range = _combine_ranges(data_range, category_range)
        if combined_range is not None:
            ranges_for_chart = [combined_range.getRangeAddress()]
        else:
            ranges_for_chart.append(category_range.getRangeAddress())
    else:
        ranges_for_chart.append(category_range.getRangeAddress())

    native_placement = payload.get("placement")
    if not isinstance(native_placement, dict):
        raise RuntimeError("create_chart requires native placement")

    chart_rect, chart_area_range = _chart_rect_and_area_from_placement(
        document,
        dest_sheet,
        native_placement,
    )
    if chart_rect is None:
        raise RuntimeError("create_chart could not resolve native placement")
    if chart_area_range is None and chart_rect is not None:
        chart_area_range = _cell_range_from_chart_rect(dest_sheet, chart_rect)
    collision_details = _chart_rect_collision_details(
        dest_sheet,
        chart_rect,
        requested_placement=native_placement,
        requested_area=chart_area_range,
    )
    if collision_details:
        raise ActionApplicationError(
            "create_chart placement overlaps existing sheet content or another chart",
            error_code="CHART_PLACEMENT_COLLISION",
            details=collision_details,
        )

    charts = dest_sheet.getCharts()
    chart_title = payload.get("title") or "Chart"
    chart_name = _unique_chart_name(charts, chart_title)

    range_addresses = tuple(addr for addr in ranges_for_chart if addr is not None)
    if not range_addresses:
        raise RuntimeError("create_chart could not resolve any data ranges")

    # Heuristically decide whether the first row/column are labels.
    # CRITICAL: if _maybe_extend_with_header grew the ranges to include a
    # header row, the first row IS labels regardless of what the agent sent.
    # Honoring an explicit ``labels_in_first_row: false`` after extension
    # causes the header text to be treated as data — corrupting the chart.
    row_labels = payload.get("labels_in_first_column")
    col_labels = payload.get("labels_in_first_row")
    if row_labels is None:
        row_labels = False
    if col_labels is None:
        col_labels = False

    if header_was_extended:
        # The range now starts one row earlier (the header row).  Force
        # col_labels=True so LibreOffice treats that row as category labels,
        # not as a data point.
        col_labels = True
        _log("create_chart: header extension detected — forcing labels_in_first_row=True")

    # Always probe category_range shape. Agents routinely send
    # labels_in_first_column=false even when category_range is a single
    # column of text — which causes the text column to be plotted as a
    # (non-numeric) data series and the chart renders empty. Force the
    # correct hint whenever the data shape contradicts the agent.
    try:
        if category_range is not None:
            c_addr = category_range.getRangeAddress()
            is_single_column = c_addr.StartColumn == c_addr.EndColumn
            is_single_row = c_addr.StartRow == c_addr.EndRow
            cat_has_text = False
            try:
                cat_data = list(map(list, category_range.getDataArray()))
                for cat_row in cat_data:
                    for v in cat_row:
                        if isinstance(v, str) and v.strip():
                            cat_has_text = True
                            break
                    if cat_has_text:
                        break
            except Exception:
                pass
            if is_single_column and cat_has_text and not row_labels:
                row_labels = True
                _log(
                    "create_chart: category_range is a single column of text — "
                    "forcing labels_in_first_column=True"
                )
            if is_single_row and cat_has_text and not col_labels:
                col_labels = True
                _log(
                    "create_chart: category_range is a single row of text — "
                    "forcing labels_in_first_row=True"
                )
    except Exception:
        pass

    try:
        if not header_was_extended and (
            payload.get("labels_in_first_row") is None
            or payload.get("labels_in_first_column") is None
        ):
            if combined_range is not None:
                rl, cl = _detect_chart_labels_from_range(combined_range)
                row_labels = row_labels or rl
                col_labels = col_labels or cl
            else:
                try:
                    if data_range is not None:
                        rl, _ = _detect_chart_labels_from_range(data_range)
                        row_labels = row_labels or rl
                except Exception:
                    pass
    except Exception:
        row_labels = bool(row_labels)
        col_labels = bool(col_labels)

    added = False
    try:
        charts.addNewByName(
            chart_name, chart_rect, range_addresses, row_labels, col_labels
        )
        added = True
    except Exception as exc:
        _log(f"create_chart primary add failed: {exc}; retrying without label hints")
        try:
            charts.addNewByName(chart_name, chart_rect, range_addresses, False, False)
            added = True
        except Exception as exc_retry:
            raise RuntimeError(f"create_chart failed to add chart: {exc_retry}")

    try:
        chart_obj = charts.getByName(chart_name)
        chart_document = chart_obj.getEmbeddedObject()
    except Exception:
        chart_document = None

    if chart_document is not None:
        # Build a proper range representation and set it immediately on the chart data.
        # Use a single combined range if available, otherwise join data and category ranges
        # with ';' so ChartData receives both ranges (Calc accepts joined ranges).
        try:
            chart_data = chart_document.getData()
            parts = []
            if combined_range is not None:
                parts.append(_range_to_representation(combined_range))
            else:
                # fallback: include data + category ranges (if resolved)
                try:
                    if data_range is not None:
                        parts.append(_range_to_representation(data_range))
                except Exception:
                    pass
                try:
                    if category_range is not None and category_range is not data_range:
                        parts.append(_range_to_representation(category_range))
                except Exception:
                    pass
            if parts:
                range_repr = ";".join(parts)
                try:
                    chart_data.setRangeRepresentation(range_repr)
                except Exception:
                    # some older engines may require setRangeAddresses; keep silent on failure.
                    pass
        except Exception:
            pass

        # Now request the diagram change. Creating the diagram from the chart's factory
        # gives the chart a diagram object compatible with the chart document.
        chart_type = _normalize_chart_type(payload.get("chart_type") or "")
        _apply_chart_type(chart_document, chart_type, payload)

        # --- Verify the chart type actually applied (pie/doughnut are prone to
        # failing silently and rendering as the default bar/column). ---
        if chart_type in ("pie", "pie3d", "doughnut"):
            _verify_chart_type_applied(chart_document, chart_type)

        _style_chart_document(chart_document, chart_type, chart_title, payload)

    cursor_ref = _resolve_range(
        document, action.get("cursor_target")
    ) or _resolve_range(document, action.get("target"))
    if cursor_ref is None and chart_rect is not None:
        # No explicit placement hint — the chart was auto-laid-out. Resolve the
        # actual anchor cell from the chart rectangle so the result metadata
        # reports where the chart really landed, not a hardcoded A1.
        cursor_ref = _cell_ref_from_chart_rect(dest_sheet, chart_rect)
    if cursor_ref is None:
        cursor_ref = dest_sheet.getCellRangeByName("A1")
    chart_area_abs = None
    chart_area_a1 = None
    if chart_area_range is not None:
        try:
            chart_area_abs = _range_to_representation(chart_area_range)
        except Exception:
            chart_area_abs = None
        chart_area_a1 = _range_to_a1_relative(chart_area_range)
    meta = action.setdefault("_result_meta", {})
    # Report actual chart rectangle so the agent knows where it landed
    actual_rect_info = None
    if chart_rect is not None:
        actual_rect_info = {
            "x": getattr(chart_rect, "X", 0),
            "y": getattr(chart_rect, "Y", 0),
            "width": getattr(chart_rect, "Width", 0),
            "height": getattr(chart_rect, "Height", 0),
        }
    meta.update(
        {
            "chart_name": chart_name,
            "chart_type": _normalize_chart_type(payload.get("chart_type") or ""),
            "destination_sheet": destination_sheet_name,
            "placement": native_placement,
            "derived_occupied_range": chart_area_abs,
            "derived_occupied_range_a1": chart_area_a1,
            "applied_table_range": chart_area_abs,
            "applied_table_range_a1": chart_area_a1,
            "actual_position": actual_rect_info,
        }
    )
    return chart_area_range or cursor_ref


def _chart_shape_for_name(sheet, chart_name: Optional[str]):
    if sheet is None or not chart_name:
        return None
    try:
        draw_page = sheet.getDrawPage()
        count = draw_page.getCount()
    except Exception:
        return None
    for idx in range(count):
        try:
            shape = draw_page.getByIndex(idx)
        except Exception:
            continue
        for attr in ("PersistName", "Name"):
            try:
                value = getattr(shape, attr, None)
                if callable(value):
                    value = value()
                if value is not None and str(value) == str(chart_name):
                    return shape
            except Exception:
                continue
    return None


def _chart_rect_for_object(chart_obj, sheet=None):
    if chart_obj is None:
        return None

    def _rect_from_position_size(obj):
        try:
            position_getter = getattr(obj, "getPosition", None)
            size_getter = getattr(obj, "getSize", None)
            position = (
                position_getter()
                if callable(position_getter)
                else getattr(obj, "Position", None)
            )
            size = size_getter() if callable(size_getter) else getattr(obj, "Size", None)
            if position is not None and size is not None:
                rectangle = uno.createUnoStruct("com.sun.star.awt.Rectangle")
                rectangle.X = int(getattr(position, "X", 0) or 0)
                rectangle.Y = int(getattr(position, "Y", 0) or 0)
                rectangle.Width = int(getattr(size, "Width", 0) or 0)
                rectangle.Height = int(getattr(size, "Height", 0) or 0)
                if rectangle.Width > 0 and rectangle.Height > 0:
                    return rectangle
        except Exception:
            pass
        return None

    direct_rect = _rect_from_position_size(chart_obj)
    if direct_rect is not None:
        return direct_rect

    if sheet is not None:
        chart_name = None
        for attr in ("getName", "Name"):
            try:
                value = getattr(chart_obj, attr, None)
                chart_name = value() if callable(value) else value
                if chart_name:
                    break
            except Exception:
                chart_name = None
        shape = _chart_shape_for_name(sheet, str(chart_name) if chart_name else None)
        shape_rect = _rect_from_position_size(shape)
        if shape_rect is not None:
            return shape_rect
    return None


def _chart_area_range_for_object(sheet, chart_obj):
    rect = _chart_rect_for_object(chart_obj, sheet=sheet)
    if rect is not None:
        chart_range = _cell_range_from_chart_rect(sheet, rect)
        if chart_range is not None:
            return chart_range
    return None


def _chart_title_for_object(chart_obj) -> Optional[str]:
    if chart_obj is None:
        return None
    chart_document = None
    try:
        chart_document = chart_obj.getEmbeddedObject()
    except Exception:
        chart_document = None
    if chart_document is None:
        return None
    try:
        title_obj = chart_document.getTitle()
    except Exception:
        title_obj = None
    if title_obj is not None:
        try:
            title_text = getattr(title_obj, "String", None)
            if title_text:
                return str(title_text)
        except Exception:
            pass
        try:
            getter = getattr(title_obj, "getString", None)
            title_text = getter() if callable(getter) else None
            if title_text:
                return str(title_text)
        except Exception:
            pass
    try:
        if _has_prop(chart_document, "Title"):
            title_obj = chart_document.getPropertyValue("Title")
            title_text = getattr(title_obj, "String", None)
            if title_text:
                return str(title_text)
    except Exception:
        pass
    return None


def _range_addresses_overlap(address_a, address_b) -> bool:
    if address_a is None or address_b is None:
        return False
    try:
        return not (
            int(address_a.EndColumn) < int(address_b.StartColumn)
            or int(address_b.EndColumn) < int(address_a.StartColumn)
            or int(address_a.EndRow) < int(address_b.StartRow)
            or int(address_b.EndRow) < int(address_a.StartRow)
        )
    except Exception:
        return False


def _chart_identity_matches(candidate: Optional[str], requested: Optional[str]) -> bool:
    if not candidate or not requested:
        return False
    return str(candidate).strip() == str(requested).strip()


def _rect_debug(rect) -> Optional[Dict[str, int]]:
    if rect is None:
        return None
    try:
        return {
            "x": int(getattr(rect, "X", 0) or 0),
            "y": int(getattr(rect, "Y", 0) or 0),
            "width": int(getattr(rect, "Width", 0) or 0),
            "height": int(getattr(rect, "Height", 0) or 0),
        }
    except Exception:
        return None


def _handle_delete_chart(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    chart_name = payload.get("chart_name")
    if chart_name is not None:
        chart_name = str(chart_name).strip()
    native_placement = payload.get("placement")
    sheet_name = (
        payload.get("sheet_name")
        or target_info.get("sheet_name")
        or payload.get("destination_sheet")
    )

    if not chart_name and not isinstance(native_placement, dict):
        raise RuntimeError("delete_chart requires chart_name or placement")

    resolved_target_range = None
    target_rect = None
    if isinstance(native_placement, dict):
        placement_sheet = _get_sheet(document, sheet_name) if sheet_name else None
        if placement_sheet is None and target_info.get("sheet_name"):
            placement_sheet = _get_sheet(document, target_info.get("sheet_name"))
        try:
            target_rect, placement_range = _chart_rect_and_area_from_placement(
                document,
                placement_sheet,
                native_placement,
            )
            if placement_range is not None:
                resolved_target_range = placement_range
                try:
                    sheet_name = _safe_sheet_name(placement_range.getSpreadsheet()) or sheet_name
                except Exception:
                    pass
        except ActionApplicationError:
            raise
        except Exception:
            target_rect = None

    sheets = _get_sheet_container(document)
    if sheets is None:
        raise RuntimeError("delete_chart unable to access document sheets collection")

    candidate_sheets = []
    if sheet_name:
        sheet = _get_sheet(document, str(sheet_name))
        if sheet is None:
            raise RuntimeError(f"delete_chart sheet '{sheet_name}' not found")
        candidate_sheets.append(sheet)
    else:
        try:
            for idx in range(sheets.getCount()):
                candidate_sheets.append(sheets.getByIndex(idx))
        except Exception:
            pass

    matches = []
    fallback_title_matches = []
    fallback_range_matches = []
    diagnostics = []
    target_address = None
    if resolved_target_range is not None:
        try:
            target_address = resolved_target_range.getRangeAddress()
        except Exception:
            target_address = None

    for sheet in candidate_sheets:
        target_address_for_sheet = target_address
        try:
            charts = sheet.getCharts()
            names = list(charts.getElementNames())
        except Exception:
            continue
        for name in names:
            try:
                chart_obj = charts.getByName(name)
            except Exception:
                continue
            visible_title = _chart_title_for_object(chart_obj)
            name_matches = _chart_identity_matches(str(name), chart_name)
            title_matches = _chart_identity_matches(visible_title, chart_name)
            identity_matches = bool(name_matches or title_matches)
            chart_area_range = _chart_area_range_for_object(sheet, chart_obj)
            chart_rect = _chart_rect_for_object(chart_obj, sheet=sheet)
            range_matches = False
            if target_address_for_sheet is not None and chart_area_range is not None:
                try:
                    range_matches = _range_addresses_overlap(
                        chart_area_range.getRangeAddress(),
                        target_address_for_sheet,
                    )
                except Exception:
                    range_matches = False
            if not range_matches and target_rect is not None and chart_rect is not None:
                range_matches = _rectangles_overlap(chart_rect, target_rect)
            if chart_name and target_rect is not None:
                if identity_matches and range_matches:
                    matches.append((sheet, charts, name, chart_area_range))
                elif identity_matches:
                    fallback_title_matches.append((sheet, charts, name, chart_area_range))
                elif range_matches:
                    fallback_range_matches.append((sheet, charts, name, chart_area_range))
            elif chart_name:
                if identity_matches:
                    matches.append((sheet, charts, name, chart_area_range))
            elif range_matches:
                matches.append((sheet, charts, name, chart_area_range))
            try:
                diagnostics.append(
                    {
                        "sheet": _safe_sheet_name(sheet),
                        "internal_name": str(name),
                        "visible_title": visible_title,
                        "name_matches": bool(name_matches),
                        "title_matches": bool(title_matches),
                        "range_matches": bool(range_matches),
                        "chart_rect": _rect_debug(chart_rect),
                        "area": _range_to_a1_relative(chart_area_range)
                        if chart_area_range is not None
                        else None,
                    }
                )
            except Exception:
                pass

    if not matches:
        if len(fallback_title_matches) == 1:
            matches = fallback_title_matches
            _log(
                "_handle_delete_chart: using unique visible-title/internal-name "
                f"fallback for chart_name={chart_name!r}; placement did not match exactly"
            )
        elif len(fallback_range_matches) == 1:
            matches = fallback_range_matches
            _log(
                "_handle_delete_chart: using unique placement fallback "
                f"for chart_name={chart_name!r}; chart identity did not match internal names"
            )
    if not matches:
        try:
            _log(
                "_handle_delete_chart: no match "
                f"chart_name={chart_name!r} placement={native_placement} "
                f"target_rect={_rect_debug(target_rect)} "
                f"target_range={_range_to_a1_relative(resolved_target_range) if resolved_target_range is not None else None} "
                f"candidates={json.dumps(diagnostics, default=str)[:4000]}"
            )
        except Exception:
            pass
        raise RuntimeError(
            f"delete_chart could not locate chart_name='{chart_name}' placement={native_placement}"
        )
    if len(matches) > 1:
        names = [match[2] for match in matches]
        raise RuntimeError(
            "delete_chart target matched multiple charts; provide a more specific "
            f"chart_name or placement. matches={names}"
        )

    sheet, charts, matched_name, chart_area_range = matches[0]
    try:
        matched_title = _chart_title_for_object(charts.getByName(matched_name))
    except Exception:
        matched_title = None
    deleted_area_abs = None
    deleted_area_a1 = None
    if chart_area_range is not None:
        try:
            deleted_area_abs = _range_to_representation(chart_area_range)
        except Exception:
            deleted_area_abs = None
        deleted_area_a1 = _range_to_a1_relative(chart_area_range)

    charts.removeByName(matched_name)
    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "deleted_chart_name": matched_name,
            "chart_name": matched_name,
            "requested_chart_name": chart_name,
            "deleted_chart_title": matched_title,
            "placement": native_placement if isinstance(native_placement, dict) else None,
            "deleted_occupied_range": deleted_area_abs,
            "deleted_occupied_range_a1": deleted_area_a1,
            "applied_table_range": deleted_area_abs,
            "applied_table_range_a1": deleted_area_a1,
        }
    )
    _log(
        "_handle_delete_chart: deleted chart "
        f"chart_name={matched_name} title={matched_title!r} placement={native_placement}"
    )
    if chart_area_range is not None:
        return chart_area_range
    if resolved_target_range is not None:
        return resolved_target_range
    return sheet.getCellRangeByName("A1")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _get_sheet_container(document):
    candidate = document
    visited = set()
    while candidate is not None and id(candidate) not in visited:
        visited.add(id(candidate))
        getter = getattr(candidate, "getSheets", None)
        if callable(getter):
            try:
                sheets = getter()
                if sheets is not None:
                    return sheets
            except Exception:
                pass
        sheets_attr = getattr(candidate, "Sheets", None)
        if sheets_attr is not None:
            return sheets_attr
        candidate = getattr(candidate, "Model", None)
    return None


_SHEET_SPACE_REGEX = re.compile(r"\s+")
_RANGE_SHEET_PREFIX_REGEX = re.compile(r"^(?P<sheet>'[^']+'|[^!.]+)[!.](?P<range>.+)$")


def _normalize_sheet_name(name: str) -> str:
    if not isinstance(name, str):
        name = str(name)
    cleaned = name.replace("_", " ").replace("-", " ")
    cleaned = cleaned.strip().strip("'\"")
    cleaned = _SHEET_SPACE_REGEX.sub(" ", cleaned)
    return cleaned.lower()


def _split_sheet_qualified_range_name(
    range_name: Optional[str],
) -> Tuple[Optional[str], Optional[str]]:
    if not range_name:
        return None, range_name
    cleaned = str(range_name).strip()
    match = _RANGE_SHEET_PREFIX_REGEX.match(cleaned)
    if not match:
        return None, cleaned.replace("$", "") if cleaned else cleaned
    sheet = match.group("sheet").strip().strip("'\"")
    range_only = match.group("range").strip()
    if range_only:
        range_only = range_only.replace("$", "")
    if not range_only:
        return None, range_name
    return sheet or None, range_only


def _resolve_sheet_and_range_name(
    document,
    sheet_name: Optional[str],
    range_name: Optional[str],
    context: str,
) -> Tuple[Optional[Any], Optional[str]]:
    range_sheet, range_only = _split_sheet_qualified_range_name(range_name)
    if range_sheet:
        if sheet_name and str(sheet_name) != str(range_sheet):
            _log(
                f"{context}: range sheet '{range_sheet}' overrides provided sheet_name '{sheet_name}'"
            )
        sheet_name = range_sheet
        range_name = range_only
    sheet = _get_sheet(document, sheet_name)
    return sheet, range_name


def _get_sheet(document, sheet_name: Optional[str]):
    if not sheet_name:
        return None
    sheets = _get_sheet_container(document)
    if sheets is None:
        return None
    # Fast path: exact match
    try:
        if sheets.hasByName(sheet_name):
            return sheets.getByName(sheet_name)
    except Exception:
        pass

    try:
        names = list(sheets.getElementNames())
    except Exception:
        names = []

    target_norm = _normalize_sheet_name(str(sheet_name))
    for name in names:
        try:
            candidate_norm = _normalize_sheet_name(str(name))
        except Exception:
            continue
        if candidate_norm == target_norm:
            try:
                _log(
                    f"_get_sheet: matched '{sheet_name}' to workbook sheet '{name}' via normalized lookup"
                )
                return sheets.getByName(name)
            except Exception:
                continue

    _log(f"_get_sheet: sheet '{sheet_name}' not found; available={names}")
    return None


def _resolve_range(document, ref: Optional[Dict[str, Any]]):
    if not ref:
        return None
    sheet, range_name = _resolve_sheet_and_range_name(
        document, ref.get("sheet_name"), ref.get("range"), "_resolve_range"
    )
    if sheet is None:
        return None

    if range_name:
        try:
            return sheet.getCellRangeByName(range_name)
        except Exception:
            return None

    start_row = ref.get("start_row")
    start_col = ref.get("start_column")
    end_row = ref.get("end_row") or start_row
    end_col = ref.get("end_column") or start_col
    if start_row is None or start_col is None:
        return None
    try:
        return sheet.getCellRangeByPosition(
            int(start_col) - 1,
            int(start_row) - 1,
            int(end_col) - 1,
            int(end_row) - 1,
        )
    except Exception:
        return None


def _resolve_postprocess_sheet_name(
    document, controller, action: Dict[str, Any], target_range=None
) -> Optional[str]:
    target = action.get("target") or {}
    payload = action.get("payload") or {}
    action_kind = str(action.get("kind") or "").strip().lower()

    if action_kind == "rename_sheet":
        renamed = payload.get("new_sheet_name")
        if isinstance(renamed, str) and renamed.strip():
            return renamed.strip()

    for raw_name in (
        target.get("sheet_name"),
        payload.get("sheet_name"),
        action.get("sheet_name"),
    ):
        if isinstance(raw_name, str):
            cleaned = raw_name.strip()
            if cleaned:
                return cleaned

    candidate_ranges = (
        target.get("range"),
        payload.get("range"),
        payload.get("target_range"),
        payload.get("start_cell"),
        payload.get("target_cell"),
        payload.get("freeze_at"),
        payload.get("anchor_cell"),
    )
    for candidate in candidate_ranges:
        if not isinstance(candidate, str):
            continue
        range_sheet, _ = _split_sheet_qualified_range_name(candidate)
        if range_sheet:
            return range_sheet

    if target_range is not None:
        getter = getattr(target_range, "getSpreadsheet", None)
        if callable(getter):
            try:
                sheet = getter()
                if sheet is not None:
                    get_name = getattr(sheet, "getName", None)
                    if callable(get_name):
                        name = get_name()
                        if isinstance(name, str) and name:
                            return name
            except Exception:
                pass
        try:
            addr = target_range.getRangeAddress()
            sheets = _get_sheet_container(document)
            if sheets is not None and hasattr(sheets, "getByIndex"):
                sheet = sheets.getByIndex(addr.Sheet)
                get_name = getattr(sheet, "getName", None)
                if callable(get_name):
                    name = get_name()
                    if isinstance(name, str) and name:
                        return name
        except Exception:
            pass

    try:
        active_sheet = controller.getActiveSheet() if controller else None
        get_name = getattr(active_sheet, "getName", None)
        if callable(get_name):
            name = get_name()
            if isinstance(name, str) and name:
                return name
    except Exception:
        pass

    return None


def _resolve_postprocess_sheet_targets(
    document, sheet_names, include_all_when_empty: bool = False
) -> List[str]:
    requested_names: List[str] = []
    requested_set: set[str] = set()
    for raw_name in sheet_names or []:
        if not isinstance(raw_name, str):
            continue
        name = raw_name.strip()
        if not name or name in requested_set:
            continue
        requested_set.add(name)
        requested_names.append(name)

    if not requested_names and not include_all_when_empty:
        return []

    names: List[str] = []
    seen: set[str] = set()
    sheets = _get_sheet_container(document)
    if sheets is None:
        return requested_names

    try:
        count = sheets.getCount()
    except Exception:
        return requested_names

    for idx in range(count):
        try:
            sheet = sheets.getByIndex(idx)
            get_name = getattr(sheet, "getName", None)
            name = get_name() if callable(get_name) else None
        except Exception:
            continue
        if not isinstance(name, str) or not name or name in seen:
            continue
        if not include_all_when_empty and name not in requested_set:
            continue
        seen.add(name)
        names.append(name)

    if names:
        return names
    return requested_names


def _record_postprocess_range(
    ranges_by_sheet: Dict[str, List[tuple[int, int, int, int]]],
    sheet_name: Optional[str],
    target_range,
) -> None:
    if not sheet_name or target_range is None:
        return
    try:
        address = target_range.getRangeAddress()
        bounds = (
            int(address.StartColumn),
            int(address.StartRow),
            int(address.EndColumn),
            int(address.EndRow),
        )
    except Exception:
        return
    ranges = ranges_by_sheet.setdefault(sheet_name, [])
    if bounds not in ranges:
        ranges.append(bounds)


def _range_to_a1_relative(cell_range) -> Optional[str]:
    if cell_range is None:
        return None
    try:
        address = cell_range.getRangeAddress()
    except Exception:
        return None
    col_start = _column_index_to_name(address.StartColumn)
    col_end = _column_index_to_name(address.EndColumn)
    row_start = int(address.StartRow) + 1
    row_end = int(address.EndRow) + 1
    start_ref = f"{col_start}{row_start}"
    end_ref = f"{col_end}{row_end}"
    if start_ref == end_ref:
        return start_ref
    return f"{start_ref}:{end_ref}"


def _normalize_target_ref(sheet_name: Optional[str], range_name: Optional[str]) -> Optional[str]:
    if not range_name:
        return None
    sheet_from_range, local_range = _split_sheet_qualified_range_name(str(range_name))
    chosen_sheet = sheet_from_range or (str(sheet_name).strip() if sheet_name else "")
    if local_range:
        local_range = local_range.replace("$", "").strip()
    if not local_range:
        return None
    if chosen_sheet:
        return f"{chosen_sheet}!{local_range}"
    return local_range


def _extract_requested_target_ref(action: Dict[str, Any]) -> Optional[str]:
    target = action.get("target") or {}
    payload = action.get("payload") or {}
    sheet_name = target.get("sheet_name") or payload.get("sheet_name")
    candidates = [
        target.get("range"),
        payload.get("range"),
        payload.get("target_range"),
        payload.get("start_cell"),
        payload.get("target_cell"),
        payload.get("anchor_cell"),
        payload.get("freeze_at"),
        payload.get("data_target_range"),
        payload.get("category_target_range"),
    ]
    for raw in candidates:
        if not isinstance(raw, str):
            continue
        cleaned = raw.strip()
        if not cleaned:
            continue
        normalized = _normalize_target_ref(sheet_name, cleaned)
        if normalized:
            return normalized
    return None


def _describe_applied_target(document, controller, action: Dict[str, Any], target_range) -> Dict[str, Any]:
    details: Dict[str, Any] = {}
    requested_target = _extract_requested_target_ref(action)
    if requested_target:
        details["requested_target"] = requested_target

    action_meta = action.get("_result_meta")
    meta_applied_abs = None
    meta_applied_a1 = None
    if isinstance(action_meta, dict):
        meta_applied_abs = action_meta.get("applied_table_range")
        meta_applied_a1 = action_meta.get("applied_table_range_a1")

    if target_range is None:
        return details

    applied_abs = None
    try:
        applied_abs = _range_to_representation(target_range)
    except Exception:
        applied_abs = None
    if isinstance(meta_applied_abs, str) and meta_applied_abs.strip():
        details["applied_range"] = meta_applied_abs
    elif applied_abs:
        details["applied_range"] = applied_abs

    applied_local = _range_to_a1_relative(target_range)
    if isinstance(meta_applied_a1, str) and meta_applied_a1.strip():
        details["applied_range_a1"] = meta_applied_a1
    elif applied_local:
        details["applied_range_a1"] = applied_local

    sheet_name = _resolve_postprocess_sheet_name(document, controller, action, target_range)
    if sheet_name:
        details["applied_sheet_name"] = sheet_name

    bounds = _range_address_bounds(target_range)
    if isinstance(bounds, dict):
        details["applied_bounds"] = bounds
        if all(
            bounds.get(k) is not None
            for k in ("start_row", "end_row", "start_column", "end_column")
        ):
            details["applied_bounds_1based"] = {
                "start_row": int(bounds["start_row"]) + 1,
                "end_row": int(bounds["end_row"]) + 1,
                "start_column": int(bounds["start_column"]) + 1,
                "end_column": int(bounds["end_column"]) + 1,
            }

    compare_local = (
        meta_applied_a1
        if isinstance(meta_applied_a1, str) and meta_applied_a1.strip()
        else applied_local
    )
    if requested_target and compare_local:
        requested_range = requested_target.split("!", 1)[-1].replace("$", "").upper()
        applied_norm = compare_local.replace("$", "").upper()
        if requested_range != applied_norm:
            details["target_shifted"] = True

    return details


def _write_cell(cell, value: Any) -> None:
    if value is None or value == "":
        cell.setString("")
        return
    if isinstance(value, bool):
        cell.setValue(1 if value else 0)
        return
    if isinstance(value, (int, float)):
        cell.setValue(float(value))
        return
    value_str = str(value)
    if value_str.startswith("="):
        try:
            sheet_name = None
            sheet_obj = getattr(cell, "getSpreadsheet", None)
            if callable(sheet_obj):
                spreadsheet = sheet_obj()
                if spreadsheet is not None:
                    get_name = getattr(spreadsheet, "getName", None)
                    if callable(get_name):
                        sheet_name = get_name()
            if sheet_name:
                document = XSCRIPTCONTEXT.getDocument()  # type: ignore[name-defined]
                _validate_formula_references(document, sheet_name, value_str)
        except Exception as validation_exc:
            raise RuntimeError(f"Formula reference check failed: {validation_exc}")
        _set_formula_with_locale(cell, value_str)
    else:
        cell.setString(value_str)


def _cell_matches_value(cell, value: Any) -> bool:
    if value is None or value == "":
        try:
            text = cell.getString()
            if text:
                return False
        except Exception:
            pass
        try:
            num = cell.getValue()
            if isinstance(num, (int, float)) and num not in (0, 0.0):
                return False
        except Exception:
            pass
        return True

    if isinstance(value, str):
        if value.startswith("="):
            try:
                return cell.getFormula() == value
            except Exception:
                return False
        try:
            return cell.getString() == value
        except Exception:
            return False

    if isinstance(value, bool):
        try:
            existing = cell.getValue()
            return bool(existing) == value
        except Exception:
            return False

    if isinstance(value, (int, float)):
        try:
            existing = cell.getValue()
            if existing is None:
                return False
            return abs(float(existing) - float(value)) < 1e-7
        except Exception:
            return False

    try:
        return cell.getString() == str(value)
    except Exception:
        return False


def _write_cell_with_validation(cell, value: Any) -> None:
    _write_cell(cell, value)
    if _cell_matches_value(cell, value):
        return
    try:
        if isinstance(value, str) and value.startswith("="):
            _set_formula_with_locale(cell, value)
        elif isinstance(value, (int, float)):
            cell.setValue(float(value))
        elif isinstance(value, bool):
            cell.setValue(1 if value else 0)
        elif value is None or value == "":
            cell.setString("")
        else:
            cell.setString(str(value))
    except Exception:
        try:
            cell.setString("" if value is None else str(value))
        except Exception:
            _log(f"_write_cell_with_validation: failed to coerce value {value!r}")


_FORMULA_REF_PATTERN = re.compile(
    r"(?:(?P<sheet>'[^']+'|[A-Za-z0-9_]+)\s*[.!])?"
    r"(?P<cell>\$?[A-Za-z]{1,3}\$?\d+)"
    r"(?::(?:(?P<sheet2>'[^']+'|[A-Za-z0-9_]+)\s*[.!])?"
    r"(?P<cell2>\$?[A-Za-z]{1,3}\$?\d+))?"
)


def _normalize_sheet_token(token: Optional[str]) -> Optional[str]:
    if not token:
        return None
    cleaned = token.strip()
    if cleaned.startswith("'") and cleaned.endswith("'") and len(cleaned) > 1:
        cleaned = cleaned[1:-1]
    return cleaned or None


def _normalize_cell_ref(cell_ref: str) -> str:
    return cell_ref.replace("$", "")


def _strip_formula_literals(formula: str) -> str:
    return re.sub(r'"([^"]|"")*"', "", formula)


def _extract_formula_references(
    formula: str, default_sheet: Optional[str]
) -> List[Dict[str, Any]]:
    if not isinstance(formula, str) or not formula.startswith("="):
        return []
    text = _strip_formula_literals(formula)
    refs: List[Dict[str, Any]] = []
    for match in _FORMULA_REF_PATTERN.finditer(text):
        sheet = _normalize_sheet_token(match.group("sheet")) or default_sheet
        cell = match.group("cell")
        sheet2 = _normalize_sheet_token(match.group("sheet2")) or sheet
        cell2 = match.group("cell2")
        if cell:
            refs.append(
                {
                    "sheet": sheet,
                    "cell": _normalize_cell_ref(cell),
                    "is_range": cell2 is not None,
                }
            )
        if cell2:
            refs.append(
                {
                    "sheet": sheet2,
                    "cell": _normalize_cell_ref(cell2),
                    "is_range": True,
                }
            )
    return refs


_KNOWN_FORMULA_ERRORS = {
    "#DIV/0!",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#REF!",
    "#VALUE!",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
    "#FIELD!",
    "#CONNECT!",
    "#BLOCKED!",
    "#UNKNOWN!",
    "ERR:508",
    "ERR:509",
    "ERR:510",
    "ERR:511",
    "ERR:512",
    "ERR:514",
    "ERR:516",
    "ERR:517",
    "ERR:519",
    "ERR:522",
    "ERR:525",
    "ERR:532",
    "ERR:533",
}
_ERR_CODE_REGEX = re.compile(r"^ERR:\d+$")


def _normalize_formula_error_token(raw_error: Any) -> Optional[str]:
    if raw_error is None:
        return None
    token = str(raw_error).strip().upper()
    if not token:
        return None
    if token in _KNOWN_FORMULA_ERRORS:
        return token
    if _ERR_CODE_REGEX.match(token):
        return token
    return None


def _scan_formula_errors(document, sheet_names: Optional[Iterable[str]] = None) -> Dict[str, Any]:
    items: List[Dict[str, Any]] = []
    by_type: Dict[str, int] = {}

    sheets = _get_sheet_container(document)
    if sheets is None:
        return {"total": 0, "by_type": {}, "items": [], "sample": []}

    requested_names: Optional[set[str]] = None
    if sheet_names is not None:
        requested_names = set()
        for raw_name in sheet_names:
            if not isinstance(raw_name, str):
                continue
            cleaned = raw_name.strip()
            if cleaned:
                requested_names.add(cleaned)

    try:
        workbook_sheet_names = list(sheets.getElementNames())
    except Exception:
        workbook_sheet_names = []

    effective_sheet_names = []
    for sheet_name in workbook_sheet_names:
        if requested_names is not None and sheet_name not in requested_names:
            continue
        effective_sheet_names.append(sheet_name)

    for sheet_name in effective_sheet_names:
        sheet = _get_sheet(document, sheet_name)
        if sheet is None:
            continue
        used = _sheet_used_area(sheet)
        if used is None:
            continue

        start_row = int(getattr(used, "StartRow", 0))
        end_row = int(getattr(used, "EndRow", -1))
        start_col = int(getattr(used, "StartColumn", 0))
        end_col = int(getattr(used, "EndColumn", -1))
        if end_row < start_row or end_col < start_col:
            continue

        # Block fetch: two UNO bridge calls per sheet instead of three per
        # cell. ``getFormulaArray`` returns the formula text for formula cells
        # (starts with "="); empty string for non-formula cells. ``getDataArray``
        # returns the displayed/computed value, which is the error token string
        # (e.g. "#VALUE!", "#REF!") for cells that evaluated to an error.
        try:
            block = sheet.getCellRangeByPosition(
                start_col, start_row, end_col, end_row
            )
            formula_grid = block.getFormulaArray()
            data_grid = block.getDataArray()
        except Exception as exc:
            _log(
                f"_scan_formula_errors: block read failed for '{sheet_name}': {exc}"
            )
            continue

        for r, formula_row in enumerate(formula_grid):
            data_row = data_grid[r] if r < len(data_grid) else ()
            for c, formula_text in enumerate(formula_row):
                if (
                    not isinstance(formula_text, str)
                    or not formula_text.startswith("=")
                ):
                    continue

                # Formula cell — inspect for error. Prefer getError() when it
                # surfaces a non-zero code (matches previous behavior); fall
                # back to the displayed value from the data grid (no extra UNO
                # call) which already carries error tokens like "#VALUE!".
                abs_col = start_col + c
                abs_row = start_row + r
                err_token: Optional[str] = None
                try:
                    cell = sheet.getCellByPosition(abs_col, abs_row)
                    get_error = getattr(cell, "getError", None)
                    if callable(get_error):
                        error_code = int(get_error())
                        if error_code:
                            err_token = _normalize_formula_error_token(
                                f"ERR:{error_code}"
                            )
                except Exception:
                    err_token = None

                if err_token is None:
                    display_value = data_row[c] if c < len(data_row) else None
                    err_token = _normalize_formula_error_token(display_value)

                if err_token is None:
                    continue

                cell_name = f"{_column_index_to_name(abs_col)}{abs_row + 1}"
                entry = {
                    "sheet_name": sheet_name,
                    "cell": cell_name,
                    "error": err_token,
                }
                items.append(entry)
                by_type[err_token] = by_type.get(err_token, 0) + 1

    sample: List[Dict[str, Any]] = []
    seen_types: set = set()
    for entry in items:
        err = entry.get("error") or "ERROR"
        if err not in seen_types:
            sample.append(entry)
            seen_types.add(err)
    if len(sample) < 8:
        for entry in items:
            if len(sample) >= 8:
                break
            if entry not in sample:
                sample.append(entry)

    return {
        "total": len(items),
        "by_type": by_type,
        "items": items,
        "sample": sample[:8],
    }


def _cell_is_effectively_empty(cell) -> bool:
    try:
        formula = cell.getFormula()
        if isinstance(formula, str) and formula.strip():
            return False
    except Exception:
        pass
    try:
        text = cell.getString()
        if isinstance(text, str) and text.strip():
            return False
    except Exception:
        pass
    return True


def _validate_formula_references(
    document, current_sheet_name: str, formula: str, allow_range_blanks: bool = True
) -> None:
    refs = _extract_formula_references(formula, current_sheet_name)
    if not refs:
        return

    for ref in refs:
        ref_sheet_name = ref.get("sheet") or current_sheet_name
        ref_cell_name = ref.get("cell")
        if not ref_sheet_name or not ref_cell_name:
            continue

        ref_sheet = _get_sheet(document, ref_sheet_name)
        if ref_sheet is None:
            raise RuntimeError(f"missing sheet '{ref_sheet_name}'")

        if allow_range_blanks and ref.get("is_range"):
            continue

        try:
            ref_cell = ref_sheet.getCellRangeByName(ref_cell_name)
        except Exception:
            raise RuntimeError(
                f"invalid reference '{ref_sheet_name}!{ref_cell_name}'"
            )

        if _cell_is_effectively_empty(ref_cell):
            if _STRICT_FORMULA_REFERENCE_CHECK:
                raise RuntimeError(
                    f"formula '{formula}' references empty cell '{ref_sheet_name}!{ref_cell_name}'"
                )
            _log(
                "ApplyExcelActionPlan: warning - formula references empty cell "
                f"'{ref_sheet_name}!{ref_cell_name}'"
            )


def _get_sheet_names(document) -> List[str]:
    sheets = _get_sheet_container(document)
    if sheets is None:
        return []
    try:
        return list(sheets.getElementNames())
    except Exception:
        return []


def _normalize_formula_sheet_names(document, formula: str) -> str:
    """
    Rewrite formula to use correct case for sheet references.
    Example: '=assumptions!B3' -> '=Assumptions!B3'
    """
    if not formula or "!" not in formula:
        return formula
    
    if not document:
        return formula

    sheet_names = _get_sheet_names(document)
    if not sheet_names:
        _log("ApplyExcelActionPlan: no sheet names found for normalization")
        return formula

    # Mapping of normalized (canonical) name to actual name
    name_map = {}
    for name in sheet_names:
        name_map[_normalize_sheet_name(name)] = name

    def replace_sheet(match):
        quoted = match.group(1)
        unquoted = match.group(2)
        original_name = quoted or unquoted
        
        norm_name = _normalize_sheet_name(original_name)
        
        if norm_name in name_map:
            real_name = name_map[norm_name]
            _log(f"ApplyExcelActionPlan: normalizing sheet '{original_name}' -> '{real_name}'")
            # If name has spaces or was quoted, it MUST be quoted.
            if " " in real_name or quoted:
                return f"'{real_name}'!"
            return f"{real_name}!"
        
        return match.group(0)

    # Use a more specific regex to avoid greediness.
    # Group 1: Quoted sheet name (content between ' ')
    # Group 2: Unquoted sheet name (alphanumeric, dots, underscores)
    # Followed by !
    regex = r"(?:'([^']+)'|([a-zA-Z0-9_\.]+))!"
    result = re.sub(regex, replace_sheet, formula)
    if result != formula:
        _log(f"ApplyExcelActionPlan: formula normalized from '{formula}' to '{result}'")
    return result

def _set_formula_with_locale(cell, formula: str) -> None:
    original_exc: Optional[Exception] = None

    # NEW: Normalize sheet name case and separator to fix #NAME? and Err:509
    try:
        # 1. Normalize Case (e.g., assumptions! -> Assumptions!)
        document = XSCRIPTCONTEXT.getDocument() # type: ignore
        formula = _normalize_formula_sheet_names(document, formula)
        
        # 2. Translate Separator (! -> .)
        # This is CRITICAL for the LibreOffice UNO API's setFormula method.
        formula = _excel_formula_to_calc(formula)
    except Exception as norm_exc:
        _log(f"ApplyExcelActionPlan: formula pre-processing failed: {norm_exc}")

    _log(f"ApplyExcelActionPlan: setting cell formula: {formula}")
    try:
        cell.setFormula(formula)
        _mark_formula_dirty()
        return
    except Exception as exc:
        original_exc = exc

    # If first attempt failed, try once more with fallback properties if exposed.
    for attr_name in ("FormulaLocal", "Formula"):
        try:
            setattr(cell, attr_name, formula)
            _mark_formula_dirty()
            return
        except Exception:
            continue
    raise original_exc or Exception("Failed to set formula")


def _excel_formula_to_calc(formula: str) -> str:
    if not formula:
        return formula
    
    # 1. Handle commas (parameter separator)
    if "," in formula:
        chars: list[str] = []
        depth = 0
        for ch in formula:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth = max(0, depth - 1)
            if ch == "," and depth > 0:
                chars.append(";")
            else:
                chars.append(ch)
        formula = "".join(chars)
    
    # 2. Handle exclamation marks (sheet separator)
    # Calc's native UNO format uses '.' instead of '!' for sheet references.
    # We replace '!' with '.' only if it looks like a sheet reference pattern.
    if "!" in formula:
        original = formula
        # Replace ! with . if it's following a word or quoted string
        # Regex: ([a-zA-Z0-9_\.]+)! (unquoted) or '([^']+)'! (quoted)
        # We use a very specific pattern to avoid matching logic operators
        formula = re.sub(r"(?<![<>])([a-zA-Z0-9_\.]+)!", r"\1.", formula)
        formula = re.sub(r"'([^']+)'!", r"'\1'.", formula)
        
        if formula != original:
            _log(f"ApplyExcelActionPlan: formula separator translated: {original} -> {formula}")

    return formula


def _mark_formula_dirty() -> None:
    global _FORMULA_RECALC_PENDING
    _FORMULA_RECALC_PENDING = True


def _maybe_calculate_formulas(
    document, controller, force: bool = False
) -> None:
    """
    Recalculate formulas after a plan finishes.

    Uses ``XCalculatable.calculate()`` (dirty-only) instead of ``calculateAll()``.
    Calc auto-marks cells dirty as we write through ``setValue`` / ``setFormula``
    / ``setDataArray``, and propagates the dirty bit to dependents. So
    ``calculate()`` recomputes only the touched cells and their downstream
    dependents — the practical "only touched sheets" recalc that LibreOffice
    exposes (``XSpreadsheet`` itself has no recalc method).

    The previous ``calculateAll()`` re-evaluated every formula in every sheet
    on every plan, which scaled with workbook size and dominated runtime as
    sheets accumulated.
    """
    global _FORMULA_RECALC_PENDING
    if not _FORMULA_RECALC_PENDING and not force:
        return
    try:
        document.calculate()
        if _FORMULA_RECALC_PENDING:
            _log("ApplyExcelActionPlan: document.calculate() after formula writes")
        else:
            _log("ApplyExcelActionPlan: document.calculate() after write actions")
        try:
            ctx = XSCRIPTCONTEXT.getComponentContext()  # type: ignore[name-defined]
            sm = ctx.getServiceManager()
            dispatcher = sm.createInstance("com.sun.star.frame.DispatchHelper")
            frame = controller.getFrame()
            dispatcher.executeDispatch(frame, ".uno:Repaint", "", 0, ())
            _log("ApplyExcelActionPlan: dispatched .uno:Repaint")
        except Exception as repaint_exc:
            _log(f"ApplyExcelActionPlan: repaint dispatch failed: {repaint_exc}")
    except Exception as exc:
        _log(f"ApplyExcelActionPlan: calculate failed: {exc}")
    finally:
        _FORMULA_RECALC_PENDING = False


def _number_format_string(number_formats, fmt_id: Optional[int]) -> Optional[str]:
    if fmt_id in (None, -1):
        return None
    try:
        fmt_obj = number_formats.getByKey(int(fmt_id))
    except Exception:
        return None
    for attr in ("FormatString", "formatstring"):
        try:
            value = getattr(fmt_obj, attr, None)
            if isinstance(value, str):
                stripped = value.strip()
                if stripped:
                    return stripped
        except Exception:
            continue
    return None


def _quoted_currency_prefix(fmt_string: str) -> str:
    if not isinstance(fmt_string, str):
        return ""
    match = re.search(r'"([^"]+)"\s*\*', fmt_string)
    if match:
        return match.group(1)
    match = re.search(r'"([^"]+)"', fmt_string)
    if match:
        return match.group(1).strip()
    for symbol in "$€£¥₹₦₩₽₪฿₫":
        if symbol in fmt_string:
            return symbol
    return ""


def _decimal_places_from_format(fmt_string: str) -> int:
    if not isinstance(fmt_string, str):
        return 0
    positive_section = fmt_string.split(";")[0]
    match = re.search(r"\.([0#]+)", positive_section)
    if not match:
        return 0
    return min(4, len(match.group(1)))


def _estimate_numeric_display_chars(value: float, fmt_string: str) -> int:
    """
    Estimate rendered width for numeric cells. Calc's OptimalWidth can
    under-measure accounting formats because the symbol is left-filled via "*".
    """
    try:
        numeric = float(value)
    except Exception:
        return 0
    if not math.isfinite(numeric):
        return 0
    fmt_string = fmt_string or ""
    is_percent = "%" in fmt_string
    display_value = abs(numeric * 100 if is_percent else numeric)
    decimals = _decimal_places_from_format(fmt_string)
    if decimals > 0:
        rendered = f"{display_value:,.{decimals}f}"
        if "#" in (fmt_string.split(".")[1].split(";")[0] if "." in fmt_string else ""):
            rendered = rendered.rstrip("0").rstrip(".")
    else:
        rendered = f"{display_value:,.0f}"
    if is_percent:
        rendered += "%"
    extra = len(_quoted_currency_prefix(fmt_string))
    if numeric < 0:
        extra += 2
    if re.search(r'"[^"]+"\s*\*', fmt_string):
        extra += _ACCOUNTING_AUTOFIT_EXTRA_CHARS
    return max(8, min(52, len(rendered) + extra + 2))


def _is_general_number_format(number_formats, fmt_id: Optional[int]) -> bool:
    if fmt_id in (None, -1, 0):
        return True
    fmt_string = _number_format_string(number_formats, fmt_id)
    if not fmt_string:
        return False
    normalized = fmt_string.strip().lower()
    return normalized in ("general", "standard")


def _postprocess_numeric_candidate(
    cell_type: Any, raw_value: Any
) -> Optional[float]:
    """
    Only postprocess real Calc numeric cells.

    This safety net exists to recover when the plan omitted an explicit number
    format, but it must not coerce text cells during the save path because that
    can leak number-format styles onto unrelated styled cells in XLSX export.
    """
    if cell_type != CELL_CONTENT_VALUE:
        return None
    if isinstance(raw_value, bool) or not isinstance(raw_value, (int, float)):
        return None
    return float(raw_value)


def _maybe_set_horizontal_alignment(cell, desired_justify) -> None:
    """
    Best-effort horizontal alignment without clobbering explicit non-default choices
    such as center/justified/distributed.
    """
    try:
        current = cell.HoriJustify
    except Exception:
        return
    if current == desired_justify:
        return
    if current not in (HORI_STANDARD, HORI_LEFT, HORI_RIGHT):
        return
    try:
        cell.HoriJustify = desired_justify
    except Exception:
        pass


_GRAB_BAG_THEME_KEYS = frozenset({
    "CharThemeColor",
    "CharThemeOriginalColor",
    "CharThemeFinalColor",
    "CharColorTheme",
    "CharColor",
    "CharThemeTint",
    "CharThemeShade",
    "CharColorLumMod",
    "CharColorLumOff",
    "CharColorTintOrShade",
    "CharComplexColor",
    "ThemeColor",
    "ThemeTint",
    "ThemeShade",
    "ColorTheme",
    "OriginalColor",
})


def _build_rgb_complex_color(color_int: int):
    """Create a ``com.sun.star.util.ComplexColor`` UNO struct for explicit RGB.

    When this struct is assigned to ``CharComplexColor``, the XLSX exporter
    sees ``Type == RGB`` and serialises ``<color rgb="FF…"/>`` instead of
    ``<color theme="…"/>``.  If the struct type does not exist in this
    LibreOffice build the function returns ``None`` and the caller falls
    back to property-level clearing.
    """
    try:
        cc = uno.createUnoStruct("com.sun.star.util.ComplexColor")
        # Field names vary across LO versions – try both conventions.
        for type_field in ("Type", "meType"):
            if hasattr(cc, type_field):
                setattr(cc, type_field, 1)  # 1 == ComplexColorType.RGB
                break
        for color_field in ("FinalColor", "mnFinalColor"):
            if hasattr(cc, color_field):
                setattr(cc, color_field, int(color_int))
                break
        for scheme_field in ("SchemeIndex", "meSchemeType", "SchemeType"):
            if hasattr(cc, scheme_field):
                setattr(cc, scheme_field, -1)  # Unknown / no theme
                break
        for trans_field in ("Transformations", "maTransformations"):
            if hasattr(cc, trans_field):
                setattr(cc, trans_field, ())
                break
        return cc
    except Exception:
        return None


def _clear_char_color_theme_state(cell) -> None:
    """Remove every theme-colour trace so the XLSX exporter serialises plain RGB.

    LibreOffice preserves OOXML theme references in several independent places:
      1. Dedicated UNO properties (``CharColorTheme``, ``CharColorLumMod``, …)
      2. ``CharComplexColor`` – a ``ComplexColor`` struct with a theme type
      3. Interop grab bags (``CharInteropGrabBag`` on Writer cells,
         ``CellInteropGrabBag`` on Calc cells) – tuples of NamedValue /
         PropertyValue items the XLSX filter reads when writing ``<color/>``.

    All three layers must be neutralised; clearing only (1) is insufficient
    because the exporter checks (2) and (3) first.
    """
    if cell is None:
        return

    # --- 1. Clear dedicated theme properties --------------------------------
    neutral_props = (
        ("CharColorTheme", -1),
        ("CharColorTintOrShade", 0),
        ("CharColorLumMod", 10000),
        ("CharColorLumOff", 0),
    )
    for prop_name, prop_value in neutral_props:
        if not _has_prop(cell, prop_name):
            continue
        _set_uno_property(cell, prop_name, prop_value)

    # --- 2. Strip theme-colour entries from ALL interop grab bags -----------
    _strip_interop_grab_bag_theme(cell, "CharInteropGrabBag")   # Writer cells
    _strip_interop_grab_bag_theme(cell, "CellInteropGrabBag")   # Calc cells


def _strip_interop_grab_bag_theme(cell, bag_prop_name: str) -> None:
    """Remove theme-colour keys from an interop grab bag property.

    Works for both ``CharInteropGrabBag`` (Writer) and
    ``CellInteropGrabBag`` (Calc).  The bag is a tuple/sequence of
    ``PropertyValue`` / ``NamedValue`` objects.  We filter out entries whose
    ``.Name`` matches a known set of theme-colour keys, then write back the
    cleaned tuple.
    """
    if cell is None:
        return
    if not _has_prop(cell, bag_prop_name):
        return
    try:
        bag = cell.getPropertyValue(bag_prop_name)
    except Exception:
        try:
            bag = getattr(cell, bag_prop_name, None)
        except Exception:
            return
    if not bag:
        return

    filtered = tuple(
        entry for entry in bag
        if getattr(entry, "Name", None) not in _GRAB_BAG_THEME_KEYS
    )
    if len(filtered) == len(bag):
        return  # nothing removed – skip the write

    try:
        cell.setPropertyValue(bag_prop_name, filtered)
    except Exception:
        try:
            setattr(cell, bag_prop_name, filtered)
        except Exception:
            pass


def _set_char_complex_color_rgb(cell, color_int: int) -> None:
    """Force the character ComplexColor to an explicit RGB value when possible."""
    if cell is None or not _has_prop(cell, "CharComplexColor"):
        return
    complex_color = _build_rgb_complex_color(int(color_int))
    if complex_color is None:
        return
    try:
        cell.setPropertyValue("CharComplexColor", complex_color)
    except Exception:
        try:
            setattr(cell, "CharComplexColor", complex_color)
        except Exception:
            pass


def _set_explicit_char_color(cell, color_int: Optional[int]) -> None:
    """Set font colour to an explicit RGB value, fully overriding any theme.

    Uses the ``.uno:Color`` dispatch command when a frame is available.
    This is critical because it properly clears the internal ComplexColor
    theme reference (Type 4 → 0), unlike ``setPropertyValue("CharColor")``
    which only sets the RGB value but leaves the ComplexColor struct pointing
    at a theme slot.  The XLSX exporter reads ComplexColor first; if its type
    is Scheme (4) it writes ``theme=N`` regardless of CharColor.

    Falls back to direct property setting when no dispatch context exists
    (e.g. in unit-test stubs).
    """
    if cell is None or color_int is None:
        return
    color_int = int(color_int)

    # The .uno:Color dispatch path is required to clear the ComplexColor
    # theme reference that the XLSX exporter otherwise serializes as
    # ``theme=N``.  This must be used for formula/value cells too: imported
    # subtotal cells can ignore direct CharColor/CharComplexColor writes on
    # export while the live Calc view still looks correct.  We only ever
    # select real ScCellObj cells here; _iterate_cells avoids text-portion
    # objects, which are the path that can coerce numeric cells to text.
    use_dispatch = _DISPATCH_CONTROLLER is not None and _DISPATCH_FRAME is not None

    if use_dispatch:
        try:
            _DISPATCH_CONTROLLER.select(cell)
            from com.sun.star.beans import PropertyValue as _PV
            _args = (_PV("Color", 0, color_int, 0),)
            ctx = uno.getComponentContext()
            dispatcher = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper", ctx)
            dispatcher.executeDispatch(
                _DISPATCH_FRAME, ".uno:Color", "", 0, _args)
        except Exception:
            pass  # fall through to direct-property path

    # --- Export guard: direct property set, with ComplexColor forced to RGB ---
    # Even after a successful .uno:Color dispatch the live view can be correct
    # while stale theme/interop color state remains. The XLSX exporter may use
    # that stale state on save, so always force the stored color model to RGB.
    #
    # Order still matters. Setting CharColor can itself leave/recreate theme
    # metadata in some Calc style paths, so clear both before and after the
    # direct color set. The final state seen by XLSX export must be explicit
    # RGB, not theme/automatic.
    _clear_char_color_theme_state(cell)
    _set_char_complex_color_rgb(cell, color_int)
    try:
        cell.CharColor = color_int
    except Exception:
        _set_uno_property(cell, "CharColor", color_int)
    _clear_char_color_theme_state(cell)
    _set_char_complex_color_rgb(cell, color_int)
    try:
        cell.CharColor = color_int
    except Exception:
        _set_uno_property(cell, "CharColor", color_int)


def _maybe_apply_thousands_separator_format(document, sheet_names) -> None:
    """
    Apply readable numeric formats to General numeric cells. Large values get
    thousands separators, fractional values render up to 4 decimal places, and
    formula outputs are included so recalculated General cells do not display
    long raw floats. This runs before autofit so column widths reflect the
    formatted display.
    """
    if not sheet_names:
        return
    try:
        number_formats = document.getNumberFormats()
        locale = _resolve_locale(number_formats)
        integer_format_string = "#,##0"
        decimal_format_string = "#,##0.####;(#,##0.####);-"
        integer_fmt_id = number_formats.queryKey(integer_format_string, locale, False)
        if integer_fmt_id == -1:
            integer_fmt_id = number_formats.addNew(integer_format_string, locale)
        decimal_fmt_id = number_formats.queryKey(decimal_format_string, locale, False)
        if decimal_fmt_id == -1:
            decimal_fmt_id = number_formats.addNew(decimal_format_string, locale)
    except Exception as exc:
        _log(f"ApplyExcelActionPlan: number format init failed: {exc}")
        return

    fmt_cache: Dict[str, int] = {}
    for sheet_name in sheet_names:
        sheet = _get_sheet(document, sheet_name)
        if sheet is None:
            continue
        try:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(True)
            addr = cursor.getRangeAddress()
            start_col, end_col = addr.StartColumn, addr.EndColumn
            start_row, end_row = addr.StartRow, addr.EndRow
        except Exception as exc:
            _log(
                f"ApplyExcelActionPlan: used-area scan failed for '{sheet_name}': {exc}"
            )
            continue

        if end_row < start_row or end_col < start_col:
            continue

        # Block fetch the whole used area: two UNO bridge calls per sheet
        # instead of three per cell. We use the result to skip empty / text
        # cells without crossing the bridge — numeric VALUE and FORMULA cells
        # get per-cell calls below.
        try:
            block = sheet.getCellRangeByPosition(
                start_col, start_row, end_col, end_row
            )
            formula_grid = block.getFormulaArray()
            data_grid = block.getDataArray()
        except Exception as exc:
            _log(
                f"ApplyExcelActionPlan: block read failed for '{sheet_name}': {exc}"
            )
            continue

        for r, formula_row in enumerate(formula_grid):
            data_row = data_grid[r] if r < len(data_grid) else ()
            for c, formula_text in enumerate(formula_row):
                # ``formula_text`` is "" for empty cells, the formula for
                # formula cells (starts with "="), and the literal value (as
                # string) for VALUE / TEXT cells. Formula cells are intentionally
                # included if their evaluated value is numeric.
                if not isinstance(formula_text, str) or formula_text == "":
                    continue
                value = data_row[c] if c < len(data_row) else None
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    continue  # text or other non-numeric — skip

                # This is a numeric VALUE or FORMULA cell. Cross the bridge now.
                try:
                    cell = sheet.getCellByPosition(start_col + c, start_row + r)
                except Exception:
                    continue
                _maybe_set_horizontal_alignment(cell, HORI_RIGHT)

                try:
                    current_fmt = cell.NumberFormat
                except Exception:
                    current_fmt = None
                fmt_string = _number_format_string(number_formats, current_fmt)
                target_fmt_id: Optional[int] = None
                if _is_general_number_format(number_formats, current_fmt):
                    numeric_value = float(value)
                    has_fraction = (
                        math.isfinite(numeric_value)
                        and not numeric_value.is_integer()
                    )
                    if has_fraction:
                        target_fmt_id = decimal_fmt_id
                    elif abs(numeric_value) >= 1000:
                        target_fmt_id = integer_fmt_id
                    else:
                        continue
                elif fmt_string and "," in fmt_string:
                    continue
                elif fmt_string:
                    grouped_fmt = _inject_grouping_separator_into_format(fmt_string)
                    if grouped_fmt != fmt_string:
                        cached_fmt = fmt_cache.get(grouped_fmt)
                        if cached_fmt is None:
                            try:
                                cached_fmt = number_formats.queryKey(
                                    grouped_fmt, locale, False
                                )
                                if cached_fmt == -1:
                                    cached_fmt = number_formats.addNew(
                                        grouped_fmt, locale
                                    )
                                fmt_cache[grouped_fmt] = cached_fmt
                            except Exception:
                                cached_fmt = None
                        target_fmt_id = cached_fmt
                if target_fmt_id in (None, -1):
                    continue
                try:
                    cell.NumberFormat = int(target_fmt_id)
                except Exception:
                    continue
                contrast_font_color = _dark_fill_contrast_font_color(cell)
                if contrast_font_color is not None:
                    _set_explicit_char_color(cell, contrast_font_color)
        _log(f"ApplyExcelActionPlan: applied numeric display format on '{sheet_name}'")


def _is_currency_number_format_string(fmt_string: str) -> bool:
    if not isinstance(fmt_string, str):
        return False
    text = fmt_string.strip()
    if not text:
        return False
    if "[$" in text:
        return True
    if re.search(r'"[^"]+"\s*\*', text):
        return True
    currency_symbols = "$€£¥₹₦₩₽₪฿₫"
    return any(symbol in text for symbol in currency_symbols)


def _is_text_number_format_string(fmt_string: str) -> bool:
    if not isinstance(fmt_string, str):
        return False
    normalized = fmt_string.strip().lower()
    return normalized in ("@", "text")


def _inject_grouping_separator_into_format(fmt_string: str) -> str:
    if not isinstance(fmt_string, str):
        return fmt_string
    if "," in fmt_string:
        return fmt_string
    match = re.search(r"([#0]+)(\.[#0]+)?", fmt_string)
    if not match:
        return fmt_string
    decimals = match.group(2) or ""
    grouped = f"#,##0{decimals}"
    return f"{fmt_string[:match.start()]}{grouped}{fmt_string[match.end():]}"


def _maybe_apply_currency_separator_format(document, sheet_names) -> None:
    """
    Ensure currency number formats include thousands separators for larger values.
    """
    if not sheet_names:
        return
    try:
        number_formats = document.getNumberFormats()
        locale = _resolve_locale(number_formats)
    except Exception as exc:
        _log(f"ApplyExcelActionPlan: currency format init failed: {exc}")
        return

    fmt_cache: Dict[str, int] = {}
    for sheet_name in sheet_names:
        sheet = _get_sheet(document, sheet_name)
        if sheet is None:
            continue
        try:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(True)
            addr = cursor.getRangeAddress()
            start_col, end_col = addr.StartColumn, addr.EndColumn
            start_row, end_row = addr.StartRow, addr.EndRow
        except Exception:
            continue

        if end_row < start_row or end_col < start_col:
            continue

        # Block fetch the whole used area. Currency reformatting only touches
        # numeric VALUE/FORMULA cells with abs(value) >= 1000 — typically a
        # small subset — so this lets us skip the bridge for every other cell.
        try:
            block = sheet.getCellRangeByPosition(
                start_col, start_row, end_col, end_row
            )
            formula_grid = block.getFormulaArray()
            data_grid = block.getDataArray()
        except Exception as exc:
            _log(
                f"ApplyExcelActionPlan: block read failed for '{sheet_name}': {exc}"
            )
            continue

        for r, formula_row in enumerate(formula_grid):
            data_row = data_grid[r] if r < len(data_grid) else ()
            for c, formula_text in enumerate(formula_row):
                # Same numeric-cell filter as the display-format pass: skip
                # empty and text cells without crossing the UNO bridge. Formula
                # cells are included if their evaluated value is numeric.
                if not isinstance(formula_text, str) or formula_text == "":
                    continue
                value = data_row[c] if c < len(data_row) else None
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    continue
                if abs(float(value)) < 1000:
                    continue

                try:
                    cell = sheet.getCellByPosition(start_col + c, start_row + r)
                except Exception:
                    continue
                try:
                    current_fmt = cell.NumberFormat
                except Exception:
                    current_fmt = None
                fmt_string = _number_format_string(number_formats, current_fmt)
                if not fmt_string or not _is_currency_number_format_string(fmt_string):
                    continue
                if "," in fmt_string:
                    continue
                grouped_fmt = _inject_grouping_separator_into_format(fmt_string)
                if grouped_fmt == fmt_string:
                    continue
                fmt_id = fmt_cache.get(grouped_fmt)
                if fmt_id is None:
                    try:
                        fmt_id = number_formats.queryKey(grouped_fmt, locale, False)
                        if fmt_id == -1:
                            fmt_id = number_formats.addNew(grouped_fmt, locale)
                        fmt_cache[grouped_fmt] = fmt_id
                    except Exception:
                        continue
                try:
                    cell.NumberFormat = fmt_id
                except Exception:
                    continue
                contrast_font_color = _dark_fill_contrast_font_color(cell)
                if contrast_font_color is not None:
                    _set_explicit_char_color(cell, contrast_font_color)


def _force_ui_refresh(controller) -> None:
    """Force Calc UI to repaint/refresh so changes are visible immediately."""
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()  # type: ignore[name-defined]
        sm = ctx.getServiceManager()
        dispatcher = sm.createInstance("com.sun.star.frame.DispatchHelper")
        frame = controller.getFrame()
        dispatcher.executeDispatch(frame, ".uno:Repaint", "", 0, ())
        dispatcher.executeDispatch(frame, ".uno:Refresh", "", 0, ())
        _log("ApplyExcelActionPlan: dispatched .uno:Repaint and .uno:Refresh")
    except Exception as refresh_exc:
        _log(f"ApplyExcelActionPlan: UI refresh failed: {refresh_exc}")


def _maybe_autofit_after_calculate(
    document,
    controller,
    sheet_names,
    excluded_columns_by_sheet: Optional[Dict[str, set]] = None,
    ranges_by_sheet: Optional[Dict[str, List[tuple[int, int, int, int]]]] = None,
) -> None:
    """
    Run column autofit on every touched sheet after a plan finishes.

    Columns in ``excluded_columns_by_sheet`` (explicitly sized by a
    ``set_column_width`` action) are skipped so the agent's deliberate
    width choices are preserved. Every other column on the touched sheet
    gets autofit with a max cap (_MAX_AUTOFIT_WIDTH_MM100) so long-text
    cells can't stretch columns across the page.
    """
    if not sheet_names:
        return
    excluded_columns_by_sheet = excluded_columns_by_sheet or {}
    ranges_by_sheet = ranges_by_sheet or {}
    for sheet_name in sheet_names:
        try:
            sheet = _get_sheet(document, sheet_name)
            if sheet is None:
                continue
            # Only autofit columns in the used area instead of all 1024+.
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(True)
            addr = cursor.getRangeAddress()
            if addr.EndColumn < 0 or addr.EndRow < 0:
                continue
            used_range = f"A1:{_column_index_to_name(addr.EndColumn)}{addr.EndRow + 1}"
            excluded_idxs = excluded_columns_by_sheet.get(sheet_name) or set()
            _handle_autofit_columns(
                document,
                controller,
                {
                    "kind": "autofit_columns",
                    "target": {"sheet_name": sheet_name, "range": used_range},
                    "payload": {
                        "sheet_name": sheet_name,
                        "_excluded_column_indexes": list(excluded_idxs),
                    },
                },
            )
            _widen_numeric_overflow_columns_after_autofit(
                document,
                sheet,
                sheet_name,
                ranges_by_sheet.get(sheet_name) or [
                    (addr.StartColumn, addr.StartRow, addr.EndColumn, addr.EndRow)
                ],
                excluded_idxs,
            )
            _log(
                f"ApplyExcelActionPlan: autofit_columns after calculate for "
                f"'{sheet_name}' (excluded cols: {sorted(excluded_idxs)})"
            )
        except Exception as exc:
            _log(
                f"ApplyExcelActionPlan: autofit_columns after calculate failed for '{sheet_name}': {exc}"
            )


def _color_luminance(color_int: int) -> float:
    red = (int(color_int) >> 16) & 0xFF
    green = (int(color_int) >> 8) & 0xFF
    blue = int(color_int) & 0xFF
    return (0.2126 * red) + (0.7152 * green) + (0.0722 * blue)


def _dark_fill_contrast_font_color(cell) -> Optional[int]:
    """Return the explicit font color required for readable dark-filled cells."""
    if cell is None:
        return None
    try:
        if bool(getattr(cell, "IsCellBackgroundTransparent", False)):
            return None
        bg_color = int(getattr(cell, "CellBackColor", -1))
    except Exception:
        return None
    if bg_color < 0 or bg_color == 0xFFFFFF:
        return None
    if _color_luminance(bg_color) >= 140:
        return None
    return 0xFFFFFF


def _maybe_repair_dark_fill_font_contrast(
    document,
    sheet_names,
    ranges_by_sheet: Optional[Dict[str, List[tuple[int, int, int, int]]]] = None,
) -> None:
    """
    Final XLSX-export guard: dark-filled cells must persist with explicit
    white RGB text, including formula/value cells.

    Calc can display a direct CharColor while still retaining a theme
    ComplexColor/interop entry. On save, XLSX export may serialize that stale
    theme color instead, so reload shows black text on dark subtotal/header
    fills. This pass runs after all writes and forces explicit white text for
    touched dark-fill cells.
    """
    if not sheet_names:
        return
    repaired = 0
    scanned = 0

    for sheet_name in sheet_names:
        try:
            sheet = _get_sheet(document, sheet_name)
            if sheet is None:
                continue
            # Always scan the used area of touched sheets. The bad exporter
            # state most often appears on existing subtotal/total formula
            # cells after later number-format or proposal passes, so limiting
            # this to only the current action ranges can leave dark rows stale.
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(True)
            addr = cursor.getRangeAddress()
            if addr.EndColumn < 0 or addr.EndRow < 0:
                continue
            scan_ranges = [
                (
                    int(addr.StartColumn),
                    int(addr.StartRow),
                    int(addr.EndColumn),
                    int(addr.EndRow),
                )
            ]

            for start_col, start_row, end_col, end_row in scan_ranges:
                for row_idx in range(int(start_row), int(end_row) + 1):
                    for col_idx in range(int(start_col), int(end_col) + 1):
                        try:
                            cell = sheet.getCellByPosition(col_idx, row_idx)
                            scanned += 1
                            font_color = _dark_fill_contrast_font_color(cell)
                            if font_color is None:
                                continue
                            _set_explicit_char_color(cell, font_color)
                            repaired += 1
                        except Exception:
                            continue
        except Exception as exc:
            _log(
                f"ApplyExcelActionPlan: dark-fill contrast repair failed for '{sheet_name}': {exc}"
            )

    if repaired:
        _log(
            "ApplyExcelActionPlan: repaired dark-fill font contrast "
            f"cells={repaired} scanned={scanned}"
        )


def _widen_numeric_overflow_columns_after_autofit(
    document,
    sheet,
    sheet_name: str,
    scan_ranges: List[tuple[int, int, int, int]],
    excluded_columns: set,
) -> None:
    """
    Post-autofit guard for numeric/formula cells that still display ###.
    This mainly protects accounting-style currency formats, where Calc can
    underestimate width even though autofit ran after recalc. It scans only
    ranges touched by the current plan, not the whole used sheet.
    """
    try:
        number_formats = document.getNumberFormats()
        columns = sheet.getColumns()
    except Exception as exc:
        _log(
            f"ApplyExcelActionPlan: numeric overflow scan failed for '{sheet_name}': {exc}"
        )
        return

    desired_width_by_col: Dict[int, int] = {}
    for start_col, start_row, end_col, end_row in scan_ranges:
        try:
            block = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
            formula_grid = block.getFormulaArray()
            data_grid = block.getDataArray()
        except Exception as exc:
            _log(
                f"ApplyExcelActionPlan: numeric overflow range scan failed for "
                f"'{sheet_name}' {start_col},{start_row}:{end_col},{end_row}: {exc}"
            )
            continue
        for r, formula_row in enumerate(formula_grid):
            data_row = data_grid[r] if r < len(data_grid) else ()
            for c, formula_text in enumerate(formula_row):
                col_idx = start_col + c
                if col_idx in excluded_columns:
                    continue
                if not isinstance(formula_text, str) or formula_text == "":
                    continue
                value = data_row[c] if c < len(data_row) else None
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    continue
                try:
                    cell = sheet.getCellByPosition(col_idx, start_row + r)
                    display_text = str(cell.getString() or "").strip()
                    fmt_string = _number_format_string(number_formats, cell.NumberFormat) or ""
                except Exception:
                    continue
                is_overflow = bool(display_text) and set(display_text) == {"#"}
                is_accounting = bool(re.search(r'"[^"]+"\s*\*', fmt_string))
                if not is_overflow and not is_accounting:
                    continue
                desired_chars = _estimate_numeric_display_chars(float(value), fmt_string)
                if desired_chars <= 0:
                    continue
                desired_width = min(_MAX_AUTOFIT_WIDTH_MM100, _chars_to_mm100(desired_chars))
                desired_width_by_col[col_idx] = max(
                    desired_width_by_col.get(col_idx, 0), desired_width
                )

    widened: List[str] = []
    for col_idx, desired_width in desired_width_by_col.items():
        try:
            column = columns.getByIndex(col_idx)
            current = int(getattr(column, "Width", 0) or 0)
            if desired_width > current:
                column.Width = desired_width
                widened.append(_column_index_to_name(col_idx))
        except Exception:
            continue
    if widened:
        _log(
            f"ApplyExcelActionPlan: widened numeric overflow/accounting columns "
            f"on '{sheet_name}': {widened}"
        )


def _columns_from_action_range(
    document, action: Dict[str, Any], default_sheet: Optional[str] = None
) -> list[tuple[str, int]]:
    """
    Extract (sheet_name, col_idx) pairs for every column covered by the action's range.
    Used to remember which columns were explicitly sized so they're excluded
    from the post-plan autofit pass.
    """
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = (
        payload.get("sheet_name")
        or target_info.get("sheet_name")
        or default_sheet
    )
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not sheet_name or not range_name:
        return []
    sheet, resolved_range = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "_columns_from_action_range"
    )
    if sheet is None or not resolved_range:
        return []
    try:
        cell_range = sheet.getCellRangeByName(resolved_range)
        address = cell_range.getRangeAddress()
    except Exception:
        return []
    return [(sheet_name, idx) for idx in range(address.StartColumn, address.EndColumn + 1)]


def _get_cell_delete_mode_up():
    global _CELL_DELETE_MODE_UP
    if _CELL_DELETE_MODE_UP is None:
        try:
            _CELL_DELETE_MODE_UP = uno.getConstantByName(
                "com.sun.star.sheet.CellDeleteMode.UP"
            )
        except Exception:
            _CELL_DELETE_MODE_UP = False
    return _CELL_DELETE_MODE_UP


def _get_cell_delete_mode_left():
    global _CELL_DELETE_MODE_LEFT
    if _CELL_DELETE_MODE_LEFT is None:
        try:
            _CELL_DELETE_MODE_LEFT = uno.getConstantByName(
                "com.sun.star.sheet.CellDeleteMode.LEFT"
            )
        except Exception:
            _CELL_DELETE_MODE_LEFT = False
    return _CELL_DELETE_MODE_LEFT


def _get_cell_insert_mode_right():
    global _CELL_INSERT_MODE_RIGHT
    if _CELL_INSERT_MODE_RIGHT is None:
        try:
            _CELL_INSERT_MODE_RIGHT = uno.getConstantByName(
                "com.sun.star.sheet.CellInsertMode.RIGHT"
            )
        except Exception:
            _CELL_INSERT_MODE_RIGHT = False
    return _CELL_INSERT_MODE_RIGHT


def _normalize_header_label(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def _normalize_column_placement_payload(
    payload: Dict[str, Any], action_name: str
) -> Tuple[str, Optional[str]]:
    placement = payload.get("placement") or {}
    position = str(placement.get("position") or "END").strip().upper()
    relative_to_column = placement.get("relative_to_column")

    if position not in {"END", "BEFORE", "AFTER"}:
        raise RuntimeError(
            f"{action_name} placement.position must be END, BEFORE, or AFTER"
        )

    if position == "END":
        if relative_to_column:
            raise RuntimeError(
                f"{action_name} placement.relative_to_column must be omitted when position is END"
            )
        return position, None

    normalized = _normalize_header_label(relative_to_column)
    if not normalized:
        raise RuntimeError(
            f"{action_name} placement.relative_to_column is required when position is {position}"
        )
    return position, str(relative_to_column)


def _resolve_column_insert_location(
    header: List[Any], address, payload: Dict[str, Any], action_name: str
) -> Tuple[str, int, int]:
    position, relative_to_column = _normalize_column_placement_payload(
        payload, action_name
    )
    start_col = int(address.StartColumn)
    end_col = int(address.EndColumn)

    if position == "END":
        new_col_index = end_col + 1
        template_col_index = max(start_col, new_col_index - 1)
        return position, new_col_index, template_col_index

    header_map: Dict[str, int] = {}
    for idx, name in enumerate(header):
        normalized = _normalize_header_label(name)
        if normalized and normalized not in header_map:
            header_map[normalized] = idx

    relative_key = _normalize_header_label(relative_to_column)
    if relative_key not in header_map:
        raise RuntimeError(
            f"{action_name} placement.relative_to_column '{relative_to_column}' not found in header"
        )

    relative_abs_col = start_col + header_map[relative_key]
    if position == "BEFORE":
        new_col_index = relative_abs_col
        template_col_index = new_col_index + 1
    else:
        new_col_index = relative_abs_col + 1
        template_col_index = max(start_col, new_col_index - 1)

    return position, new_col_index, template_col_index


def _copy_cell_style(source, target) -> None:
    if source is None or target is None:
        return
    style_attrs = [
        "CellStyle",
        "ConditionalFormat",
        "CharStyleName",
        "CharWeight",
        "CharColor",
        "CharHeight",
        "CharUnderline",
        "CharPosture",
        "CellBackColor",
        "HoriJustify",
        "VertJustify",
        "NumberFormat",
        "CellProtection",
    ]
    for attr in style_attrs:
        try:
            setattr(target, attr, getattr(source, attr))
        except Exception:
            continue


def _get_cell_property(cell, prop_name: str) -> Any:
    if cell is None or not prop_name:
        return None
    try:
        if _has_prop(cell, prop_name):
            getter = getattr(cell, "getPropertyValue", None)
            if callable(getter):
                return getter(prop_name)
            return getattr(cell, prop_name)
    except Exception:
        return None
    return None


def _set_cell_property(cell, prop_name: str, value: Any) -> bool:
    if cell is None or not prop_name:
        return False
    try:
        if not _has_prop(cell, prop_name):
            return False
        setter = getattr(cell, "setPropertyValue", None)
        if callable(setter):
            setter(prop_name, value)
            return True
        setattr(cell, prop_name, value)
        return True
    except Exception:
        return False


def _capture_column_property_matrix(
    sheet,
    prop_name: str,
    start_col: int,
    end_col: int,
    start_row: int,
    end_row: int,
) -> Optional[Dict[int, List[Any]]]:
    if start_col > end_col or start_row > end_row:
        return {}

    captured: Dict[int, List[Any]] = {}
    saw_property = False
    for col in range(start_col, end_col + 1):
        column_values: List[Any] = []
        for row in range(start_row, end_row + 1):
            try:
                cell = sheet.getCellByPosition(col, row)
            except Exception:
                column_values.append(None)
                continue
            value = _get_cell_property(cell, prop_name)
            if value is not None:
                saw_property = True
            column_values.append(value)
        captured[col] = column_values
    if not saw_property:
        return None
    return captured


def _shift_column_property_right_from_snapshot(
    sheet,
    prop_name: str,
    snapshot: Optional[Dict[int, List[Any]]],
    inserted_col: int,
    original_end_col: int,
    start_row: int,
    end_row: int,
) -> None:
    if not snapshot:
        return

    for source_col in range(inserted_col, original_end_col + 1):
        values = snapshot.get(source_col)
        if values is None:
            continue
        dest_col = source_col + 1
        for offset, row in enumerate(range(start_row, end_row + 1)):
            try:
                cell = sheet.getCellByPosition(dest_col, row)
            except Exception:
                continue
            if offset >= len(values):
                continue
            _set_cell_property(cell, prop_name, values[offset])


def _copy_column_styles(
    sheet, source_col_index: int, target_col_index: int, start_row: int, end_row: int
):
    if source_col_index < 0 or target_col_index < 0:
        return
    for row in range(start_row, end_row + 1):
        try:
            source_cell = sheet.getCellByPosition(source_col_index, row)
            target_cell = sheet.getCellByPosition(target_col_index, row)
            _copy_cell_style(source_cell, target_cell)
        except Exception:
            continue


def _cell_has_annotation(cell) -> bool:
    if cell is None:
        return False
    potential = [
        getattr(cell, attr, None) for attr in ("Annotation", "AnnotationObject")
    ]
    for annotation in potential:
        if annotation:
            try:
                text = getattr(annotation, "String", None)
                if text not in (None, ""):
                    return True
                # Even empty annotations count as meaningful content (comments).
                return True
            except Exception:
                return True
    getter = getattr(cell, "getAnnotation", None)
    if callable(getter):
        try:
            annotation = getter()
            if annotation:
                text = getattr(annotation, "String", None)
                if text not in (None, ""):
                    return True
                return True
        except Exception:
            pass
    return False


def _row_from_data_has_content(row: Iterable[Any]) -> bool:
    if row is None:
        return False
    for value in row:
        if value is None:
            continue
        if isinstance(value, str):
            if value.strip():
                return True
        else:
            return True
    return False


def _row_contains_content(sheet, row_index: int, start_col: int, end_col: int) -> bool:
    if row_index < 0:
        return False
    for col in range(start_col, end_col + 1):
        try:
            cell = sheet.getCellByPosition(col, row_index)
        except Exception:
            continue
        if _cell_has_visible_value(cell) or _cell_has_annotation(cell):
            return True
    return False


def _find_last_data_row(
    sheet, start_col: int, end_col: int, initial_last_row: int
) -> int:
    last_row = initial_last_row
    try:
        max_rows = sheet.getRows().getCount()
    except Exception:
        max_rows = 1_048_576  # Excel compatibility upper bound
    candidate = last_row + 1
    while candidate < max_rows and _row_contains_content(
        sheet, candidate, start_col, end_col
    ):
        last_row = candidate
        candidate += 1
    return max(last_row, initial_last_row)


def _insert_rows_via_sheet(
    sheet, start_row: int, row_count: int, start_col: int, end_col: int
) -> bool:
    try:
        max_col_index = sheet.getColumns().getCount() - 1
    except Exception:
        max_col_index = max(end_col, 0)
    try:
        insert_range = sheet.getCellRangeByPosition(
            start_col,
            start_row,
            min(max_col_index, max(start_col, end_col)),
            start_row + row_count - 1,
        )
        sheet.insertCells(insert_range.getRangeAddress(), INSERT_DOWN)
        return True
    except Exception as exc:
        _log(f"append_rows: sheet.insertCells DOWN fallback failed: {exc}")
    try:
        insert_range = sheet.getCellRangeByPosition(
            0,
            start_row,
            max_col_index,
            start_row + row_count - 1,
        )
        sheet.insertCells(insert_range.getRangeAddress(), INSERT_ROWS)
        return True
    except Exception as exc:
        _log(f"append_rows: sheet.insertCells ROWS fallback failed: {exc}")
        return False


def _insert_columns_via_sheet(
    sheet, start_col: int, column_count: int, start_row: int, end_row: int
) -> bool:
    mode_right = _get_cell_insert_mode_right()
    if not mode_right:
        return False
    try:
        insert_range = sheet.getCellRangeByPosition(
            start_col,
            start_row,
            start_col + column_count - 1,
            end_row,
        )
        sheet.insertCells(insert_range.getRangeAddress(), mode_right)
        return True
    except Exception as exc:
        _log(f"insert_column: sheet.insertCells RIGHT failed: {exc}")
        return False


def _shift_rows_down_manual(
    sheet, start_col: int, end_col: int, start_row: int, row_count: int
) -> bool:
    try:
        max_rows = sheet.getRows().getCount()
    except Exception:
        max_rows = 1_048_576
    last_row = _find_last_data_row(sheet, start_col, end_col, start_row)
    if start_row + row_count >= max_rows:
        return False
    try:
        for row in range(last_row, start_row - 1, -1):
            dest_row = row + row_count
            if dest_row >= max_rows:
                continue
            src_range = sheet.getCellRangeByPosition(start_col, row, end_col, row)
            dest_range = sheet.getCellRangeByPosition(
                start_col, dest_row, end_col, dest_row
            )
            try:
                sheet.copyRange(dest_range.getCellAddress(), src_range)
            except Exception:
                data = src_range.getDataArray()
                dest_range.setDataArray(data)
        _clear_range(sheet, start_col, end_col, start_row, start_row + row_count - 1)
        return True
    except Exception as exc:
        _log(f"append_rows: manual shift failed: {exc}")
        return False


def _shift_columns_right_manual(
    sheet, start_col: int, end_col: int, start_row: int, end_row: int, column_count: int
) -> bool:
    try:
        max_cols = int(sheet.getColumns().getCount())
    except Exception:
        max_cols = 1024
    if end_col + column_count >= max_cols:
        return False
    try:
        for col in range(end_col, start_col - 1, -1):
            dest_col = col + column_count
            src_range = sheet.getCellRangeByPosition(col, start_row, col, end_row)
            dest_range = sheet.getCellRangeByPosition(
                dest_col, start_row, dest_col, end_row
            )
            try:
                sheet.copyRange(dest_range.getCellAddress(), src_range)
            except Exception:
                data = src_range.getDataArray()
                dest_range.setDataArray(data)
        _clear_range(sheet, start_col, start_col + column_count - 1, start_row, end_row)
        return True
    except Exception as exc:
        _log(f"insert_column: manual right shift failed: {exc}")
        return False


def _clear_range(
    sheet, start_col: int, end_col: int, start_row: int, end_row: int
) -> None:
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            try:
                cell = sheet.getCellByPosition(col, row)
            except Exception:
                continue
            try:
                cell.setString("")
            except Exception:
                pass
            try:
                cell.setValue(0)
            except Exception:
                pass


def _clone_column_format(
    sheet, source_col_index: int, target_col_index: int, start_row: int, end_row: int
):
    if source_col_index < 0 or target_col_index < 0:
        return
    source_range = sheet.getCellRangeByPosition(
        source_col_index, start_row, source_col_index, end_row
    )
    target_range = sheet.getCellRangeByPosition(
        target_col_index, start_row, target_col_index, end_row
    )
    try:
        sheet.copyRange(target_range.getCellAddress(), source_range)
    except Exception:
        _copy_column_styles(
            sheet, source_col_index, target_col_index, start_row, end_row
        )


def _cell_has_visible_value(cell) -> bool:
    if cell is None:
        return False
    try:
        text = cell.getString()
        if text:
            return True
    except Exception:
        pass
    try:
        value = cell.getValue()
        if isinstance(value, (int, float)) and value != 0:
            return True
    except Exception:
        pass
    return False


def _resolve_locale(number_formats) -> Optional[Any]:
    try:
        settings = number_formats.getNumberFormatSettings()
        locale = getattr(settings, "Locale", None)
        if locale:
            return locale
    except Exception:
        pass
    try:
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language = "en"
        locale.Country = "US"
        return locale
    except Exception:
        return None


def _parse_hex_color_to_int(value: Any) -> Optional[int]:
    """
    Accept common color encodings and return a LibreOffice color int (0xRRGGBB).

    Supported inputs:
    - "RRGGBB"
    - "#RRGGBB"
    - "0xRRGGBB"
    - "RGB" / "#RGB" shorthand
    - 8-hex values (e.g., AARRGGBB): last 6 hex digits are used
    - int (passed through)
    """
    if value is None:
        return None
    if isinstance(value, int):
        return value

    try:
        text = str(value).strip()
    except Exception:
        return None

    if not text:
        return None

    if text.startswith("#"):
        text = text[1:]
    if text.lower().startswith("0x"):
        text = text[2:]

    # Allow alpha-prefixed values; keep RGB portion.
    if len(text) == 8:
        text = text[-6:]

    # Expand shorthand.
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)

    try:
        return int(text, 16)
    except Exception:
        return None


def _apply_format_to_cell(
    cell, format_spec: Dict[str, Any], number_formats, locale
) -> None:
    if not format_spec:
        return
    force_no_wrap = bool(format_spec.get("_force_no_wrap"))
    reapply_font_color_after_number_format: Optional[int] = None

    font_spec = format_spec.get("font") or {}
    explicit_font_color_int = None
    if font_spec:
        if font_spec.get("name"):
            cell.CharFontName = font_spec["name"]
        if font_spec.get("size"):
            try:
                cell.CharHeight = float(font_spec["size"])
            except Exception:
                pass
        if font_spec.get("bold") is True:
            cell.CharWeight = FONT_WEIGHT_BOLD
        elif font_spec.get("bold") is False:
            cell.CharWeight = 100
        if font_spec.get("italic") is True:
            cell.CharPosture = FONT_SLANT_ITALIC
        elif font_spec.get("italic") is False:
            cell.CharPosture = 0
        if font_spec.get("underline") is True:
            cell.CharUnderline = FONT_UNDERLINE_SINGLE
        elif font_spec.get("underline") is False:
            cell.CharUnderline = 0
        if font_spec.get("color"):
            color_int = _parse_hex_color_to_int(font_spec.get("color"))
            if color_int is not None:
                explicit_font_color_int = color_int

    fill_spec = format_spec.get("fill") or {}
    if fill_spec.get("color"):
        color_int = _parse_hex_color_to_int(fill_spec.get("color"))
        if color_int is not None:
            _set_uno_property(cell, "IsCellBackgroundTransparent", False)
            _set_uno_property(cell, "CellBackColor", color_int)
            red = (color_int >> 16) & 0xFF
            green = (color_int >> 8) & 0xFF
            blue = color_int & 0xFF
            luminance = (0.2126 * red) + (0.7152 * green) + (0.0722 * blue)
            desired_font_color = 0xFFFFFF if luminance < 140 else 0x000000

            # Contrast guard:
            # - If no explicit font color was provided, force readable text.
            # - If an explicit black/white color conflicts with fill luminance,
            #   override to keep dark fills readable after reloads/new plans.
            final_font_color = explicit_font_color_int
            if explicit_font_color_int is None:
                final_font_color = desired_font_color
            elif (
                (explicit_font_color_int == 0x000000 and desired_font_color == 0xFFFFFF)
                or (explicit_font_color_int == 0xFFFFFF and desired_font_color == 0x000000)
            ):
                final_font_color = desired_font_color

            # Font colour MUST be applied after fill/background changes.
            # Calc can re-derive colour/theme state when CellBackColor is set,
            # which leaves numeric/formula cells with black text on dark fills.
            _set_explicit_char_color(cell, final_font_color)
            reapply_font_color_after_number_format = final_font_color
    elif explicit_font_color_int is not None:
        _set_explicit_char_color(cell, explicit_font_color_int)
        reapply_font_color_after_number_format = explicit_font_color_int
        # No fill in this format spec, but the agent set an explicit font
        # color.  If the cell ALREADY has a fill that creates low contrast
        # (dark fill + dark font, or light fill + light font), clear the
        # fill so the agent's font is readable.
        #
        # Why: agents using APPLY_FORMATTING to "reset" body cells with
        # ``fill: {}`` and a font color expect the result to be a clean
        # black-on-white body row, not black-on-stale-navy.  Without this
        # cleanup, a body row that previously inherited a dark fill (from
        # prior plans, header style copy, or filter style) keeps it and
        # the contrast guard in the *previous* IF-branch never fires
        # (because there is no fill_spec.color to trigger it).
        try:
            current_bg = int(getattr(cell, "CellBackColor", -1))
            is_transparent = bool(getattr(cell, "IsCellBackgroundTransparent", False))
        except Exception:
            current_bg = -1
            is_transparent = True
        if (
            not is_transparent
            and current_bg >= 0
            and current_bg != 0xFFFFFF
        ):
            try:
                bg_red = (current_bg >> 16) & 0xFF
                bg_green = (current_bg >> 8) & 0xFF
                bg_blue = current_bg & 0xFF
                bg_lum = (
                    0.2126 * bg_red + 0.7152 * bg_green + 0.0722 * bg_blue
                )
                font_red = (explicit_font_color_int >> 16) & 0xFF
                font_green = (explicit_font_color_int >> 8) & 0xFF
                font_blue = explicit_font_color_int & 0xFF
                font_lum = (
                    0.2126 * font_red + 0.7152 * font_green + 0.0722 * font_blue
                )
                # Both dark (luminance < 140) or both light (>= 140) ⇒
                # contrast failure.  Clear the existing fill so the agent's
                # explicit font color renders as intended.
                low_contrast = (bg_lum < 140 and font_lum < 140) or (
                    bg_lum >= 140 and font_lum >= 140
                )
                if low_contrast:
                    _set_uno_property(cell, "IsCellBackgroundTransparent", True)
                    # Reset to white explicitly: some Calc paths render
                    # IsCellBackgroundTransparent=True correctly only when
                    # CellBackColor is also a sane neutral.
                    _set_uno_property(cell, "CellBackColor", 0xFFFFFF)
            except Exception:
                pass

    alignment_spec = format_spec.get("alignment") or {}
    if alignment_spec:
        horiz = (alignment_spec.get("horizontal") or "").lower()
        try:
            if horiz == "left":
                cell.HoriJustify = HORI_LEFT
            elif horiz == "center":
                cell.HoriJustify = HORI_CENTER
            elif horiz == "right":
                cell.HoriJustify = HORI_RIGHT
        except AttributeError:
            pass  # Cell doesn't support HoriJustify

        vert = (alignment_spec.get("vertical") or "").lower()
        try:
            if vert == "top":
                cell.VertJustify = VERT_TOP
            elif vert == "center":
                cell.VertJustify = VERT_CENTER
            elif vert == "bottom":
                cell.VertJustify = VERT_BOTTOM
        except AttributeError:
            pass  # Cell doesn't support VertJustify

        if alignment_spec.get("wrap_text") is not None:
            cell.IsTextWrapped = False if force_no_wrap else bool(alignment_spec["wrap_text"])
        elif force_no_wrap:
            cell.IsTextWrapped = False
    elif force_no_wrap:
        cell.IsTextWrapped = False

    if format_spec.get("number_format"):
        try:
            format_string = format_spec["number_format"]
            if locale is None:
                locale = _resolve_locale(number_formats)
            fmt_id = number_formats.queryKey(format_string, locale, False)
            if fmt_id == -1:
                fmt_id = number_formats.addNew(format_string, locale)
            cell.NumberFormat = fmt_id
            if reapply_font_color_after_number_format is not None:
                _set_explicit_char_color(cell, reapply_font_color_after_number_format)
            else:
                contrast_font_color = _dark_fill_contrast_font_color(cell)
                if contrast_font_color is not None:
                    _set_explicit_char_color(cell, contrast_font_color)
        except Exception:
            pass


def _extract_cell_value(cell) -> Any:
    text_value = ""
    try:
        text_value = cell.getString()
    except Exception:
        text_value = ""

    numeric_value: Optional[float] = None
    try:
        numeric_value = cell.getValue()
    except Exception:
        numeric_value = None

    if text_value:
        return text_value
    if numeric_value is not None:
        return numeric_value
    return text_value


def _expand_to_single_cell(cell):
    """
    Ensure the cursor target is a single-cell range.
    """
    if cell is None:
        return None

    get_cell_address = getattr(cell, "getCellAddress", None)
    if callable(get_cell_address):
        address = get_cell_address()
        sheet = cell.getSpreadsheet()
        return sheet.getCellRangeByPosition(
            address.Column, address.Row, address.Column, address.Row
        )

    get_range_address = getattr(cell, "getRangeAddress", None)
    if callable(get_range_address):
        address = get_range_address()
        sheet = cell.getSpreadsheet()
        return sheet.getCellRangeByPosition(
            address.StartColumn, address.StartRow, address.StartColumn, address.StartRow
        )

    raise RuntimeError("Cannot resolve a single-cell cursor target from this UNO object")


def _coerce_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def _coerce_plain_numeric_string(value: Any) -> Optional[float]:
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    if not stripped:
        return None
    if not _PLAIN_NUMERIC_TEXT_RE.fullmatch(stripped):
        return None
    normalized = stripped.replace(",", "")
    try:
        return float(normalized)
    except ValueError:
        return None


def _aggregate_row(
    values: Iterable[Optional[float]], operation: str
) -> Optional[float]:
    op = operation.upper()
    numeric = [v for v in values if v is not None]
    if op == "SUM":
        return sum(numeric) if numeric else 0.0
    if op == "AVERAGE":
        return sum(numeric) / len(numeric) if numeric else None
    if op == "MAX":
        return max(numeric) if numeric else None
    if op == "MIN":
        return min(numeric) if numeric else None
    if op == "COUNT":
        return float(len([v for v in values if v is not None]))
    return None


def _coerce_string(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _compare_values(cell_value: Any, operator: str, expected: Any) -> bool:
    op = operator.upper()
    left_num = _coerce_number(cell_value)
    right_num = _coerce_number(expected)

    if op in {"GREATER_THAN", "GREATER_OR_EQUAL", "LESS_THAN", "LESS_OR_EQUAL"}:
        if left_num is None or right_num is None:
            return False
        if op == "GREATER_THAN":
            return left_num > right_num
        if op == "GREATER_OR_EQUAL":
            return left_num >= right_num
        if op == "LESS_THAN":
            return left_num < right_num
        if op == "LESS_OR_EQUAL":
            return left_num <= right_num

    left_str = _coerce_string(cell_value).strip()
    right_str = _coerce_string(expected).strip()

    if op == "EQUALS":
        return left_str == right_str
    if op == "NOT_EQUALS":
        return left_str != right_str
    if op == "CONTAINS":
        return right_str.lower() in left_str.lower()
    if op == "NOT_CONTAINS":
        return right_str.lower() not in left_str.lower()
    if op == "STARTS_WITH":
        return left_str.startswith(right_str)
    if op == "ENDS_WITH":
        return left_str.endswith(right_str)
    if op == "IN":
        expected_list = expected if isinstance(expected, list) else [expected]
        return left_str in [str(item) for item in expected_list]
    if op == "NOT_IN":
        expected_list = expected if isinstance(expected, list) else [expected]
        return left_str not in [str(item) for item in expected_list]
    if op == "BETWEEN":
        if not isinstance(expected, (list, tuple)) or len(expected) != 2:
            return False
        low = _coerce_number(expected[0])
        high = _coerce_number(expected[1])
        if left_num is None or low is None or high is None:
            return False
        return low <= left_num <= high

    return False


def _row_matches_conditions(
    row_values: list[Any],
    header_map: Dict[str, int],
    conditions: Iterable[Dict[str, Any]],
    width: Optional[int] = None,
    start_column: int = 0,
) -> bool:
    if not conditions:
        return True
    resolved_width = width if width is not None else len(row_values)
    for condition in conditions:
        column_name = condition.get("column")
        operator = condition.get("operator")
        value = condition.get("value")
        col_index = _resolve_update_cells_column_index(
            header_map, column_name, resolved_width, start_column
        )
        if col_index is None or not operator:
            return False
        if col_index >= len(row_values):
            return False
        cell_value = row_values[col_index]
        if not _compare_values(cell_value, operator, value):
            return False
    return True


def _handle_autofit_columns(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    padding_mm100 = payload.get("padding_mm100")
    padding_chars = payload.get("padding_chars")
    if padding_mm100 is None and padding_chars is not None:
        try:
            padding_mm100 = _chars_to_mm100(float(padding_chars))
        except Exception:
            padding_mm100 = None
    if padding_mm100 is None:
        padding_mm100 = _DEFAULT_AUTOFIT_PADDING_MM100
    try:
        padding_mm100 = int(padding_mm100)
    except Exception:
        padding_mm100 = _DEFAULT_AUTOFIT_PADDING_MM100

    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "autofit_columns"
    )
    if sheet is None:
        raise RuntimeError(f"autofit_columns sheet '{sheet_name}' not found")
    columns = sheet.getColumns()

    # Columns the caller wants skipped (e.g. ones explicitly sized by the
    # agent via set_column_width). Pairs are 0-based column indexes.
    excluded = set(payload.get("_excluded_column_indexes") or [])

    if range_name:
        column_range = sheet.getCellRangeByName(range_name)
        address = column_range.getRangeAddress()
        for idx in range(address.StartColumn, address.EndColumn + 1):
            if idx in excluded:
                continue
            try:
                column = columns.getByIndex(idx)
                _autofit_column_with_padding(column, padding_mm100)
            except Exception:
                continue
        return column_range

    columns_count = columns.getCount()
    for idx in range(columns_count):
        if idx in excluded:
            continue
        try:
            column = columns.getByIndex(idx)
            _autofit_column_with_padding(column, padding_mm100)
        except Exception:
            continue
    return sheet.getCellRangeByName("A1")


def _resolve_sheet_and_range(
    document, ref: str, default_sheet
) -> tuple[Optional[Any], Optional[Any]]:
    sheet_name, range_body = _split_sheet_qualified_range_name(ref)
    target_sheet = _get_sheet(document, sheet_name) if sheet_name else default_sheet
    if target_sheet is None or range_body is None:
        return None, None
    try:
        cell_range = target_sheet.getCellRangeByName(range_body)
        return target_sheet, cell_range
    except Exception:
        return None, None



def _combine_ranges(range_a, range_b):
    if range_a is None:
        return range_b
    if range_b is None:
        return range_a
    address_a = range_a.getRangeAddress()
    address_b = range_b.getRangeAddress()
    if address_a.Sheet != address_b.Sheet:
        return None
    start_col = min(address_a.StartColumn, address_b.StartColumn)
    end_col = max(address_a.EndColumn, address_b.EndColumn)
    start_row = min(address_a.StartRow, address_b.StartRow)
    end_row = max(address_a.EndRow, address_b.EndRow)
    sheet = range_a.getSpreadsheet()
    try:
        return sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    except Exception:
        return None


def _range_address_bounds(cell_range):
    if cell_range is None:
        return None
    try:
        address = cell_range.getRangeAddress()
    except Exception:
        return None
    return {
        "start_row": getattr(address, "StartRow", None),
        "end_row": getattr(address, "EndRow", None),
        "start_column": getattr(address, "StartColumn", None),
        "end_column": getattr(address, "EndColumn", None),
    }


def _maybe_extend_with_header(cell_range):
    """Extend a range upward by one row if the row above contains a text header.

    Returns ``(extended_range, was_extended)`` so callers can adjust label hints
    when the range was silently grown to include headers.
    """
    if cell_range is None:
        return None, False
    try:
        address = cell_range.getRangeAddress()
        sheet = cell_range.getSpreadsheet()
    except Exception:
        return cell_range, False
    if address.StartRow <= 0:
        return cell_range, False
    try:
        header_cell = sheet.getCellByPosition(address.StartColumn, address.StartRow - 1)
        header_value = header_cell.getString()
    except Exception:
        header_value = ""
    if not header_value:
        return cell_range, False
    try:
        extended = sheet.getCellRangeByPosition(
            address.StartColumn,
            address.StartRow - 1,
            address.EndColumn,
            address.EndRow,
        )
        return extended, True
    except Exception:
        return cell_range, False


def _chart_rect_overlaps_existing(sheet, chart_rect) -> bool:
    """Return True if *chart_rect* overlaps any existing chart or used-area content on *sheet*.

    This is the runtime safety-net: if the agent requested an explicit placement
    that would land on top of another chart (or on top of cell content like a
    title row), we detect it here and reject the action.
    """
    if sheet is None or chart_rect is None:
        return False

    # Check against existing charts
    try:
        charts = sheet.getCharts()
        if charts:
            for name in charts.getElementNames():
                try:
                    chart_obj = charts.getByName(name)
                    existing_rect = _chart_rect_for_object(chart_obj, sheet=sheet)
                    if existing_rect is not None and _rectangles_overlap(
                        chart_rect,
                        existing_rect,
                    ):
                        return True
                except Exception:
                    continue
    except Exception:
        pass

    # Check the cells covered by the chart rectangle. This avoids treating a
    # sparse sheet's whole used-area bounding box as occupied.
    try:
        chart_area = _cell_range_from_chart_rect(sheet, chart_rect)
        if chart_area is not None and _cell_range_has_content(chart_area):
            return True
    except Exception:
        pass

    return False


def _rect_to_dict(rect) -> Optional[Dict[str, int]]:
    if rect is None:
        return None
    try:
        return {
            "x_mm100": int(getattr(rect, "X", 0) or 0),
            "y_mm100": int(getattr(rect, "Y", 0) or 0),
            "width_mm100": int(getattr(rect, "Width", 0) or 0),
            "height_mm100": int(getattr(rect, "Height", 0) or 0),
        }
    except Exception:
        return None


def _placement_from_rect(rect) -> Optional[Dict[str, Any]]:
    return _rect_to_dict(rect)


def _range_diagnostics(cell_range) -> Dict[str, Optional[str]]:
    if cell_range is None:
        return {"range": None, "range_a1": None}
    range_abs = None
    range_a1 = None
    try:
        range_abs = _range_to_representation(cell_range)
    except Exception:
        range_abs = None
    try:
        range_a1 = _range_to_a1_relative(cell_range)
    except Exception:
        range_a1 = None
    return {"range": range_abs, "range_a1": range_a1}


def _sample_non_empty_cells(cell_range, limit: int = 10) -> List[Dict[str, Any]]:
    samples: List[Dict[str, Any]] = []
    if cell_range is None:
        return samples
    try:
        address = cell_range.getRangeAddress()
        sheet = cell_range.getSpreadsheet()
    except Exception:
        return samples
    try:
        sheet_name = _safe_sheet_name(sheet)
    except Exception:
        sheet_name = None
    try:
        for row in range(address.StartRow, address.EndRow + 1):
            for col in range(address.StartColumn, address.EndColumn + 1):
                if len(samples) >= limit:
                    return samples
                cell = sheet.getCellByPosition(col, row)
                value = None
                try:
                    value = cell.getString()
                except Exception:
                    value = None
                if isinstance(value, str) and not value.strip():
                    value = None
                if value is None:
                    try:
                        numeric = cell.getValue()
                        if numeric not in (None, 0):
                            value = numeric
                    except Exception:
                        value = None
                if value is None:
                    continue
                cell_ref = f"{_column_index_to_name(col)}{row + 1}"
                samples.append(
                    {
                        "sheet_name": sheet_name,
                        "cell": cell_ref,
                        "value_preview": str(value)[:80],
                    }
                )
    except Exception:
        return samples
    return samples


def _chart_rect_collision_details(
    sheet,
    chart_rect,
    *,
    requested_placement: Optional[Dict[str, Any]] = None,
    requested_area=None,
) -> Optional[Dict[str, Any]]:
    if sheet is None or chart_rect is None:
        return None
    sheet_name = _safe_sheet_name(sheet)
    area = requested_area or _cell_range_from_chart_rect(sheet, chart_rect)
    area_info = _range_diagnostics(area)
    details: Dict[str, Any] = {
        "error_code": "CHART_PLACEMENT_COLLISION",
        "sheet_name": sheet_name,
        "requested_placement": requested_placement,
        "requested_rect": _rect_to_dict(chart_rect),
        "requested_occupied_range": area_info.get("range"),
        "requested_occupied_range_a1": area_info.get("range_a1"),
    }

    try:
        charts = sheet.getCharts()
        names = list(charts.getElementNames()) if charts else []
    except Exception:
        names = []
    for name in names:
        try:
            chart_obj = charts.getByName(name)
            existing_rect = _chart_rect_for_object(chart_obj, sheet=sheet)
            if existing_rect is None or not _rectangles_overlap(chart_rect, existing_rect):
                continue
            existing_area = _chart_area_range_for_object(sheet, chart_obj)
            existing_area_info = _range_diagnostics(existing_area)
            details.update(
                {
                    "collision_type": "existing_chart",
                    "colliding_chart_name": str(name),
                    "colliding_chart_title": _chart_title_for_object(chart_obj),
                    "colliding_placement": _placement_from_rect(existing_rect),
                    "colliding_rect": _rect_to_dict(existing_rect),
                    "colliding_occupied_range": existing_area_info.get("range"),
                    "colliding_occupied_range_a1": existing_area_info.get("range_a1"),
                    "suggestion": (
                        "Choose a placement whose rectangle does not intersect the "
                        "colliding chart's occupied area."
                    ),
                }
            )
            return details
        except Exception:
            continue

    try:
        if area is not None and _cell_range_has_content(area):
            samples = _sample_non_empty_cells(area)
            details.update(
                {
                    "collision_type": "sheet_content",
                    "occupied_range": area_info.get("range"),
                    "occupied_range_a1": area_info.get("range_a1"),
                    "non_empty_cells_sample": samples,
                    "non_empty_cell_count_sampled": len(samples),
                    "suggestion": (
                        "Choose a placement whose covered cells are empty, or move "
                        "the chart away from the listed occupied cells."
                    ),
                }
            )
            return details
    except Exception:
        pass

    return None


def _shape_rect(shape):
    if shape is None:
        return None
    try:
        pos = shape.getPosition()
        size = shape.getSize()
        return _rectangle_from_values(
            getattr(pos, "X", None),
            getattr(pos, "Y", None),
            getattr(size, "Width", None),
            getattr(size, "Height", None),
        )
    except Exception:
        return None


def _image_rect_collision_details(
    sheet,
    image_rect,
    *,
    requested_placement: Optional[Dict[str, Any]] = None,
    requested_area=None,
    ignore_shape=None,
) -> Optional[Dict[str, Any]]:
    if sheet is None or image_rect is None:
        return None
    sheet_name = _safe_sheet_name(sheet)
    area = requested_area or _cell_range_from_chart_rect(sheet, image_rect)
    area_info = _range_diagnostics(area)
    details: Dict[str, Any] = {
        "error_code": "IMAGE_PLACEMENT_COLLISION",
        "sheet_name": sheet_name,
        "requested_placement": requested_placement,
        "requested_rect": _rect_to_dict(image_rect),
        "requested_occupied_range": area_info.get("range"),
        "requested_occupied_range_a1": area_info.get("range_a1"),
    }

    try:
        charts = sheet.getCharts()
        names = list(charts.getElementNames()) if charts else []
    except Exception:
        names = []
    for name in names:
        try:
            chart_obj = charts.getByName(name)
            existing_rect = _chart_rect_for_object(chart_obj, sheet=sheet)
            if existing_rect is None or not _rectangles_overlap(image_rect, existing_rect):
                continue
            details.update(
                {
                    "collision_type": "existing_chart",
                    "colliding_chart_name": str(name),
                    "colliding_chart_title": _chart_title_for_object(chart_obj),
                    "colliding_placement": _placement_from_rect(existing_rect),
                    "suggestion": "Choose image placement that does not intersect the existing chart.",
                }
            )
            return details
        except Exception:
            continue

    try:
        draw_page = sheet.getDrawPage()
        count = int(draw_page.getCount())
    except Exception:
        count = 0
    for idx in range(count):
        try:
            shape = draw_page.getByIndex(idx)
            if ignore_shape is not None and shape == ignore_shape:
                continue
            existing_rect = _shape_rect(shape)
            if existing_rect is None or not _rectangles_overlap(image_rect, existing_rect):
                continue
            details.update(
                {
                    "collision_type": "existing_drawing",
                    "colliding_shape_name": getattr(shape, "Name", None),
                    "colliding_placement": _placement_from_rect(existing_rect),
                    "suggestion": "Choose image placement that does not intersect an existing drawing object.",
                }
            )
            return details
        except Exception:
            continue

    try:
        if area is not None and _cell_range_has_content(area):
            samples = _sample_non_empty_cells(area)
            details.update(
                {
                    "collision_type": "sheet_content",
                    "occupied_range": area_info.get("range"),
                    "occupied_range_a1": area_info.get("range_a1"),
                    "non_empty_cells_sample": samples,
                    "non_empty_cell_count_sampled": len(samples),
                    "suggestion": "Choose image placement whose covered cells are empty.",
                }
            )
            return details
    except Exception:
        pass

    return None


def _image_rect_has_collision(sheet, image_rect) -> bool:
    return _image_rect_collision_details(sheet, image_rect) is not None


def _image_rect_and_area_from_placement(dest_sheet, placement):
    required_fields = ("x_mm100", "y_mm100", "width_mm100", "height_mm100")
    if not isinstance(placement, dict):
        raise ActionApplicationError(
            "image placement must be native LibreOffice geometry",
            error_code="IMAGE_PLACEMENT_INVALID",
            details={"required_fields": list(required_fields)},
        )
    legacy_keys = [
        key
        for key in ("mode", "top_left_cell", "bottom_right_cell", "anchor_cell", "width", "height")
        if key in placement
    ]
    if legacy_keys:
        raise ActionApplicationError(
            "image placement uses native geometry only",
            error_code="IMAGE_PLACEMENT_INVALID",
            details={
                "required_fields": list(required_fields),
                "rejected_fields": legacy_keys,
                "unit": "1/100 mm; x_mm100=24838 means 248.38 mm from the sheet left edge",
            },
        )
    missing = [key for key in required_fields if placement.get(key) is None]
    if missing:
        raise ActionApplicationError(
            "image placement is missing native geometry fields",
            error_code="IMAGE_PLACEMENT_INVALID",
            details={
                "required_fields": list(required_fields),
                "missing_fields": missing,
                "received_fields": sorted(str(key) for key in placement.keys()),
            },
        )
    rect = _rectangle_from_values(
        placement.get("x_mm100"),
        placement.get("y_mm100"),
        placement.get("width_mm100"),
        placement.get("height_mm100"),
    )
    if rect is None:
        raise ActionApplicationError(
            "image placement geometry values are invalid",
            error_code="IMAGE_PLACEMENT_INVALID",
            details={
                "required_fields": list(required_fields),
                "received_placement": placement,
                "require_positive": ["width_mm100", "height_mm100"],
            },
        )
    area = _cell_range_from_chart_rect(dest_sheet, rect) if dest_sheet is not None else None
    return rect, area


def _resolve_legacy_image_rect(sheet, anchor_cell: str, width: int, height: int):
    """Resolve older anchor_cell image actions to a non-overlapping drawing rect."""
    try:
        base_cell = sheet.getCellRangeByName(anchor_cell)
        base_rect = _get_cell_range_rect(sheet, base_cell)
    except Exception:
        base_cell = None
        base_rect = None
    if not base_rect:
        raise RuntimeError(f"insert_image could not resolve anchor cell '{anchor_cell}'")

    direct_rect = _rectangle_from_values(base_rect["X"], base_rect["Y"], width, height)
    direct_area = _cell_range_from_chart_rect(sheet, direct_rect)
    if not _image_rect_collision_details(sheet, direct_rect, requested_area=direct_area):
        return direct_rect, direct_area, base_cell

    start_col, start_row = _cell_name_to_indices(anchor_cell)
    if start_col is None:
        start_col = 0
    if start_row is None:
        start_row = 0

    # First try the blank drawing space to the right of the requested cell, then
    # continue in chart-like grid rows below it. This keeps row tables readable
    # when an old agent asks to insert several images at C5/C6/C7.
    for row_slot in range(_IMAGE_MAX_ROW_SCAN):
        candidate_row = max(0, start_row + row_slot * _IMAGE_ROW_SLOT_STRIDE)
        for col_slot in range(_IMAGE_MAX_COLUMN_SCAN):
            candidate_col = max(0, start_col + 1 + col_slot * 4)
            try:
                candidate_cell = sheet.getCellByPosition(candidate_col, candidate_row)
                candidate_props = _get_cell_range_rect(sheet, candidate_cell)
            except Exception:
                continue
            if not candidate_props:
                continue
            rect = _rectangle_from_values(candidate_props["X"], candidate_props["Y"], width, height)
            if rect is None:
                continue
            area = _cell_range_from_chart_rect(sheet, rect)
            if not _image_rect_collision_details(sheet, rect, requested_area=area):
                return rect, area, candidate_cell

    ua = _sheet_used_area(sheet)
    fallback_col = start_col
    fallback_row = start_row + _IMAGE_ROW_SLOT_STRIDE
    if ua is not None:
        fallback_col = max(0, int(getattr(ua, "EndColumn", fallback_col) or fallback_col) + 2)
        fallback_row = max(0, int(getattr(ua, "StartRow", fallback_row) or fallback_row))
    for row_slot in range(_IMAGE_MAX_ROW_SCAN):
        candidate_row = fallback_row + row_slot * _IMAGE_ROW_SLOT_STRIDE
        try:
            candidate_cell = sheet.getCellByPosition(fallback_col, candidate_row)
            candidate_props = _get_cell_range_rect(sheet, candidate_cell)
        except Exception:
            continue
        if not candidate_props:
            continue
        rect = _rectangle_from_values(candidate_props["X"], candidate_props["Y"], width, height)
        area = _cell_range_from_chart_rect(sheet, rect)
        if not _image_rect_collision_details(sheet, rect, requested_area=area):
            return rect, area, candidate_cell

    return direct_rect, direct_area, base_cell


def _cell_ref_from_chart_rect(sheet, chart_rect):
    """Resolve the cell containing a chart rectangle's top-left corner.

    ``chart_rect`` is a ``com.sun.star.awt.Rectangle`` in mm/100 (the native
    units for sheet column/row Width and Height). Walks the column widths and
    row heights to find which cell owns (X, Y).
    """
    if sheet is None or chart_rect is None:
        return None
    try:
        target_x = int(getattr(chart_rect, "X", 0))
        target_y = int(getattr(chart_rect, "Y", 0))
    except Exception:
        return None

    try:
        columns = sheet.getColumns()
        rows = sheet.getRows()
    except Exception:
        return None

    col_idx = 0
    cum_x = 0
    # Cap at 1024 to bound the scan on defensive grounds.
    for _ in range(1024):
        try:
            width = int(getattr(columns.getByIndex(col_idx), "Width", 0) or 0)
        except Exception:
            break
        if width <= 0:
            break
        if cum_x + width > target_x:
            break
        cum_x += width
        col_idx += 1

    row_idx = 0
    cum_y = 0
    for _ in range(100000):
        try:
            height = int(getattr(rows.getByIndex(row_idx), "Height", 0) or 0)
        except Exception:
            break
        if height <= 0:
            break
        if cum_y + height > target_y:
            break
        cum_y += height
        row_idx += 1

    try:
        return sheet.getCellByPosition(col_idx, row_idx)
    except Exception:
        return None


def _cell_range_from_chart_rect(sheet, chart_rect):
    """Resolve the cell range covered by a chart rectangle."""
    if sheet is None or chart_rect is None:
        return None

    try:
        start_x = int(getattr(chart_rect, "X", 0))
        start_y = int(getattr(chart_rect, "Y", 0))
        end_x = max(start_x, start_x + int(getattr(chart_rect, "Width", 0)) - 1)
        end_y = max(start_y, start_y + int(getattr(chart_rect, "Height", 0)) - 1)
    except Exception:
        return None

    try:
        columns = sheet.getColumns()
        rows = sheet.getRows()
    except Exception:
        return None

    def _count(container, default_count: int) -> int:
        try:
            return max(1, min(int(container.getCount()), default_count))
        except Exception:
            return default_count

    def _index_for_coordinate(container, attr: str, coordinate: int, max_count: int) -> int:
        cumulative = 0
        last_index = 0
        for idx in range(max_count):
            last_index = idx
            try:
                span = int(getattr(container.getByIndex(idx), attr, 0) or 0)
            except Exception:
                break
            span = max(1, span)
            if cumulative + span > coordinate:
                return idx
            cumulative += span
        return last_index

    max_cols = _count(columns, 16384)
    max_rows = _count(rows, 1048576)
    start_col = _index_for_coordinate(columns, "Width", start_x, max_cols)
    end_col = _index_for_coordinate(columns, "Width", end_x, max_cols)
    start_row = _index_for_coordinate(rows, "Height", start_y, max_rows)
    end_row = _index_for_coordinate(rows, "Height", end_y, max_rows)

    try:
        return sheet.getCellRangeByPosition(
            min(start_col, end_col),
            min(start_row, end_row),
            max(start_col, end_col),
            max(start_row, end_row),
        )
    except Exception:
        return None


def _cell_range_has_content(cell_range) -> bool:
    if cell_range is None:
        return False
    try:
        data = cell_range.getDataArray()
    except Exception:
        data = ()
    try:
        for row in data or ():
            for value in row or ():
                if value is None:
                    continue
                if isinstance(value, str):
                    if value.strip():
                        return True
                    continue
                return True
    except Exception:
        pass
    try:
        address = cell_range.getRangeAddress()
        sheet = cell_range.getSpreadsheet()
        max_cells = 10000
        checked = 0
        for row in range(address.StartRow, address.EndRow + 1):
            for col in range(address.StartColumn, address.EndColumn + 1):
                checked += 1
                if checked > max_cells:
                    return True
                cell = sheet.getCellByPosition(col, row)
                text = ""
                try:
                    text = cell.getString()
                except Exception:
                    pass
                if text:
                    return True
                try:
                    value = cell.getValue()
                    if value not in (None, 0):
                        return True
                except Exception:
                    pass
    except Exception:
        pass
    return False


def _compute_chart_rectangle(sheet, target_info: Dict[str, Any], data_bounds=None):
    if isinstance(target_info, dict):
        try_range = (target_info.get("range") or "").strip()
        if try_range:
            _, range_only = _split_sheet_qualified_range_name(try_range)
            if range_only:
                try_range = range_only
            if ":" in try_range:
                try:
                    chart_area = sheet.getCellRangeByName(try_range)
                    rect = _build_chart_rectangle_from_range(chart_area)
                    if rect is not None:
                        return rect
                except Exception:
                    pass
    anchor = _resolve_chart_anchor(sheet, target_info, data_bounds=data_bounds)
    if anchor is None:
        raise RuntimeError("create_chart could not resolve placement anchor")
    return _build_chart_rectangle_from_anchor(anchor)


def _resolve_chart_anchor(
    sheet, target_info: Optional[Dict[str, Any]], data_bounds=None
):
    if sheet is None:
        return None

    if isinstance(target_info, dict):
        try_range = (target_info.get("range") or "").strip()
        if try_range:
            _, range_only = _split_sheet_qualified_range_name(try_range)
            if range_only:
                try_range = range_only
            try:
                return sheet.getCellRangeByName(try_range)
            except Exception:
                pass
        start_row = target_info.get("start_row")
        start_col = target_info.get("start_column")
        if start_row is not None and start_col is not None:
            try:
                row = max(0, int(start_row) - 1)
                col = max(0, int(start_col) - 1)
                return sheet.getCellRangeByPosition(col, row, col, row)
            except Exception:
                pass

    auto_anchor = _next_auto_chart_anchor(sheet, data_bounds=data_bounds)
    if auto_anchor is not None:
        return auto_anchor

    fallback_targets = [
        _DEFAULT_CHART_START_CELL,
        "A1",
    ]
    for candidate in fallback_targets:
        try:
            return sheet.getCellRangeByName(candidate)
        except Exception:
            continue
    try:
        return sheet.getCellRangeByPosition(0, 0, 0, 0)
    except Exception:
        return None


def _get_chart_sheet_state(sheet, data_bounds=None):
    global _CHART_LAYOUT_STATE
    if _CHART_LAYOUT_STATE is None:
        return None
    per_sheet = _CHART_LAYOUT_STATE.setdefault("per_sheet", {})
    sheet_name = _safe_sheet_name(sheet)
    if sheet_name is None:
        return None
    state = per_sheet.get(sheet_name)
    if state is None:
        # Seed with existing charts to continue after what already exists.
        try:
            existing = sheet.getCharts().getCount()
        except Exception:
            existing = 0
        base_col, base_row = _cell_name_to_indices(_DEFAULT_CHART_START_CELL)
        # If no data bounds and no existing charts, anchor tight to the default.
        if existing or data_bounds:
            ua = _sheet_used_area(sheet)
            if ua is not None:
                try:
                    end_row = ua.EndRow
                    if end_row is not None and isinstance(end_row, int):
                        base_row = max(
                            base_row or 0, int(end_row) + 1 + _CHART_ROW_PADDING
                        )
                except Exception:
                    pass
                try:
                    end_col = ua.EndColumn
                    if end_col is not None and isinstance(end_col, int):
                        base_col = max(
                            base_col or 0, int(end_col) + 1 + _CHART_COLUMN_PADDING
                        )
                except Exception:
                    pass
            # Also account for existing charts so the next chart continues after them.
            cb = _chart_bounds(sheet)
            if cb:
                if isinstance(cb.get("end_row"), int):
                    base_row = max(
                        base_row or 0, int(cb.get("end_row")) + 1 + _CHART_ROW_PADDING
                    )
                if isinstance(cb.get("end_col"), int):
                    base_col = max(
                        base_col or 0,
                        int(cb.get("end_col")) + 1 + _CHART_COLUMN_PADDING,
                    )
        state = {
            "sheet_name": sheet_name,
            "chart_count": int(existing or 0),
            "base_col": int(base_col or 0),
            "base_row": int(base_row or 0),
        }
        per_sheet[sheet_name] = state

    if data_bounds:
        end_row = data_bounds.get("end_row")
        end_col = data_bounds.get("end_column")
        if isinstance(end_row, int):
            state["base_row"] = max(
                int(state.get("base_row") or 0), int(end_row) + 1 + _CHART_ROW_PADDING
            )
        if isinstance(end_col, int):
            state["base_col"] = max(
                int(state.get("base_col") or 0),
                int(end_col) + 1 + _CHART_COLUMN_PADDING,
            )

    return state


def _next_auto_chart_anchor(sheet, data_bounds=None):
    state = _get_chart_sheet_state(sheet, data_bounds=data_bounds)
    if state is None:
        return None
    count = int(state.get("chart_count") or 0)
    base_col = state.get("base_col")
    base_row = state.get("base_row")
    if base_col is None or base_row is None:
        base_col, base_row = _cell_name_to_indices(_DEFAULT_CHART_START_CELL)
        if base_col is None or base_row is None:
            base_col, base_row = 7, 1  # H2 fallback
    col_slot = count % _CHART_GRID_COLUMNS
    row_slot = count // _CHART_GRID_COLUMNS
    col_step = _CHART_COLUMN_STRIDE + max(_CHART_COLUMN_GAP, 0)
    row_step = _CHART_ROW_STRIDE + max(_CHART_ROW_GAP, 0)
    target_col = base_col + (col_slot * col_step)
    target_row = base_row + (row_slot * row_step)
    try:
        anchor = sheet.getCellRangeByPosition(
            target_col, target_row, target_col, target_row
        )
    except Exception:
        try:
            anchor = sheet.getCellRangeByName(_DEFAULT_CHART_START_CELL)
        except Exception:
            anchor = None
    state["chart_count"] = count + 1
    return anchor


def _sheet_used_area(sheet):
    """Return the used area range address or None if unavailable.

    Uses a cursor to detect the bounding box of non-empty cells on the sheet.
    """
    try:
        cursor = sheet.createCursor()
        try:
            cursor.gotoStartOfUsedArea(False)
        except Exception:
            pass
        try:
            cursor.gotoEndOfUsedArea(True)
        except Exception:
            pass
        try:
            return cursor.getRangeAddress()
        except Exception:
            return None
    except Exception:
        return None


def _safe_sheet_name(sheet):
    if sheet is None:
        return None
    getter = getattr(sheet, "getName", None)
    if not callable(getter):
        return None
    try:
        return getter()
    except Exception:
        return None


def _chart_bounds(sheet):
    """Return max row/col occupied by existing charts on the sheet."""
    try:
        charts = sheet.getCharts()
    except Exception:
        return None
    try:
        names = charts.getElementNames()
    except Exception:
        names = []
    max_row = None
    max_col = None
    for name in names:
        try:
            chart = charts.getByName(name)
        except Exception:
            continue
        chart_area = _chart_area_range_for_object(sheet, chart)
        if chart_area is None:
            continue
        try:
            addr = chart_area.getRangeAddress()
            end_row = getattr(addr, "EndRow", None)
            end_col = getattr(addr, "EndColumn", None)
        except Exception:
            end_row = None
            end_col = None
        if isinstance(end_row, int):
            max_row = end_row if max_row is None else max(max_row, end_row)
        if isinstance(end_col, int):
            max_col = end_col if max_col is None else max(max_col, end_col)
    if max_row is None and max_col is None:
        return None
    return {"end_row": max_row, "end_col": max_col}


def _build_chart_rectangle_from_anchor(anchor):
    rectangle = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    if anchor is None:
        rectangle.X = 2000
        rectangle.Y = 2000
        rectangle.Width = 15000
        rectangle.Height = 8000
        return rectangle

    try:
        sheet = anchor.getSpreadsheet()
        address = anchor.getRangeAddress()
    except Exception:
        try:
            position = anchor.getPosition()
            size = anchor.getSize()
            rectangle.X = getattr(position, "X", 2000)
            rectangle.Y = getattr(position, "Y", 2000)
            rectangle.Width = getattr(size, "Width", 15000)
            rectangle.Height = getattr(size, "Height", 8000)
            return rectangle
        except Exception:
            rectangle.X = 2000
            rectangle.Y = 2000
            rectangle.Width = 15000
            rectangle.Height = 8000
            return rectangle

    end_col = address.StartColumn + max(0, _CHART_COLUMN_STRIDE - 1)
    end_row = address.StartRow + max(0, _CHART_ROW_STRIDE - 1)

    try:
        chart_area_range = sheet.getCellRangeByPosition(
            address.StartColumn,
            address.StartRow,
            end_col,
            end_row,
        )
        rect_props = _get_cell_range_rect(sheet, chart_area_range)
        if rect_props:
            rectangle.X = rect_props["X"]
            rectangle.Y = rect_props["Y"]
            rectangle.Width = rect_props["Width"]
            rectangle.Height = rect_props["Height"]
        else:
            rectangle.X = 2000
            rectangle.Y = 2000
            rectangle.Width = 15000
            rectangle.Height = 8000
    except Exception as exc:
        _log(f"Error creating chart rectangle from anchor: {exc}")
        rectangle.X = 2000
        rectangle.Y = 2000
        rectangle.Width = 15000
        rectangle.Height = 8000

    return rectangle


def _build_chart_rectangle_from_range(cell_range):
    rectangle = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    if cell_range is None:
        return None
    try:
        sheet = cell_range.getSpreadsheet()
        rect_props = _get_cell_range_rect(sheet, cell_range)
        if rect_props:
            rectangle.X = rect_props["X"]
            rectangle.Y = rect_props["Y"]
            rectangle.Width = rect_props["Width"]
            rectangle.Height = rect_props["Height"]
            return rectangle
    except Exception as exc:
        _log(f"Error creating chart rectangle from range: {exc}")
    return None


def _chart_rect_from_range_ref(document, dest_sheet, range_ref: str):
    if not range_ref:
        return None
    try:
        sheet, cell_range = _resolve_sheet_and_range(document, range_ref, dest_sheet)
    except Exception:
        sheet, cell_range = dest_sheet, None
    if cell_range is None:
        return None
    return _build_chart_rectangle_from_range(cell_range)


def _range_ref_is_multi_cell_area(range_ref: Optional[str]) -> bool:
    if not range_ref:
        return False
    _, range_only = _split_sheet_qualified_range_name(str(range_ref))
    if not range_only:
        return False
    return ":" in str(range_only)


def _rectangle_from_values(x, y, width, height):
    try:
        rectangle = uno.createUnoStruct("com.sun.star.awt.Rectangle")
        rectangle.X = int(x)
        rectangle.Y = int(y)
        rectangle.Width = int(width)
        rectangle.Height = int(height)
        if rectangle.Width <= 0 or rectangle.Height <= 0:
            return None
        return rectangle
    except Exception:
        return None


def _rectangles_overlap(rect_a, rect_b) -> bool:
    if rect_a is None or rect_b is None:
        return False
    try:
        ax = int(getattr(rect_a, "X", 0) or 0)
        ay = int(getattr(rect_a, "Y", 0) or 0)
        aw = int(getattr(rect_a, "Width", 0) or 0)
        ah = int(getattr(rect_a, "Height", 0) or 0)
        bx = int(getattr(rect_b, "X", 0) or 0)
        by = int(getattr(rect_b, "Y", 0) or 0)
        bw = int(getattr(rect_b, "Width", 0) or 0)
        bh = int(getattr(rect_b, "Height", 0) or 0)
    except Exception:
        return False
    if aw <= 0 or ah <= 0 or bw <= 0 or bh <= 0:
        return False
    return not (
        ax + aw <= bx
        or bx + bw <= ax
        or ay + ah <= by
        or by + bh <= ay
    )


def _cell_ref_without_sheet(ref: Optional[str]) -> Optional[str]:
    if not ref:
        return None
    _sheet_name, range_only = _split_sheet_qualified_range_name(str(ref))
    return range_only


def _chart_rect_and_area_from_placement(document, dest_sheet, placement):
    required_fields = ("x_mm100", "y_mm100", "width_mm100", "height_mm100")
    if not isinstance(placement, dict):
        raise ActionApplicationError(
            "chart placement must be native LibreOffice geometry",
            error_code="CHART_PLACEMENT_INVALID",
            details={"required_fields": list(required_fields)},
        )
    legacy_keys = [
        key
        for key in ("mode", "top_left_cell", "bottom_right_cell", "anchor_cell", "width_mm", "height_mm")
        if key in placement
    ]
    if legacy_keys:
        raise ActionApplicationError(
            "chart placement uses native geometry only",
            error_code="CHART_PLACEMENT_INVALID",
            details={
                "required_fields": list(required_fields),
                "rejected_fields": legacy_keys,
                "unit": "1/100 mm; x_mm100=24838 means 248.38 mm from the sheet left edge",
            },
        )
    missing = [key for key in required_fields if placement.get(key) is None]
    if missing:
        raise ActionApplicationError(
            "chart placement is missing native geometry fields",
            error_code="CHART_PLACEMENT_INVALID",
            details={
                "required_fields": list(required_fields),
                "missing_fields": missing,
                "received_fields": sorted(str(key) for key in placement.keys()),
            },
        )
    rect = _rectangle_from_values(
        placement.get("x_mm100"),
        placement.get("y_mm100"),
        placement.get("width_mm100"),
        placement.get("height_mm100"),
    )
    if rect is None:
        raise ActionApplicationError(
            "chart placement geometry values are invalid",
            error_code="CHART_PLACEMENT_INVALID",
            details={
                "required_fields": list(required_fields),
                "received_placement": placement,
                "require_positive": ["width_mm100", "height_mm100"],
            },
        )
    area = (
        _cell_range_from_chart_rect(dest_sheet, rect)
        if rect is not None and dest_sheet is not None
        else None
    )
    return rect, area


def _unique_chart_name(charts, base_name: str) -> str:
    safe_base = base_name.strip() or "Chart"
    candidate = safe_base
    attempt = 1
    while charts.hasByName(candidate):
        attempt += 1
        candidate = f"{safe_base}_{attempt}"
    return candidate


def _column_index_to_name(idx: int) -> str:
    idx += 1
    name = ""
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        name = chr(65 + remainder) + name
    return name


_CELL_REF_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")


def _cell_name_to_indices(cell_ref: str) -> Tuple[Optional[int], Optional[int]]:
    if not cell_ref:
        return None, None
    cleaned = cell_ref.strip()
    match = _CELL_REF_RE.match(cleaned)
    if not match:
        return None, None
    col_str, row_str = match.groups()
    col_idx = 0
    for ch in col_str.upper():
        col_idx = col_idx * 26 + (ord(ch) - 64)
    col_idx -= 1
    row_idx = int(row_str) - 1
    if col_idx < 0:
        col_idx = 0
    if row_idx < 0:
        row_idx = 0
    return col_idx, row_idx


def _range_to_representation(cell_range) -> str:
    """Return an absolute A1-style range with safe sheet quoting.

    Example: 'Sheet 1'.$A$1:$C$10
    """
    address = cell_range.getRangeAddress()
    sheet = cell_range.getSpreadsheet()
    sheet_get_name = getattr(sheet, "getName", None)
    if callable(sheet_get_name):
        sheet_name = sheet.getName()
    else:
        sheet_name = ""
    col_start = _column_index_to_name(address.StartColumn)
    col_end = _column_index_to_name(address.EndColumn)
    row_start = address.StartRow + 1
    row_end = address.EndRow + 1
    a1_start = f"${col_start}${row_start}"
    a1_end = f"${col_end}${row_end}"
    prefix = f"{_quote_sheet_name(sheet_name)}." if sheet_name else ""
    if row_start == row_end and col_start == col_end:
        return f"{prefix}{a1_start}"
    return f"{prefix}{a1_start}:{a1_end}"


def _quote_sheet_name(name: str) -> str:
    """Quote a sheet name when needed; escape internal quotes."""
    safe = name.replace("'", "''")
    if not re.match(r"^[A-Za-z0-9_]+$", safe):
        return f"'{safe}'"
    return safe


def _detect_chart_labels_from_range(cell_range) -> Tuple[bool, bool]:
    """Heuristically detect if the first row/column are labels. Returns (row, col)."""
    try:
        data = list(map(list, cell_range.getDataArray()))
    except Exception:
        data = []
    if not data:
        return False, False

    def _is_text(v: Any) -> bool:
        return isinstance(v, str) and v.strip() != ""

    first_row = data[0] if data else []
    first_col = [r[0] for r in data if r]
    row_labels = any(_is_text(v) for v in first_row)
    col_labels = any(_is_text(v) for v in first_col)
    return bool(row_labels), bool(col_labels)


_CHART_TYPE_ALIASES = {
    "surface": "surface3d",
    "donut": "doughnut",
}


def _normalize_chart_type(chart_type: str) -> str:
    value = str(chart_type or "").strip().lower()
    if not value:
        return ""
    return _CHART_TYPE_ALIASES.get(value, value)


def _chart_type_to_uno(chart_type: str) -> Optional[str]:
    chart_type = _normalize_chart_type(chart_type)
    mapping = {
        "bar": "com.sun.star.chart2.BarDiagram",
        "bar3d": "com.sun.star.chart2.BarDiagram",
        "column": "com.sun.star.chart2.ColumnDiagram",
        "column3d": "com.sun.star.chart2.ColumnDiagram",
        "line": "com.sun.star.chart2.LineDiagram",
        "line3d": "com.sun.star.chart2.LineDiagram",
        "pie": "com.sun.star.chart2.PieDiagram",
        "pie3d": "com.sun.star.chart2.PieDiagram",
        "doughnut": "com.sun.star.chart2.DoughnutDiagram",
        "area": "com.sun.star.chart2.AreaDiagram",
        "radar": "com.sun.star.chart2.NetDiagram",
        "scatter": "com.sun.star.chart2.XYDiagram",
        "bubble": "com.sun.star.chart2.BubbleDiagram",
        "stock": "com.sun.star.chart2.CandleStickDiagram",
        "combo": "com.sun.star.chart2.ColumnDiagram",
        "surface3d": "com.sun.star.chart2.SurfaceDiagram",
    }
    return mapping.get(chart_type)


def _chart_type_to_classic(chart_type: str) -> Optional[str]:
    """Map generic chart_type to classic Chart (chart1) services.

    These are the services used in LO SDK examples like ChartTypeChange.py and
    tend to be widely supported across hosts.
    """
    normalized = _normalize_chart_type(chart_type)
    mapping = {
        "bar": "com.sun.star.chart.BarDiagram",
        "bar3d": "com.sun.star.chart.BarDiagram",
        "column": "com.sun.star.chart.ColumnDiagram",
        "column3d": "com.sun.star.chart.ColumnDiagram",
        "line": "com.sun.star.chart.LineDiagram",
        "line3d": "com.sun.star.chart.LineDiagram",
        "pie": "com.sun.star.chart.PieDiagram",
        "pie3d": "com.sun.star.chart.PieDiagram",
        "doughnut": "com.sun.star.chart.DonutDiagram",  # sometimes DonutDiagram
        "radar": "com.sun.star.chart.NetDiagram",
        "scatter": "com.sun.star.chart.XYDiagram",
        "area": "com.sun.star.chart.AreaDiagram",
        "combo": "com.sun.star.chart.ColumnDiagram",
        # extras when requested by payloads
        "stock": "com.sun.star.chart.StockDiagram",
    }
    # Fallback: treat column as bar if classic ColumnDiagram is not recognized by host
    service = mapping.get(normalized)
    if service is None and normalized in {"column", "column3d"}:
        service = mapping.get("bar")
    return service


def _set_chart_3d_hint(diagram, chart_type: str):
    if diagram is None:
        return
    is_3d = chart_type in {"bar3d", "column3d", "line3d", "pie3d", "surface3d"}
    if hasattr(diagram, "Dim3D"):
        try:
            diagram.Dim3D = bool(is_3d)
        except Exception:
            pass


def _verify_chart_type_applied(chart_document, expected_type: str):
    """Check that a pie/doughnut chart was actually set. If not, retry with
    the classic chart API path which is more reliable for these types."""
    if chart_document is None:
        return
    expected_lower = _normalize_chart_type(expected_type)
    # Map expected type to service substrings we'd find in the diagram
    expected_markers = {
        "pie": ("pie", "Pie"),
        "pie3d": ("pie", "Pie"),
        "doughnut": ("donut", "Donut", "Doughnut", "doughnut"),
    }
    markers = expected_markers.get(expected_lower, ())
    if not markers:
        return

    # Check what diagram type is actually set
    try:
        diagram = chart_document.getDiagram()
        if diagram is None:
            return
        # Check the diagram's service name / implementation name
        diag_type = ""
        for attr in ("DiagramType", "getDiagramType"):
            try:
                val = getattr(diagram, attr, None)
                if callable(val):
                    val = val()
                if isinstance(val, str):
                    diag_type = val
                    break
            except Exception:
                continue

        if not diag_type:
            # Try supportsService as a check
            for marker in markers:
                svc = f"com.sun.star.chart.{marker.capitalize()}Diagram"
                try:
                    if diagram.supportsService(svc):
                        return  # correct type is set
                except Exception:
                    pass

        if diag_type:
            for marker in markers:
                if marker.lower() in diag_type.lower():
                    return  # correct type is set

        # Type mismatch — try the classic API as a more direct retry
        _log(
            f"_verify_chart_type_applied: expected {expected_lower} but got "
            f"'{diag_type}'; retrying with classic chart service"
        )
        classic_service = _chart_type_to_classic(expected_lower)
        if classic_service:
            try:
                new_diagram = chart_document.createInstance(classic_service)
                if new_diagram is not None:
                    _set_chart_3d_hint(new_diagram, expected_lower)
                    chart_document.setDiagram(new_diagram)
                    _log(f"_verify_chart_type_applied: retry with {classic_service} succeeded")
                    return
            except Exception as exc:
                _log(f"_verify_chart_type_applied: classic retry failed: {exc}")
    except Exception as exc:
        _log(f"_verify_chart_type_applied: verification error: {exc}")


def _apply_chart_type(chart_document, chart_type: str, payload: Optional[Dict[str, Any]] = None):
    chart_type_lower = _normalize_chart_type(chart_type) or ""
    base_type = "column" if chart_type_lower == "combo" else chart_type_lower
    diagram_service = _chart_type_to_uno(chart_type_lower)
    if chart_type_lower == "combo":
        diagram_service = _chart_type_to_uno("column")
    if not diagram_service or chart_document is None:
        return

    # 1) Prefer creating the diagram via the chart document's own factory or createInstance.
    direct_factory = getattr(chart_document, "createInstance", None)
    if callable(direct_factory):
        try:
            new_diagram = direct_factory(diagram_service)
            if new_diagram is not None:
                _set_chart_3d_hint(new_diagram, chart_type_lower)
                chart_document.setDiagram(new_diagram)
                return
        except Exception:
            pass

    # Some chart documents expose getFactory() which returns an XMultiServiceFactory.
    try:
        factory = None
        try:
            if callable(getattr(chart_document, "getFactory", None)):
                factory = chart_document.getFactory()
        except Exception:
            factory = None

        if factory is not None:
            try:
                new_diagram = factory.createInstance(diagram_service)
                if new_diagram is not None:
                    _set_chart_3d_hint(new_diagram, chart_type_lower)
                    chart_document.setDiagram(new_diagram)
                    return
            except Exception:
                # fall through to other fallbacks
                pass
    except Exception:
        pass

    # --- Classic Chart (com.sun.star.chart.*) fallback ---
    classic_service = _chart_type_to_classic(base_type)
    if classic_service:
        # Try createInstance on the chart document
        try:
            new_diagram = chart_document.createInstance(classic_service)
            if new_diagram is not None:
                try:
                    # Some classic diagrams expose Dim3D; default to 2D
                    if hasattr(new_diagram, "Dim3D"):
                        setattr(
                            new_diagram,
                            "Dim3D",
                            chart_type_lower
                            in {"bar3d", "column3d", "line3d", "pie3d", "surface3d"},
                        )
                except Exception:
                    pass
                chart_document.setDiagram(new_diagram)
                return
        except Exception:
            pass
        # Try the document's factory if available
        try:
            if callable(getattr(chart_document, "getFactory", None)):
                factory = chart_document.getFactory()
            else:
                factory = None
            if factory is not None:
                try:
                    new_diagram = factory.createInstance(classic_service)
                    if new_diagram is not None:
                        _set_chart_3d_hint(new_diagram, chart_type_lower)
                        chart_document.setDiagram(new_diagram)
                        return
                except Exception:
                    pass
        except Exception:
            pass
        # Try global service manager last
        try:
            ctx = None
            try:
                ctx = uno.getComponentContext()
            except Exception:
                ctx = None
            if ctx is not None:
                try:
                    smgr = ctx.getServiceManager()
                    new_diagram = smgr.createInstanceWithContext(classic_service, ctx)
                    if new_diagram is not None:
                        _set_chart_3d_hint(new_diagram, chart_type_lower)
                        chart_document.setDiagram(new_diagram)
                        return
                except Exception:
                    pass
        except Exception:
            pass

    # 2) Fallback to creating diagram via global service manager (older hosts).
    try:
        ctx = None
        try:
            ctx = uno.getComponentContext()
        except Exception:
            ctx = None
        if ctx is not None:
            try:
                smgr = ctx.getServiceManager()
                new_diagram = smgr.createInstanceWithContext(diagram_service, ctx)
                if new_diagram is not None:
                    _set_chart_3d_hint(new_diagram, chart_type_lower)
                    chart_document.setDiagram(new_diagram)
                    return
            except Exception:
                pass
    except Exception:
        pass

    # 3) Fallback to chart type manager helper if available.
    try:
        manager = getattr(chart_document, "getChartTypeManager", None)
        if callable(manager):
            try:
                mgr = chart_document.getChartTypeManager()
                if mgr is not None:
                    mgr.changeDiagramType(chart_document, diagram_service)
                    return
            except Exception:
                pass
    except Exception:
        pass

    # 4) Last resort: try to call setDiagramType on existing diagram (legacy).
    try:
        diagram = chart_document.getDiagram()
        if diagram is not None:
            setter = getattr(diagram, "setDiagramType", None)
            if callable(setter):
                try:
                    setter(diagram_service)
                    return
                except Exception:
                    pass
    except Exception:
        pass


def _style_chart_document(
    chart_document,
    chart_type: str,
    chart_title: Optional[str],
    payload: Optional[Dict[str, Any]] = None,
):
    if chart_document is None:
        return

    chart_type_lower = _normalize_chart_type(chart_type)
    is_non_cartesian = ("pie" in chart_type_lower) or ("doughnut" in chart_type_lower)
    x_axis_title = (payload or {}).get("x_axis_title")
    y_axis_title = (payload or {}).get("y_axis_title")
    legend_visible = (payload or {}).get("legend")
    legend_position = (payload or {}).get("legend_position")
    show_gridlines = (payload or {}).get("show_gridlines")
    show_data_labels = (payload or {}).get("show_data_labels")
    x_axis_rotation = (payload or {}).get("x_axis_label_rotation")
    y_axis_rotation = (payload or {}).get("y_axis_label_rotation")
    y_axis_number_format = (payload or {}).get("y_axis_number_format")
    stack_mode = (payload or {}).get("stack_mode")

    if legend_visible is None:
        legend_visible = True
    if show_data_labels is None and is_non_cartesian:
        show_data_labels = True

    chart_title_text = str(chart_title).strip() if chart_title is not None else ""
    x_axis_title_text = (
        str(x_axis_title).strip() if x_axis_title is not None else None
    )
    y_axis_title_text = (
        str(y_axis_title).strip() if y_axis_title is not None else None
    )

    # Title (chart2 + classic fallback): set text and explicitly show.
    try:
        has_main_title_prop = _has_prop(chart_document, "HasMainTitle") or hasattr(
            chart_document, "HasMainTitle"
        )
        if chart_title is not None and has_main_title_prop:
            try:
                chart_document.setPropertyValue("HasMainTitle", bool(chart_title_text))
            except Exception:
                try:
                    chart_document.HasMainTitle = bool(chart_title_text)
                except Exception:
                    pass
        if chart_title_text:
            try:
                title_obj = chart_document.getTitle()
            except Exception:
                title_obj = None
            if title_obj is not None:
                if hasattr(title_obj, "String"):
                    title_obj.String = chart_title_text
                if hasattr(title_obj, "Show"):
                    title_obj.Show = True
            elif _has_prop(chart_document, "Title"):
                # Some classic interfaces expose Title directly.
                try:
                    title_obj = chart_document.getPropertyValue("Title")
                    if title_obj is not None and hasattr(title_obj, "String"):
                        title_obj.String = chart_title_text
                except Exception:
                    pass
    except Exception as e:
        _log(f"_style_chart_document: title error: {e}")

    # Legend: explicitly show and set position if possible.
    try:
        has_legend_prop = _has_prop(chart_document, "HasLegend") or hasattr(
            chart_document, "HasLegend"
        )
        if legend_visible is not None and has_legend_prop:
            try:
                chart_document.setPropertyValue("HasLegend", bool(legend_visible))
            except Exception:
                try:
                    chart_document.HasLegend = bool(legend_visible)
                except Exception:
                    pass

        legend_obj = chart_document.getLegend()
        if legend_obj is not None:
            if hasattr(legend_obj, "Show") and legend_visible is not None:
                legend_obj.Show = bool(legend_visible)
            if legend_position:
                try:
                    from com.sun.star.chart.LegendPosition import (
                        RIGHT as LEGEND_RIGHT,
                        LEFT as LEGEND_LEFT,
                        TOP as LEGEND_TOP,
                        BOTTOM as LEGEND_BOTTOM,
                    )

                    pos_map = {
                        "right": LEGEND_RIGHT,
                        "left": LEGEND_LEFT,
                        "top": LEGEND_TOP,
                        "bottom": LEGEND_BOTTOM,
                    }
                    pos_key = str(legend_position).lower()
                    pos_val = pos_map.get(pos_key)
                    if pos_val is not None:
                        if hasattr(legend_obj, "Alignment"):
                            legend_obj.Alignment = pos_val
                        elif _has_prop(legend_obj, "AnchorPosition"):
                            legend_obj.setPropertyValue("AnchorPosition", pos_val)
                    elif pos_key in ("none", "off", "false", "no"):
                        if hasattr(legend_obj, "Show"):
                            legend_obj.Show = False
                except Exception:
                    pass
    except Exception as e:
        _log(f"_style_chart_document: legend error: {e}")

    try:
        diagram = chart_document.getDiagram()
    except Exception:
        diagram = None
    chart2_diagram = None
    try:
        get_first_diagram = getattr(chart_document, "getFirstDiagram", None)
        if callable(get_first_diagram):
            chart2_diagram = chart_document.getFirstDiagram()
    except Exception:
        chart2_diagram = None
    if chart2_diagram is None and diagram is not None:
        # Some hosts return a chart2-capable object from getDiagram.
        if hasattr(diagram, "getCoordinateSystems"):
            chart2_diagram = diagram
    diagram_for_axes = chart2_diagram or diagram
    if diagram_for_axes is None:
        return

    _apply_stack_mode(diagram_for_axes, stack_mode)

    # Axes and titles (chart2 path)
    x_axis_title_applied = False
    y_axis_title_applied = False
    for axis_dimension in (0, 1):
        try:
            axis = diagram_for_axes.getAxisByDimension(axis_dimension, 0)  # type: ignore[attr-defined]
            if axis is None:
                continue
            if hasattr(axis, "Show"):
                axis.Show = not is_non_cartesian
            # Prefer text categories
            if hasattr(axis, "AxisType"):
                try:
                    axis.AxisType = 2
                except Exception:
                    pass
            if axis_dimension == 0 and x_axis_rotation is not None:
                try:
                    axis.LabelRotation = int(x_axis_rotation)
                except Exception:
                    pass
            if axis_dimension == 1 and y_axis_rotation is not None:
                try:
                    axis.LabelRotation = int(y_axis_rotation)
                except Exception:
                    pass
            title_for_axis = (
                x_axis_title_text if axis_dimension == 0 else y_axis_title_text
            )
            title_applied = _set_axis_title(axis, title_for_axis)
            if axis_dimension == 0 and title_applied:
                x_axis_title_applied = True
            if axis_dimension == 1 and title_applied:
                y_axis_title_applied = True
            if not is_non_cartesian and axis_dimension == 1:
                _set_axis_number_format(chart_document, axis, y_axis_number_format)
        except Exception:
            continue

    # Axes and titles (classic fallback)
    if diagram is not None:
        if x_axis_title_text is not None and not x_axis_title_applied:
            try:
                has_x_title_prop = _has_prop(diagram, "HasXAxisTitle") or hasattr(
                    diagram, "HasXAxisTitle"
                )
                if has_x_title_prop:
                    try:
                        diagram.setPropertyValue(
                            "HasXAxisTitle", bool(x_axis_title_text)
                        )
                    except Exception:
                        diagram.HasXAxisTitle = bool(x_axis_title_text)
                if x_axis_title_text and hasattr(diagram, "XAxisTitle"):
                    axis_title_obj = diagram.XAxisTitle
                    if axis_title_obj is not None and hasattr(axis_title_obj, "String"):
                        axis_title_obj.String = x_axis_title_text
            except Exception:
                pass
        if y_axis_title_text is not None and not y_axis_title_applied:
            try:
                has_y_title_prop = _has_prop(diagram, "HasYAxisTitle") or hasattr(
                    diagram, "HasYAxisTitle"
                )
                if has_y_title_prop:
                    try:
                        diagram.setPropertyValue(
                            "HasYAxisTitle", bool(y_axis_title_text)
                        )
                    except Exception:
                        diagram.HasYAxisTitle = bool(y_axis_title_text)
                if y_axis_title_text and hasattr(diagram, "YAxisTitle"):
                    axis_title_obj = diagram.YAxisTitle
                    if axis_title_obj is not None and hasattr(axis_title_obj, "String"):
                        axis_title_obj.String = y_axis_title_text
            except Exception:
                pass

    # Gridlines toggle
    diagram_for_grids = diagram_for_axes or diagram
    if show_gridlines is not None:
        for grid_prop in (
            "HasXAxisGrid",
            "HasYAxisGrid",
            "HasSecondaryXAxisGrid",
            "HasSecondaryYAxisGrid",
        ):
            if _has_prop(diagram_for_grids, grid_prop):
                try:
                    diagram_for_grids.setPropertyValue(grid_prop, bool(show_gridlines))
                except Exception:
                    pass
    else:
        for grid_prop in (
            "HasXAxisGrid",
            "HasYAxisGrid",
            "HasSecondaryXAxisGrid",
            "HasSecondaryYAxisGrid",
        ):
            if _has_prop(diagram_for_grids, grid_prop):
                try:
                    diagram_for_grids.setPropertyValue(grid_prop, False)
                except Exception:
                    pass

    # Apply series styling (chart2 preferred; classic fallback if needed)
    _apply_series_styles(
        chart_document,
        chart_type_lower,
        show_data_labels,
        diagram_hint=(chart2_diagram or diagram),
    )
    _apply_combo_overrides(
        chart_document,
        payload or {},
        diagram_hint=(chart2_diagram or diagram),
    )


def _set_uno_property(obj, name: str, value: Any) -> bool:
    if obj is None:
        return False
    try:
        if _has_prop(obj, name):
            obj.setPropertyValue(name, value)
            return True
    except Exception:
        pass
    try:
        setattr(obj, name, value)
        return True
    except Exception:
        return False


def _apply_stack_mode(diagram, stack_mode: Optional[str]) -> None:
    if diagram is None or not stack_mode:
        return

    mode = str(stack_mode).strip().lower()
    if mode not in {"none", "stacked", "percent"}:
        return

    stacked = mode in {"stacked", "percent"}
    percent = mode == "percent"

    _set_uno_property(diagram, "Stacked", stacked)
    _set_uno_property(diagram, "Percent", percent)


def _apply_combo_overrides(
    chart_document,
    payload: Dict[str, Any],
    diagram_hint=None,
) -> None:
    chart_type = _normalize_chart_type(payload.get("chart_type") or "")
    if chart_type != "combo" or chart_document is None:
        return

    diagram = diagram_hint
    if diagram is None:
        try:
            diagram = chart_document.getDiagram()
        except Exception:
            diagram = None
    if diagram is None:
        return

    coordinate_systems = None
    try:
        getter = getattr(diagram, "getCoordinateSystems", None)
        if callable(getter):
            coordinate_systems = getter()
    except Exception:
        coordinate_systems = None
    if not coordinate_systems:
        return

    chart_types = None
    try:
        chart_types = coordinate_systems[0].getChartTypes()
    except Exception:
        chart_types = None
    if not chart_types:
        return

    try:
        data_series = chart_types[0].getDataSeries()
    except Exception:
        data_series = None
    if not data_series:
        return

    combo_config = payload.get("combo_config") if isinstance(payload.get("combo_config"), dict) else {}
    line_indices = set()
    secondary_indices = set()

    for field_name in ("line_series_indices",):
        values = combo_config.get(field_name)
        if isinstance(values, list):
            for value in values:
                if isinstance(value, int) and value >= 0:
                    line_indices.add(value)

    for field_name in ("secondary_series_indices",):
        values = combo_config.get(field_name)
        if isinstance(values, list):
            for value in values:
                if isinstance(value, int) and value >= 0:
                    secondary_indices.add(value)

    direct_secondary = payload.get("secondary_axis_series")
    if isinstance(direct_secondary, list):
        for value in direct_secondary:
            if isinstance(value, int) and value >= 0:
                secondary_indices.add(value)

    series_config = payload.get("series_config")
    if isinstance(series_config, list):
        for idx, series in enumerate(series_config):
            if not isinstance(series, dict):
                continue
            if series.get("use_secondary_axis") is True:
                secondary_indices.add(idx)
            if str(series.get("chart_type") or "").strip().lower() == "line":
                line_indices.add(idx)

    for idx, series in enumerate(data_series):
        if idx in secondary_indices:
            _set_uno_property(series, "AttachedAxisIndex", 1)
            _set_uno_property(series, "AxisIndex", 1)
        if idx in line_indices:
            # Approximate line rendering by emphasizing stroke and removing fill.
            _set_uno_property(series, "FillTransparence", 100)
            _set_uno_property(series, "LineWidth", 90)
            _set_uno_property(series, "LineStyle", 1)


def _set_axis_number_format(chart_document, axis, format_string: Optional[str] = None):
    """Set a readable currency/number format on the provided axis when possible."""
    if chart_document is None or axis is None:
        return
    try:
        formats_supplier = None
        if hasattr(chart_document, "getNumberFormats"):
            formats_supplier = chart_document.getNumberFormats()
        elif hasattr(chart_document, "NumberFormats"):
            formats_supplier = chart_document.NumberFormats
        if formats_supplier is None:
            return
        locale = uno.createUnoStruct("com.sun.star.lang.Locale")
        locale.Language = "en"
        locale.Country = "US"
        fmt_str = format_string or "$#,##0"
        try:
            fmt_key = formats_supplier.queryKey(fmt_str, locale, True)
        except Exception:
            fmt_key = 0
        if fmt_key == 0:
            try:
                fmt_key = formats_supplier.addNew(fmt_str, locale)
            except Exception:
                fmt_key = 0
        if fmt_key and _has_prop(axis, "NumberFormat"):
            try:
                axis.setPropertyValue("NumberFormat", fmt_key)
            except Exception:
                pass
    except Exception:
        pass


def _set_axis_title(axis, title_text: Optional[str]) -> bool:
    if axis is None or not title_text:
        return False
    try:
        getter = getattr(axis, "getTitle", None)
        title_obj = getter() if callable(getter) else None
    except Exception:
        title_obj = None
    if title_obj is not None and hasattr(title_obj, "String"):
        try:
            title_obj.String = title_text
            if hasattr(title_obj, "Show"):
                title_obj.Show = True
            return True
        except Exception as e:
            _log(f"_set_axis_title: failed to set '{title_text}': {e}")
    return False


def _set_chart2_series_label_struct(
    series,
    chart_type: str,
    show_data_labels: Optional[bool],
) -> bool:
    if series is None:
        return False
    has_label_prop = _has_prop(series, "Label") or hasattr(series, "Label")
    if not has_label_prop:
        return False
    try:
        if _has_prop(series, "Label"):
            label = series.getPropertyValue("Label")
        else:
            label = getattr(series, "Label", None)
    except Exception:
        label = None
    if label is None:
        try:
            label = uno.createUnoStruct("com.sun.star.chart2.DataPointLabel")
        except Exception:
            return False

    try:
        if show_data_labels is False:
            if hasattr(label, "ShowNumber"):
                label.ShowNumber = False
            if hasattr(label, "ShowNumberInPercent"):
                label.ShowNumberInPercent = False
            if hasattr(label, "ShowCategoryName"):
                label.ShowCategoryName = False
        elif "pie" in chart_type or "doughnut" in chart_type:
            if hasattr(label, "ShowNumber"):
                label.ShowNumber = True
            if hasattr(label, "ShowNumberInPercent"):
                label.ShowNumberInPercent = True
            if hasattr(label, "ShowCategoryName"):
                label.ShowCategoryName = True
        else:
            if hasattr(label, "ShowNumber"):
                label.ShowNumber = True
            if hasattr(label, "ShowNumberInPercent"):
                label.ShowNumberInPercent = False
            if hasattr(label, "ShowCategoryName"):
                label.ShowCategoryName = False

        if _has_prop(series, "Label"):
            series.setPropertyValue("Label", label)
        else:
            series.Label = label
        return True
    except Exception:
        return False


def _chart_data_caption_constant(name: str, fallback: int) -> int:
    try:
        return int(uno.getConstantByName(f"com.sun.star.chart.ChartDataCaption.{name}"))
    except Exception:
        return int(fallback)


def _apply_series_styles_classic(
    diagram,
    chart_type: str,
    show_data_labels: Optional[bool] = None,
):
    if diagram is None:
        return

    palette = [
        0x1F77B4,
        0xFF7F0E,
        0x2CA02C,
        0xD62728,
        0x9467BD,
        0x8C564B,
        0xE377C2,
        0x7F7F7F,
        0xBCBD22,
        0x17BECF,
    ]

    caption_none = _chart_data_caption_constant("NONE", 0)
    caption_value = _chart_data_caption_constant("VALUE", 1)
    caption_percent = _chart_data_caption_constant("PERCENT", 2)
    caption_text = _chart_data_caption_constant("TEXT", 4)

    if show_data_labels is False:
        label_caption = caption_none
    elif "pie" in chart_type or "doughnut" in chart_type:
        label_caption = caption_value | caption_percent | caption_text
    else:
        label_caption = caption_value

    try:
        if _has_prop(diagram, "DataCaption"):
            try:
                diagram.setPropertyValue("DataCaption", label_caption)
            except Exception:
                setattr(diagram, "DataCaption", label_caption)
        elif hasattr(diagram, "DataCaption"):
            setattr(diagram, "DataCaption", label_caption)
        try:
            _log(
                f"_apply_series_styles_classic: diagram DataCaption={getattr(diagram, 'DataCaption', None)}"
            )
        except Exception:
            pass
    except Exception:
        pass

    # Series colors + label captions.
    for series_index in range(0, 64):
        try:
            row_props = diagram.getDataRowProperties(series_index)
        except Exception:
            break
        color = palette[series_index % len(palette)]
        try:
            if hasattr(row_props, "FillColor"):
                row_props.FillColor = color
            if hasattr(row_props, "LineColor"):
                row_props.LineColor = color
        except Exception:
            pass
        caption_set = False
        if _has_prop(row_props, "DataCaption"):
            try:
                row_props.setPropertyValue("DataCaption", label_caption)
                caption_set = True
            except Exception:
                pass
        if not caption_set and hasattr(row_props, "DataCaption"):
            try:
                row_props.DataCaption = label_caption
                caption_set = True
            except Exception:
                pass
        if series_index == 0:
            try:
                _log(
                    f"_apply_series_styles_classic: row0 DataCaption={getattr(row_props, 'DataCaption', None)} set_ok={caption_set}"
                )
            except Exception:
                pass

    if (
        "pie" in chart_type
        or "doughnut" in chart_type
        or chart_type in {"bar", "bar3d", "column", "column3d"}
    ):
        try:
            if hasattr(diagram, "VaryColorsByPoint"):
                diagram.VaryColorsByPoint = True
        except Exception:
            pass

        # Per-point styling fallback for classic API.
        for point_index in range(0, 256):
            try:
                point_props = diagram.getDataPointProperties(0, point_index)
            except Exception:
                break
            color = palette[point_index % len(palette)]
            try:
                if hasattr(point_props, "FillColor"):
                    point_props.FillColor = color
            except Exception:
                pass
            if "pie" in chart_type or "doughnut" in chart_type:
                try:
                    if hasattr(point_props, "DataCaption"):
                        point_props.DataCaption = label_caption
                except Exception:
                    pass


def _apply_series_styles(
    chart_document,
    chart_type: str,
    show_data_labels: Optional[bool] = None,
    diagram_hint=None,
):
    """Apply base styling using chart2 properties and enable labels."""
    try:
        diagram = diagram_hint if diagram_hint is not None else chart_document.getDiagram()
        if diagram is None:
            return
        get_coordinate_systems = getattr(diagram, "getCoordinateSystems", None)
        coordinate_systems = (
            get_coordinate_systems() if callable(get_coordinate_systems) else None
        )
    except Exception:
        coordinate_systems = None
        diagram = diagram_hint

    if not coordinate_systems:
        _apply_series_styles_classic(diagram, chart_type, show_data_labels)
        return
    try:
        chart_types = coordinate_systems[0].getChartTypes()
    except Exception:
        chart_types = None
    if not chart_types:
        _apply_series_styles_classic(diagram, chart_type, show_data_labels)
        return
    try:
        data_series = chart_types[0].getDataSeries()
    except Exception:
        data_series = None
    if not data_series:
        _apply_series_styles_classic(diagram, chart_type, show_data_labels)
        return

    palette = [
        0x1F77B4,
        0xFF7F0E,
        0x2CA02C,
        0xD62728,
        0x9467BD,
        0x8C564B,
        0xE377C2,
        0x7F7F7F,
        0xBCBD22,
        0x17BECF,
    ]

    for series_index, series in enumerate(data_series):
        base_color = palette[series_index % len(palette)]
        try:
            if hasattr(series, "FillStyle"):
                series.FillStyle = 1  # solid fill
            if hasattr(series, "FillColor"):
                series.FillColor = base_color
            if hasattr(series, "LineStyle"):
                series.LineStyle = 1  # solid line
            if hasattr(series, "LineColor"):
                series.LineColor = base_color
            if hasattr(series, "LineWidth"):
                series.LineWidth = 50  # visible line width (1/100 mm)
        except Exception as e:
            _log(f"_apply_series_styles: series style error: {e}")

        # Configure data labels.
        try:
            get_labels = getattr(series, "getDataLabels", None)
            labels = get_labels() if callable(get_labels) else None
        except Exception:
            labels = None

        _set_chart2_series_label_struct(series, chart_type, show_data_labels)

        if labels is not None and show_data_labels is not False:
            try:
                if "pie" in chart_type or "doughnut" in chart_type:
                    if hasattr(labels, "ShowNumber"):
                        labels.ShowNumber = True
                    if hasattr(labels, "ShowNumberInPercent"):
                        labels.ShowNumberInPercent = True
                    if hasattr(labels, "ShowCategoryName"):
                        labels.ShowCategoryName = True
                else:
                    if hasattr(labels, "ShowNumber"):
                        labels.ShowNumber = True
                    if hasattr(labels, "ShowCategoryName"):
                        labels.ShowCategoryName = False
            except Exception:
                pass
        elif labels is not None and show_data_labels is False:
            try:
                if hasattr(labels, "ShowNumber"):
                    labels.ShowNumber = False
                if hasattr(labels, "ShowNumberInPercent"):
                    labels.ShowNumberInPercent = False
                if hasattr(labels, "ShowCategoryName"):
                    labels.ShowCategoryName = False
            except Exception:
                pass

        # Per-point colors for pie/doughnut and single-series bars/columns.
        if (
            "pie" in chart_type
            or "doughnut" in chart_type
            or (
                chart_type in {"bar", "bar3d", "column", "column3d"}
                and len(data_series) == 1
            )
        ):
            try:
                if hasattr(diagram, "VaryColorsByPoint"):
                    diagram.VaryColorsByPoint = True
                points = series.getDataPoints()
            except Exception:
                points = None
            if points:
                for idx, point in enumerate(points):
                    try:
                        color = palette[idx % len(palette)]
                        if hasattr(point, "FillColor"):
                            point.FillColor = color
                    except Exception:
                        continue


def _handle_set_freeze_panes(document, controller, action: Dict[str, Any]):
    """Freeze rows/columns at a specific cell position so the title row,
    subtitle/description row, and column-header row stay visible while the
    user scrolls the data area.

    Three details that the previous implementation got wrong:
      1. ``freezeAtPosition`` operates on the controller's *active* sheet —
         so we set the target sheet active first.  Without this, freezes
         could land on the wrong sheet in multi-sheet workbooks.
      2. ``lockControllers()`` (set globally for the plan) suppresses the
         view-state broadcast that Collabora's client needs to render the
         freeze.  Briefly unlocking around the call lets the broadcast
         reach the renderer (same pattern as the merge fix).
      3. The freeze is view state, not cell content.  After applying it, we
         refresh the controller's view-data payload and mark the document
         modified so the existing host-managed save/WOPI path serializes it.
    """
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    freeze_at = payload.get("freeze_at") or target_info.get("range") or "A2"
    sheet, freeze_at = _resolve_sheet_and_range_name(
        document, sheet_name, freeze_at, "set_freeze_panes"
    )
    if sheet is None:
        raise RuntimeError(f"set_freeze_panes sheet '{sheet_name}' not found")

    cell = sheet.getCellRangeByName(freeze_at)
    skip_entry = _find_duplicate_header_freeze_skip(sheet, cell)
    action_name = str(action.get("action_name") or "")
    if skip_entry and action_name == "Freeze Panes Fallback":
        meta = action.setdefault("_result_meta", {})
        meta["skipped_due_to_duplicate_header_write"] = True
        meta["requested_freeze_at"] = freeze_at
        meta["skipped_header_range_a1"] = skip_entry.get("skipped_header_range_a1")
        meta["existing_header_range_a1"] = skip_entry.get("existing_header_range_a1")
        _log(
            "_handle_set_freeze_panes: skipped Freeze Panes Fallback generated "
            f"for duplicate WRITE_DATA header row {skip_entry.get('skipped_header_range_a1')}"
        )
        return cell

    address = cell.getCellAddress()
    col, row = address.Column, address.Row

    # 1. Make the target sheet active so freezeAtPosition lands here.
    try:
        controller.setActiveSheet(sheet)
    except Exception as exc:
        _log(f"_handle_set_freeze_panes: setActiveSheet failed: {exc}")
    try:
        controller.select(cell)
    except Exception as exc:
        _log(f"_handle_set_freeze_panes: select freeze cell failed: {exc}")

    # 2. Briefly unlock controllers so the freeze actually broadcasts.
    try:
        document.unlockControllers()
    except Exception:
        pass
    freeze_ok = True
    try:
        controller.freezeAtPosition(col, row)
    except Exception as exc:
        freeze_ok = False
        _log(f"_handle_set_freeze_panes: freezeAtPosition raised {exc}")
    try:
        document.lockControllers()
    except Exception:
        pass

    view_data_ok = False
    if freeze_ok:
        try:
            view_data = controller.getViewData()
            if view_data:
                controller.restoreViewData(view_data)
                view_data_ok = True
        except Exception as exc:
            _log(f"_handle_set_freeze_panes: view-data refresh failed: {exc}")
        try:
            document.setModified(True)
        except Exception as exc:
            _log(f"_handle_set_freeze_panes: setModified failed: {exc}")

    _log(
        f"_handle_set_freeze_panes: froze at {freeze_at} "
        f"(col={col}, row={row}) ok={freeze_ok} view_data_ok={view_data_ok}"
    )
    return cell


def _handle_add_data_validation(document, controller, action: Dict[str, Any]):
    """Add data validation to a cell range."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    target_range = target_info.get("range") or payload.get("range") or payload.get("target_range")
    validation_type = payload.get("validation_type", "list")
    formula1 = payload.get("formula1")
    if formula1 in (None, ""):
        formula1 = payload.get("values") or ""
    formula2 = payload.get("formula2") or ""
    show_dropdown = payload.get("show_dropdown", True)
    error_message = payload.get("error_message") or ""
    error_title = payload.get("error_title") or "Validation Error"

    sheet, target_range = _resolve_sheet_and_range_name(
        document, sheet_name, target_range, "add_data_validation"
    )
    if sheet is None:
        raise RuntimeError(f"add_data_validation sheet '{sheet_name}' not found")

    if not target_range:
        raise RuntimeError("add_data_validation requires range")

    cell_range = sheet.getCellRangeByName(target_range)

    # Get validation object
    validation = cell_range.getPropertyValue("Validation")

    # Map validation types
    type_map = {
        "list": 6,  # com.sun.star.sheet.ValidationType.LIST
        "whole": 1,  # WHOLE
        "decimal": 2,  # DECIMAL
        "date": 3,  # DATE
        "text_length": 7,  # TEXT_LEN
    }
    val_type = type_map.get(validation_type, 6)

    validation.Type = val_type
    validation.ShowInputMessage = True
    validation.ShowErrorMessage = True
    validation.ErrorMessage = error_message
    validation.ErrorTitle = error_title

    if validation_type == "list":
        formula1 = _normalize_list_validation_formula(sheet, cell_range, formula1)
        validation.setFormula1(formula1)
        for attr_name, attr_value in (
            ("ShowList", bool(show_dropdown)),
            ("IgnoreBlankCells", bool(payload.get("allow_blank", True))),
        ):
            try:
                setattr(validation, attr_name, attr_value)
            except Exception:
                pass
    else:
        validation.setFormula1(formula1)
        if formula2:
            validation.setFormula2(formula2)

    cell_range.setPropertyValue("Validation", validation)
    try:
        document.setModified(True)
    except Exception:
        pass
    _log(
        f"_handle_add_data_validation: added validation to {target_range} "
        f"type={validation_type} formula1={formula1!r}"
    )
    return cell_range


def _normalize_list_validation_formula(sheet, cell_range, raw_formula: Any) -> str:
    """Return a durable, display-cased inline list formula for Calc/XLSX.

    Agents often pass inline list values as lowercase comma-separated text
    (``hr,finance,it``).  If the target cells already contain display-cased
    values (``HR``, ``Finance``, ``IT``), use those spellings for the dropdown.
    XLSX inline lists should be quoted as one formula string, e.g.
    ``"HR,Finance,IT"``.
    """
    if isinstance(raw_formula, (list, tuple)):
        raw_text = ",".join(str(item) for item in raw_formula if item is not None)
    else:
        raw_text = str(raw_formula or "").strip()

    if _looks_like_validation_reference(raw_text):
        return raw_text

    items = _split_inline_validation_items(raw_text)
    if not items:
        items = _existing_validation_values(sheet, cell_range)
    if not items:
        return raw_text

    display_by_key: Dict[str, str] = {}
    for display_value in _existing_validation_values(sheet, cell_range):
        key = display_value.strip().casefold()
        if key and key not in display_by_key:
            display_by_key[key] = display_value.strip()

    normalized: List[str] = []
    seen: set = set()
    for item in items:
        text = str(item).strip()
        if not text:
            continue
        display_text = display_by_key.get(text.casefold(), text)
        key = display_text.casefold()
        if key in seen:
            continue
        seen.add(key)
        normalized.append(display_text)

    escaped = ",".join(value.replace('"', '""') for value in normalized)
    return f'"{escaped}"'


def _looks_like_validation_reference(text: str) -> bool:
    if not text:
        return False
    stripped = text.strip()
    if stripped.startswith("="):
        return True
    unquoted = _strip_outer_quotes(stripped)
    if "," in unquoted or ";" in unquoted:
        return False
    return bool(re.search(r"[$]?[A-Za-z]{1,3}[$]?\d+[:.][$]?[A-Za-z]{1,3}[$]?\d+", unquoted))


def _split_inline_validation_items(text: str) -> List[str]:
    if not text:
        return []
    inner = _strip_outer_quotes(text.strip())
    if not inner:
        return []
    delimiter = ";" if ";" in inner and "," not in inner else ","
    return [part.strip().replace('""', '"') for part in inner.split(delimiter)]


def _strip_outer_quotes(text: str) -> str:
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        return text[1:-1]
    return text


def _existing_validation_values(sheet, cell_range) -> List[str]:
    values: List[str] = []
    seen: set = set()
    for cell in _iterate_cells(sheet, cell_range):
        try:
            value = str(cell.getString() or "").strip()
        except Exception:
            value = ""
        if not value:
            continue
        key = value.casefold()
        if key in seen:
            continue
        seen.add(key)
        values.append(value)
    return values


def _handle_protect_sheet(document, controller, action: Dict[str, Any]):
    """Protect a sheet from editing."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    password = payload.get("password") or ""

    sheet = _get_sheet(document, sheet_name)
    if sheet is None:
        raise RuntimeError(f"protect_sheet sheet '{sheet_name}' not found")

    sheet.protect(password)
    _log(f"_handle_protect_sheet: protected sheet '{sheet_name}'")
    return sheet.getCellRangeByName("A1")


def _handle_insert_image(document, controller, action: Dict[str, Any]):
    """Insert an image into a sheet."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    image_url = payload.get("image_url") or payload.get("image_path")
    anchor_cell = payload.get("anchor_cell") or target_info.get("range") or "A1"
    image_ref = payload.get("image_ref")

    if not image_url:
        raise RuntimeError("insert_image requires image_url")

    sheet, anchor_cell = _resolve_sheet_and_range_name(
        document, sheet_name, anchor_cell, "insert_image"
    )
    if sheet is None:
        raise RuntimeError(f"insert_image sheet '{sheet_name}' not found")

    native_placement = payload.get("placement")
    if isinstance(native_placement, dict):
        image_rect, image_area_range = _image_rect_and_area_from_placement(
            sheet,
            native_placement,
        )
        collision_details = _image_rect_collision_details(
            sheet,
            image_rect,
            requested_placement=native_placement,
            requested_area=image_area_range,
        )
        if collision_details:
            raise ActionApplicationError(
                "insert_image placement overlaps existing sheet content or another drawing",
                error_code="IMAGE_PLACEMENT_COLLISION",
                details=collision_details,
            )
        anchor_for_shape = _cell_ref_from_chart_rect(sheet, image_rect)
    else:
        width = _image_dimension_to_mm100(payload.get("width"), 10000)
        height = _image_dimension_to_mm100(payload.get("height"), 8000)
        image_rect, image_area_range, anchor_for_shape = _resolve_legacy_image_rect(
            sheet,
            anchor_cell,
            width,
            height,
        )

    # Get draw page
    draw_page = sheet.getDrawPage()

    # Create graphic shape
    shape = document.createInstance("com.sun.star.drawing.GraphicObjectShape")

    # Set the image URL. Calc is most reliable when external images are first
    # downloaded and assigned via a local file URI.
    shape.GraphicURL = _resolve_image_graphic_url(str(image_url))
    name_seed = image_ref or f"{sheet_name}_{anchor_cell}"
    shape.Name = re.sub(r"[^A-Za-z0-9_]+", "_", str(name_seed)).strip("_")[:120] or "image"

    try:
        if anchor_for_shape is not None:
            shape.Anchor = anchor_for_shape
    except Exception:
        pass
    shape.setPosition(
        uno.createUnoStruct(
            "com.sun.star.awt.Point",
            int(getattr(image_rect, "X", 0) or 0),
            int(getattr(image_rect, "Y", 0) or 0),
        )
    )
    shape.setSize(
        uno.createUnoStruct(
            "com.sun.star.awt.Size",
            int(getattr(image_rect, "Width", 0) or 0),
            int(getattr(image_rect, "Height", 0) or 0),
        )
    )

    draw_page.add(shape)
    _log(
        "_handle_insert_image: inserted image placement="
        f"{_placement_from_rect(image_rect)} occupied={_range_diagnostics(image_area_range).get('range_a1')}"
    )
    return image_area_range or anchor_for_shape


def _image_dimension_to_mm100(value: Any, default: int) -> int:
    """Convert schema pixel dimensions to LibreOffice 1/100mm units.

    Older callers may have sent native 1/100mm values, so values above 2000 are
    treated as already converted. Schema-sized values such as 180x120 are
    interpreted as pixels at 96 DPI.
    """
    if value is None:
        return int(default)
    try:
        numeric = float(value)
    except Exception:
        return int(default)
    if numeric <= 0:
        return int(default)
    if numeric <= 2000:
        return max(100, int(round(numeric * 2540 / 96)))
    return int(round(numeric))


def _is_unvalidated_external_image_url(raw: str) -> bool:
    parsed = urlparse(str(raw or "").strip())
    if parsed.scheme not in {"http", "https"}:
        return False
    netloc = (parsed.netloc or "").lower()
    if netloc.startswith(("127.0.0.1", "localhost")):
        return False
    if netloc in {"storage.googleapis.com", "firebasestorage.googleapis.com"}:
        return False
    if "/api/files/" in parsed.path:
        return False
    return True


def _resolve_image_graphic_url(image_url: str) -> str:
    raw = str(image_url or "").strip()
    if not raw:
        raise RuntimeError("insert_image requires image_url")
    parsed = urlparse(raw)
    if parsed.scheme in {"http", "https"}:
        if _is_unvalidated_external_image_url(raw):
            raise RuntimeError(
                "insert_image image_url is an unvalidated external URL; "
                "validate and save it with web_fetch_image_tool first"
            )
        suffix = os.path.splitext(parsed.path or "")[1] or ".png"
        if len(suffix) > 8:
            suffix = ".png"
        headers = {"User-Agent": "Mozilla/5.0"}
        request = urllib.request.Request(raw, headers=headers)
        with urllib.request.urlopen(
            request,
            timeout=30,
            context=_UNVERIFIED_SSL_CONTEXT,
        ) as response:
            content = response.read()
        if not content:
            raise RuntimeError(f"insert_image downloaded empty content from {raw}")
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp:
            temp.write(content)
            temp_path = temp.name
        return Path(temp_path).resolve().as_uri()
    if parsed.scheme == "file":
        return raw
    if os.path.exists(raw):
        return Path(raw).resolve().as_uri()
    return raw


def _find_graphic_shape_on_sheet(
    sheet, image_ref: Optional[str] = None, anchor_cell: Optional[str] = None
):
    """Locate a graphic shape by name or top-left anchored cell position."""
    draw_page = sheet.getDrawPage()
    target_pos = None
    normalized_names = []

    if image_ref:
        normalized_names.append(str(image_ref))
        normalized = re.sub(r"[^A-Za-z0-9_]+", "_", str(image_ref)).strip("_")
        if normalized and normalized not in normalized_names:
            normalized_names.append(normalized)

    if anchor_cell:
        try:
            target_rect = _get_cell_range_rect(sheet, sheet.getCellRangeByName(anchor_cell))
            if target_rect:
                target_pos = uno.createUnoStruct(
                    "com.sun.star.awt.Point", target_rect["X"], target_rect["Y"]
                )
        except Exception as exc:
            _log(
                f"_find_graphic_shape_on_sheet: failed to resolve anchor cell '{anchor_cell}': {exc}"
            )

    count = draw_page.getCount()
    for idx in range(count):
        shape = draw_page.getByIndex(idx)
        try:
            if not shape.supportsService("com.sun.star.drawing.GraphicObjectShape"):
                continue
        except Exception:
            continue

        try:
            shape_name = getattr(shape, "Name", None)
            if normalized_names and shape_name in normalized_names:
                return shape
        except Exception:
            pass

        if target_pos is not None:
            try:
                pos = shape.getPosition()
                if pos.X == target_pos.X and pos.Y == target_pos.Y:
                    return shape
            except Exception:
                continue
    return None


def _handle_delete_image(document, controller, action: Dict[str, Any]):
    """Delete an existing image from a sheet."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    image_ref = payload.get("image_ref")
    target_range = (
        payload.get("range")
        or payload.get("target_range")
        or payload.get("anchor_cell")
        or target_info.get("range")
    )

    if not image_ref and not target_range:
        raise RuntimeError("delete_image requires image_ref or target_range")

    sheet, target_range = _resolve_sheet_and_range_name(
        document, sheet_name, target_range or "A1", "delete_image"
    )
    if sheet is None:
        raise RuntimeError(f"delete_image sheet '{sheet_name}' not found")

    shape = _find_graphic_shape_on_sheet(
        sheet, image_ref=image_ref, anchor_cell=target_range
    )
    if shape is None:
        raise RuntimeError(
            f"delete_image could not locate image_ref='{image_ref}' target_range='{target_range}'"
        )

    shape.dispose()
    _log(
        "_handle_delete_image: deleted image "
        f"image_ref={image_ref} target_range={target_range}"
    )
    return sheet.getCellRangeByName(target_range or "A1")


def _handle_add_hyperlink(document, controller, action: Dict[str, Any]):
    """Add a hyperlink to a cell."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    target_cell = payload.get("target_cell") or target_info.get("range")
    url = payload.get("url")
    display_text = payload.get("display_text")

    if not target_cell or not url:
        raise RuntimeError("add_hyperlink requires target_cell and url")

    sheet, target_cell = _resolve_sheet_and_range_name(
        document, sheet_name, target_cell, "add_hyperlink"
    )
    if sheet is None:
        raise RuntimeError(f"add_hyperlink sheet '{sheet_name}' not found")

    cell = sheet.getCellRangeByName(target_cell)

    # Set the text
    text = display_text or url
    cell.setString(text)

    # Create hyperlink via text field
    text_obj = cell.getText()
    cursor = text_obj.createTextCursor()
    cursor.gotoStart(False)
    cursor.gotoEnd(True)
    cursor.setPropertyValue("HyperLinkURL", url)
    cursor.setPropertyValue("HyperLinkTarget", "_blank")

    _log(f"_handle_add_hyperlink: added hyperlink to {target_cell}")
    return cell


def _safe_excel_table_name(raw_name: Any, prefix: str = "AutoFilter") -> str:
    raw_text = str(raw_name or "")
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", raw_text)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned:
        cleaned = prefix
    if not re.match(r"^[A-Za-z_]", cleaned):
        cleaned = f"{prefix}_{cleaned}"
    suffix = hashlib.sha1(raw_text.encode("utf-8", errors="ignore")).hexdigest()[:8]
    max_base_len = 240 - len(suffix) - 1
    return f"{cleaned[:max_base_len]}_{suffix}"


def _handle_apply_filter(document, controller, action: Dict[str, Any]):
    """Apply auto-filter to a data range."""
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    target_range = target_info.get("range") or payload.get("range") or payload.get("target_range")

    if not target_range:
        raise RuntimeError("apply_filter requires range")

    sheet, target_range = _resolve_sheet_and_range_name(
        document, sheet_name, target_range, "apply_filter"
    )
    if sheet is None:
        raise RuntimeError(f"apply_filter sheet '{sheet_name}' not found")

    cell_range = sheet.getCellRangeByName(target_range)

    # Create database range for auto-filter
    db_ranges = document.DatabaseRanges
    db_name = _safe_excel_table_name(
        f"AutoFilter_{_sheet_name_for_state(sheet, sheet_name)}_{target_range}"
    )

    if db_ranges.hasByName(db_name):
        db_ranges.removeByName(db_name)

    address = cell_range.getRangeAddress()
    db_ranges.addNewByName(db_name, address)

    db_range = db_ranges.getByName(db_name)
    db_range.AutoFilter = True

    _log(f"_handle_apply_filter: applied auto-filter to {target_range}")
    meta = action.setdefault("_result_meta", {})
    meta.update(
        {
            "filter_range": target_range,
            "database_range_name": db_name,
        }
    )
    return cell_range


def _handle_hide_rows(document, controller, action: Dict[str, Any], hidden: bool):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("hide_rows/show_rows requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "hide_rows"
    )
    if sheet is None:
        raise RuntimeError(f"hide_rows/show_rows sheet '{sheet_name}' not found")
    cell_range = sheet.getCellRangeByName(range_name)
    address = cell_range.getRangeAddress()
    rows = sheet.getRows()
    for row_idx in range(address.StartRow, address.EndRow + 1):
        row = rows.getByIndex(row_idx)
        row.IsVisible = not hidden
    _log(
        f"_handle_hide_rows: {'hidden' if hidden else 'shown'} rows {range_name} on '{sheet_name}'"
    )
    return cell_range


def _handle_hide_columns(document, controller, action: Dict[str, Any], hidden: bool):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    if not range_name:
        raise RuntimeError("hide_columns/show_columns requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "hide_columns"
    )
    if sheet is None:
        raise RuntimeError(f"hide_columns/show_columns sheet '{sheet_name}' not found")
    cell_range = sheet.getCellRangeByName(range_name)
    address = cell_range.getRangeAddress()
    columns = sheet.getColumns()
    for col_idx in range(address.StartColumn, address.EndColumn + 1):
        col = columns.getByIndex(col_idx)
        col.IsVisible = not hidden
    _log(
        f"_handle_hide_columns: {'hidden' if hidden else 'shown'} columns {range_name} on '{sheet_name}'"
    )
    return cell_range


def _handle_hide_rows_action(document, controller, action: Dict[str, Any]):
    return _handle_hide_rows(document, controller, action, True)


def _handle_show_rows_action(document, controller, action: Dict[str, Any]):
    return _handle_hide_rows(document, controller, action, False)


def _handle_hide_columns_action(document, controller, action: Dict[str, Any]):
    return _handle_hide_columns(document, controller, action, True)


def _handle_show_columns_action(document, controller, action: Dict[str, Any]):
    return _handle_hide_columns(document, controller, action, False)


def _handle_set_row_height(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    height = payload.get("height")
    if range_name is None or height is None:
        raise RuntimeError("set_row_height requires range and height")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "set_row_height"
    )
    if sheet is None:
        raise RuntimeError(f"set_row_height sheet '{sheet_name}' not found")
    cell_range = sheet.getCellRangeByName(range_name)
    address = cell_range.getRangeAddress()
    rows = sheet.getRows()
    mm100 = _points_to_mm100(float(height))
    for row_idx in range(address.StartRow, address.EndRow + 1):
        row = rows.getByIndex(row_idx)
        row.Height = mm100
    _log(f"_handle_set_row_height: set {height}pt for {range_name} on '{sheet_name}'")
    return cell_range


def _handle_set_column_width(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    width = payload.get("width")
    if range_name is None or width is None:
        raise RuntimeError("set_column_width requires range and width")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "set_column_width"
    )
    if sheet is None:
        raise RuntimeError(f"set_column_width sheet '{sheet_name}' not found")
    cell_range = sheet.getCellRangeByName(range_name)
    address = cell_range.getRangeAddress()
    columns = sheet.getColumns()
    mm100 = _chars_to_mm100(float(width))
    for col_idx in range(address.StartColumn, address.EndColumn + 1):
        col = columns.getByIndex(col_idx)
        col.Width = mm100
    _log(
        f"_handle_set_column_width: set {width}ch for {range_name} on '{sheet_name}'"
    )
    return cell_range


def _handle_merge_cells(document, controller, action: Dict[str, Any]):
    payload = action.get("payload") or {}
    target_info = action.get("target") or {}
    sheet_name = payload.get("sheet_name") or target_info.get("sheet_name")
    range_name = target_info.get("range") or payload.get("range") or payload.get("target_range")
    merge = payload.get("merge", True)
    if not range_name:
        raise RuntimeError("merge_cells requires range")
    sheet, range_name = _resolve_sheet_and_range_name(
        document, sheet_name, range_name, "merge_cells"
    )
    if sheet is None:
        raise RuntimeError(f"merge_cells sheet '{sheet_name}' not found")
    cell_range = sheet.getCellRangeByName(range_name)
    cell_range.merge(bool(merge))
    _log(
        f"_handle_merge_cells: {'merged' if merge else 'unmerged'} {range_name} on '{sheet_name}'"
    )
    return cell_range


def _get_cell_range_rect(sheet, cell_range):
    """
    Return the rectangle for a cell range in 1/100th mm units.

    We derive this from column widths and row heights instead of relying on
    getPosition()/getSize(), which are not reliably exposed on cell objects
    in all Calc hosts and can force us into a fallback rectangle.
    """
    if sheet is None or cell_range is None:
        return None

    try:
        addr = cell_range.getRangeAddress()
    except Exception:
        return None

    try:
        columns = sheet.getColumns()
        rows = sheet.getRows()
    except Exception:
        return None

    try:
        col_count = columns.getCount()
    except Exception:
        col_count = 0
    try:
        row_count = rows.getCount()
    except Exception:
        row_count = 0

    start_col = max(0, getattr(addr, "StartColumn", 0))
    end_col = max(start_col, getattr(addr, "EndColumn", start_col))
    start_row = max(0, getattr(addr, "StartRow", 0))
    end_row = max(start_row, getattr(addr, "EndRow", start_row))

    if col_count <= 0 or row_count <= 0:
        return None

    # Clamp indices to existing rows/columns to avoid index errors on huge
    # EndColumn/EndRow values.
    end_col = min(end_col, col_count - 1)
    end_row = min(end_row, row_count - 1)
    start_col = min(start_col, end_col)
    start_row = min(start_row, end_row)

    def _safe_col_width(index: int) -> int:
        try:
            return max(0, int(columns.getByIndex(index).Width))
        except Exception:
            return 0

    def _safe_row_height(index: int) -> int:
        try:
            return max(0, int(rows.getByIndex(index).Height))
        except Exception:
            return 0

    # Position is the sum of all preceding column widths / row heights.
    x_pos = sum(_safe_col_width(i) for i in range(start_col))
    y_pos = sum(_safe_row_height(i) for i in range(start_row))

    width = sum(_safe_col_width(i) for i in range(start_col, end_col + 1))
    height = sum(_safe_row_height(i) for i in range(start_row, end_row + 1))

    # Guard against degenerate rectangles which can confuse chart placement.
    min_size = 1000  # 10 mm
    width = max(min_size, width)
    height = max(min_size, height)

    return {
        "X": int(x_pos),
        "Y": int(y_pos),
        "Width": int(width),
        "Height": int(height),
    }


_ACTION_HANDLERS: Dict[str, Any] = {
    "set_active_sheet": _handle_set_active_sheet,
    "create_sheet": _handle_create_sheet,
    "rename_sheet": _handle_rename_sheet,
    "delete_sheet": _handle_delete_sheet,
    "write_range": _handle_write_range,
    "set_cell_formula": _handle_set_cell_formula,
    "append_rows": _handle_append_rows,
    "create_column": _handle_create_column,
    "delete_columns": _handle_delete_columns,
    "create_conditional_column": _handle_create_conditional_column,
    "delete_rows": _handle_delete_rows,
    "find_and_replace": _handle_find_and_replace,
    "manipulate_text_column": _handle_manipulate_text_column,
    "update_cells": _handle_update_cells,
    "apply_formatting": _handle_apply_formatting,
    "apply_conditional_formatting": _handle_apply_conditional_formatting,
    "autofit_columns": _handle_autofit_columns,
    "sort_range": _handle_sort_range,
    "create_chart": _handle_create_chart,
    "delete_chart": _handle_delete_chart,
    "add_title": _handle_add_title,
    "insert_comment": _handle_insert_comment,
    # New operations
    "set_freeze_panes": _handle_set_freeze_panes,
    "add_data_validation": _handle_add_data_validation,
    "protect_sheet": _handle_protect_sheet,
    "insert_image": _handle_insert_image,
    "delete_image": _handle_delete_image,
    "add_hyperlink": _handle_add_hyperlink,
    "apply_filter": _handle_apply_filter,
    "hide_rows": _handle_hide_rows_action,
    "show_rows": _handle_show_rows_action,
    "hide_columns": _handle_hide_columns_action,
    "show_columns": _handle_show_columns_action,
    "set_row_height": _handle_set_row_height,
    "set_column_width": _handle_set_column_width,
    "merge_cells": _handle_merge_cells,
}


_FORMULA_RECALC_PENDING = False
_CELL_DELETE_MODE_UP = None
_CELL_DELETE_MODE_LEFT = None
_CELL_INSERT_MODE_RIGHT = None


g_exportedScripts = (applyExcelActionPlan,)
