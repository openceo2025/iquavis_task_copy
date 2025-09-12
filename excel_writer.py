import json
import os
import re
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple


def _is_primitive(x: Any) -> bool:
    return x is None or isinstance(x, (str, int, float, bool))


def flatten_dict(obj: Dict[str, Any], parent_key: str = "", sep: str = ".") -> Dict[str, Any]:
    """
    Flatten nested dicts using dot notation.
    - Lists/tuples are serialized as compact JSON strings to preserve information without column explosion.
    - None stays None; consumers may render as blank.
    """
    items: List[Tuple[str, Any]] = []
    for k, v in (obj or {}).items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else str(k)
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        elif isinstance(v, (list, tuple)):
            try:
                items.append((new_key, json.dumps(v, ensure_ascii=False)))
            except Exception:
                # Fallback to str if not JSON-serializable
                items.append((new_key, str(v)))
        else:
            items.append((new_key, v))
    return dict(items)


def collect_headers(rows: Iterable[Dict[str, Any]], extra_headers: Iterable[str] = ()) -> List[str]:
    """
    Build the union of keys from flattened rows. Returns an ordered list:
    - Start with a preferred key order, then append remaining keys sorted.
    """
    preferred = [
        "Id",
        "Name",
        "Type",
        "StartDate",
        "EndDate",
        "ProjectId",
        "TaskDomainId",
    ]
    seen: Set[str] = set(extra_headers or [])
    for row in rows:
        seen.update(row.keys())

    ordered: List[str] = []
    for key in preferred:
        if key in seen:
            ordered.append(key)
            seen.remove(key)
    # Append remaining keys in stable sorted order for predictability
    ordered.extend(sorted(seen))
    return ordered


def sanitize_filename(name: str) -> str:
    # Remove/replace characters invalid on common filesystems
    name = name.strip()
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)  # Windows-forbidden
    # Avoid trailing dots/spaces on Windows
    name = name.rstrip(" .")
    return name or "project"


def next_available_path(base_dir: str, base_name: str) -> str:
    """
    If file exists, add suffix ' (n)'. Returns an available absolute path.
    """
    root, ext = os.path.splitext(base_name)
    candidate = os.path.join(base_dir, f"{root}{ext}")
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(base_dir, f"{root} ({n}){ext}")
        n += 1
    return candidate


def write_tasks_xlsx(
    tasks: List[Dict[str, Any]],
    project_name: str,
    out_dir: str,
    extra_headers: Iterable[str] = (),
    project_sheet_rows: Optional[Iterable[Iterable[Any]]] = None,
) -> str:
    """
    Write tasks to an .xlsx file with a header containing the union of all
    flattened keys. A "project" sheet is created as the left-most sheet using
    ``project_sheet_rows`` if provided. Returns the absolute path to the written
    file.
    """
    try:
        from openpyxl import Workbook
    except Exception as e:
        raise RuntimeError("openpyxl is required to export .xlsx. Please install via 'pip install openpyxl'.") from e

    # First flatten rows; keep in memory for simplicity and consistent headers
    flat_rows: List[Dict[str, Any]] = [flatten_dict(t) for t in tasks]
    headers = collect_headers(flat_rows, extra_headers=extra_headers)

    wb = Workbook()
    ws_project = wb.active
    ws_project.title = "project"

    for row in project_sheet_rows or []:
        ws_project.append(list(row))

    ws_tasks = wb.create_sheet("tasks")
    ws_tasks.append(headers)

    # Rows
    for row in flat_rows:
        ws_tasks.append([row.get(h) for h in headers])

    safe_name = sanitize_filename(project_name)
    file_name = f"tasks_{safe_name}.xlsx"
    out_path = next_available_path(out_dir, file_name)

    wb.save(out_path)
    return out_path
