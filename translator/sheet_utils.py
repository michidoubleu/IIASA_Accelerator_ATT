"""Generic helpers for locating header rows/markers in loosely-structured worksheets."""


def find_header_row(ws, sequence, max_row=None):
    """Find a row containing `sequence` as consecutive, case-insensitive cell values.

    Returns (row, first_col).
    """
    max_row = max_row or ws.max_row
    seq_norm = [s.strip().lower() for s in sequence]
    n = len(seq_norm)
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column - n + 2):
            vals = [ws.cell(row=r, column=c + i).value for i in range(n)]
            vals_norm = [v.strip().lower() if isinstance(v, str) else None for v in vals]
            if vals_norm == seq_norm:
                return r, c
    raise ValueError(f"Header sequence {sequence!r} not found in sheet {ws.title!r}")


def find_marker_row(ws, col, value, max_row=None):
    """Find the row where column `col` holds exactly `value` (case-insensitive)."""
    max_row = max_row or ws.max_row
    target = value.strip().lower()
    for r in range(1, max_row + 1):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.strip().lower() == target:
            return r
    raise ValueError(f"Marker {value!r} not found in column {col} of sheet {ws.title!r}")


def clear_row_range(ws, row, col_start, col_end):
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).value = None
