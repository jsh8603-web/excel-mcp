"""xl_verify.py, xl_whatif.py 공용 함수 (formulon 기반, 원본 파일은 절대 저장하지 않음)."""
import re
from pathlib import Path

import formulon as f

VOLATILE = re.compile(r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(", re.I)
STRINGS = re.compile(r'"(?:[^"]|"")*"')
# 파일 안에서 외부 통합문서 참조는 [1]Sheet!A1 처럼 번호로 저장된다. 표 참조 T[Amt], [@Amt]는 제외
EXTERNAL = re.compile(r"(?<![\w\]])'?\[\d+\]")
_ADDR = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^!']+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$")


def load(path: str) -> f.Workbook:
    return f.Workbook.load(Path(path).expanduser().read_bytes())


def sheet_index(wb: f.Workbook, name: str | None) -> int:
    if name is None:
        return 0
    for i in range(wb.sheet_count()):
        if wb.sheet_name(i).lower() == name.lower():
            return i
    raise SystemExit(f"시트를 찾을 수 없음: {name}")


def parse_addr(wb: f.Workbook, text: str, default_sheet: str | None = None) -> tuple[int, int, int]:
    """'Sheet1!B3', "'O''Brien'!B3", 'B3' -> (sheet, row0, col0)"""
    m = _ADDR.match(text.strip())
    if not m:
        raise SystemExit(f"셀 주소 형식 오류: {text}")
    sheet = (m.group(1).replace("''", "'") if m.group(1) else m.group(2)) or default_sheet
    col = 0
    for ch in m.group(3).upper():
        col = col * 26 + (ord(ch) - 64)
    return sheet_index(wb, sheet), int(m.group(4)) - 1, col - 1


def parse_range(wb: f.Workbook, text: str, default_sheet: str | None = None):
    """'Sheet!A1:C5' -> (sheet, r0, c0, r1, c1)"""
    head, _, tail = text.rpartition(":")
    if not head:
        s, r, c = parse_addr(wb, text, default_sheet)
        return s, r, c, r, c
    s, r0, c0 = parse_addr(wb, head, default_sheet)
    _, r1, c1 = parse_addr(wb, tail, wb.sheet_name(s))
    return s, min(r0, r1), min(c0, c1), max(r0, r1), max(c0, c1)


def a1(row0: int, col0: int) -> str:
    s, c = "", col0 + 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return f"{s}{row0 + 1}"


def py(v: f.Value):
    k = v.kind
    if k == f.ValueKind.NUMBER:
        return v.number
    if k == f.ValueKind.TEXT:
        return v.text
    if k == f.ValueKind.BOOL:
        return v.boolean
    if k == f.ValueKind.ERROR:
        try:
            return f"#ERR:{f.error_display_name(v.error_code)}"
        except Exception:
            return f"#ERR:{v.error_code}"
    return None


def same(a, b, rel: float = 1e-9, abs_: float = 1e-6) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return type(a) is type(b) and a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= max(abs_, rel * max(abs(a), abs(b)))
    if a in (None, "") and b in (None, ""):
        return True
    return a == b


def fmt(x) -> str:
    if isinstance(x, bool) or not isinstance(x, (int, float)):
        return str(x)
    return f"{x:,.2f}" if abs(x) >= 1000 else f"{x:.6g}"


def recalc(wb: f.Workbook) -> None:
    wb.recalc()
    if hasattr(wb, "recalc_data_tables"):
        wb.recalc_data_tables()
