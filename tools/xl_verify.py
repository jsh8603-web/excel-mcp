# /// script
# requires-python = ">=3.10"
# dependencies = ["formulon>=0.18"]
# ///
"""formulon(엑셀 호환 엔진)으로 통합문서를 다시 계산해서 점검한다. 파일은 저장하지 않는다.

세 가지 모드
  기본          엑셀이 마지막 저장 때 계산해 둔 값 vs 재계산 값 → 깨진 수식, 손으로 덮어쓴 값 찾기
  --baseline 원본  편집한 파일을 재계산해서 '편집 전 원본(엑셀 저장값)'과 비교 → 편집 영향 보고
  --print 범위   재계산한 값을 그대로 출력 (LAMBDA 등 다른 도구가 계산 못한 셀 확인용)

사용
  python xl_verify.py 파일.xlsx [--sheet 시트]
  python xl_verify.py 편집본.xlsx --baseline 원본.xlsx
  python xl_verify.py 파일.xlsx --print 'Schedule!A1:F20'

종료 코드: 0 문제 없음 / 1 불일치·오류값 있음 / 3 기준값이 없어 비교 불가
(MCP 도구로 저장한 파일은 엑셀 계산값이 지워지거나(haris-musa) 옛 값이 남는다(xlide).
 편집 후 점검은 --baseline으로 원본과 비교해야 정확하다.)
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from xlcommon import (EXTERNAL, STRINGS, VOLATILE, a1, fmt, load, parse_range,  # noqa: E402
                      py, recalc, same, sheet_index)


def formula_cells(wb, sheets):
    out = {}
    for s in sheets:
        for c in wb.iter_cells(s):
            if c.formula:
                out[(s, c.row, c.col)] = (c.formula, py(c.value))
    return out


def skipped_cells(wb, cells) -> set:
    """휘발성 함수·외부 참조 셀과 그 셀을 참조하는 모든 셀 (엑셀 값과 비교 불가)."""
    roots = {k for k, (formula, _) in cells.items()
             if VOLATILE.search(formula) or EXTERNAL.search(STRINGS.sub('""', formula))}
    out = set(roots)
    for s, r, c in roots:
        for n in wb.dependents(s, r, c, depth=64):
            out.add((n.sheet, n.row, n.col))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--sheet")
    ap.add_argument("--baseline", help="편집 전 원본 파일 (엑셀로 저장된 것)")
    ap.add_argument("--print", dest="print_range", help="재계산 값을 출력할 범위 (예: 'Sheet1!A1:D10')")
    ap.add_argument("--show", type=int, default=40)
    args = ap.parse_args()

    wb = load(args.path)
    sheets = [sheet_index(wb, args.sheet)] if args.sheet else list(range(wb.sheet_count()))
    cells = formula_cells(wb, sheets)
    if args.baseline:
        base = load(args.baseline)
        names = {base.sheet_name(i): i for i in range(base.sheet_count())}
        cached = {}
        for (s, r, c) in cells:
            bs = names.get(wb.sheet_name(s))
            cached[(s, r, c)] = py(base.get_value(bs, r, c)) if bs is not None else None
    else:
        cached = {k: v for k, (_, v) in cells.items()}

    recalc(wb)

    if args.print_range:
        s, r0, c0, r1, c1 = parse_range(wb, args.print_range, args.sheet)
        print(f"{wb.sheet_name(s)}!{a1(r0, c0)}:{a1(r1, c1)} (formulon 재계산 값)")
        for r in range(r0, r1 + 1):
            print("\t".join(fmt(py(wb.get_value(s, r, c))) for c in range(c0, c1 + 1)))
        return 0

    skip = skipped_cells(wb, cells)
    ok, diffs, nobase, errors = 0, [], [], []
    for key, (formula, _) in cells.items():
        s, r, c = key
        before, after = cached[key], py(wb.get_value(s, r, c))
        where = f"{wb.sheet_name(s)}!{a1(r, c)}"
        if isinstance(after, str) and after.startswith("#ERR"):
            errors.append((where, formula, after))
        if key in skip:
            continue
        if before is None:
            nobase.append(where)
        elif same(before, after):
            ok += 1
        else:
            diffs.append((where, formula, before, after))

    label = "원본" if args.baseline else "엑셀"
    print(f"수식 셀 {len(cells):,}개 | 같음 {ok:,} | {'바뀜' if args.baseline else '불일치'} {len(diffs):,} | "
          f"비교제외(휘발성·외부참조와 그 의존 셀) {len(skip):,} | 기준없음 {len(nobase):,} | 오류값 {len(errors):,}")
    for where, formula, before, after in diffs[: args.show]:
        print(f"  {'→' if args.baseline else '✗'} {where:<22} {label}={fmt(before):<18} 재계산={fmt(after):<18} {formula[:70]}")
    for where, formula, after in errors[: args.show]:
        print(f"  ! {where:<22} {after:<14} {formula[:80]}")

    compared = len(cells) - len(skip)
    if compared and len(nobase) > compared / 2:
        print("\n⚠ 비교할 기준값이 대부분 없습니다. MCP 도구로 저장한 파일이면 편집 전 원본과 비교하세요: "
              "--baseline 원본.xlsx")
        return 3
    if args.baseline:
        return 1 if errors else 0
    return 1 if diffs or errors else 0


if __name__ == "__main__":
    sys.exit(main())
