# /// script
# requires-python = ">=3.10"
# dependencies = ["formulon>=0.18"]
# ///
"""What-If 민감도 분석 (엑셀 '데이터 표'와 같은 결과, 원본 파일은 저장하지 않음).

입력 셀 1~2개를 여러 값으로 바꿔 가며 통합문서 전체를 다시 계산하고, 결과 셀 값을 표로 출력한다.

사용 예
  # 할인율 1개 변수 -> 리스부채 기초잔액, 월 상각액
  ~/.xlkit/venv/bin/python ~/.xlkit/tools/xl_whatif.py lease.xlsx --sheet Assump --in B3 --vals 0.04,0.045,0.05 \
      --out 'Schedule!E5' --out 'Schedule!F5'

  # 2개 변수 (할인율 x 리스기간)
  ~/.xlkit/venv/bin/python ~/.xlkit/tools/xl_whatif.py lease.xlsx --sheet Assump --in B3 --vals 0.04,0.05 \
      --in2 B4 --vals2 36,48,60 --out 'Schedule!E5' --csv result.csv
"""
import argparse
import csv
import itertools
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from xlcommon import fmt, load, parse_addr, py, recalc  # noqa: E402


def nums(text: str) -> list[float]:
    return [float(x) for x in text.split(",") if x.strip()]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--sheet", help="--in/--out에 시트명을 안 적었을 때 기본 시트")
    ap.add_argument("--in", dest="in1", required=True)
    ap.add_argument("--vals", required=True)
    ap.add_argument("--in2")
    ap.add_argument("--vals2")
    ap.add_argument("--out", action="append", required=True)
    ap.add_argument("--csv")
    args = ap.parse_args()

    if bool(args.in2) != bool(args.vals2):
        ap.error("--in2와 --vals2는 함께 줘야 합니다")
    raw = Path(args.path).expanduser().read_bytes()
    probe = load(args.path)
    c1 = parse_addr(probe, args.in1, args.sheet)
    c2 = parse_addr(probe, args.in2, args.sheet) if args.in2 else None
    outs = [(o, parse_addr(probe, o, args.sheet)) for o in args.out]
    for addr in [c1] + ([c2] if c2 else []):
        if probe.get_formula(*addr):
            print(f"경고: 입력 셀 {addr}에 수식이 있습니다. 값으로 덮어써서 계산합니다.", file=sys.stderr)

    grid1, grid2 = nums(args.vals), nums(args.vals2) if c2 else [None]
    header = [args.in1] + ([args.in2] if c2 else []) + [o for o, _ in outs]
    rows = []
    import formulon as f
    for v1, v2 in itertools.product(grid1, grid2):
        wb = f.Workbook.load(raw)  # 매 조합마다 원본에서 새로 읽음 (상태 섞임 방지)
        wb.set_number(*c1, v1)
        if c2:
            wb.set_number(*c2, v2)
        recalc(wb)
        rows.append([v1] + ([v2] if c2 else []) + [py(wb.get_value(*addr)) for _, addr in outs])

    widths = [max(len(str(h)), 14) for h in header]
    print("  ".join(str(h).rjust(w) for h, w in zip(header, widths)))
    for r in rows:
        print("  ".join(fmt(x).rjust(w) for x, w in zip(r, widths)))
    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as fh:
            csv.writer(fh).writerows([header] + rows)
        print(f"\nCSV 저장: {args.csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
