#!/bin/bash
# Mac 엑셀에서 Power Query를 새로 고친 뒤 저장하고 닫는다.
# 사용: tools/pq_refresh.sh 파일.xlsx
#
# ⚠️ Mac에서 아직 검증하지 않은 스크립트. 처음엔 복사본으로 시험할 것.
# 전제
#  - 처음 실행 시 "터미널이 Microsoft Excel을 제어" 허용 창이 뜬다 → 허용
#  - 쿼리 속성에서 '백그라운드 새로 고침 사용'을 끄면 새로 고침이 끝난 뒤에 저장된다
#    (데이터 > 쿼리 및 연결 > 쿼리 우클릭 > 속성)
set -euo pipefail
f="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
[ -f "$f" ] || { echo "파일 없음: $1"; exit 1; }
n="$(basename "$f")"; d="$(dirname "$f")"
if [ -e "$d/~\$$n" ] || [ -e "$d/~\$${n:2}" ]; then
  echo "이미 엑셀에서 열려 있습니다. 저장 안 된 변경을 덮어쓰지 않도록 저장하고 닫은 뒤 다시 실행하세요."; exit 2
fi
WAIT="${PQ_WAIT:-15}"   # 새로 고침을 기다리는 초. 쿼리가 오래 걸리면 PQ_WAIT=60 처럼 늘린다

osascript - "$f" "$WAIT" <<'APPLESCRIPT'
on run argv
    set p to item 1 of argv
    tell application "Microsoft Excel"
        activate
        open (POSIX file p as alias)
        set wb to active workbook
        refresh all wb
        delay (item 2 of argv as integer)
        save wb
        close wb saving no
    end tell
    return "refreshed: " & p
end run
APPLESCRIPT
