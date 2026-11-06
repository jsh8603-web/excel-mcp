#!/bin/bash
# xlkit 설치 (macOS, Apple Silicon) — 한 번만 실행하면 모든 프로젝트에서 쓸 수 있다.
#
#   bash setup.sh             모든 프로젝트에서 자동으로 쓰기 (사용자 범위 등록)
#   bash setup.sh --project   설치만 하고, 필요한 프로젝트에서 직접 등록:
#       cd 프로젝트폴더 && claude mcp add -s project xlkit -- ~/.xlkit/bin/xlkit
#
# 하는 일
#  1) ~/.xlkit 에 전용 Python 환경을 만들고 wheels/ 에서 오프라인 설치 (PyPI 접속 불필요)
#  2) MCP 서버 'xlkit' 하나를 Claude Code 사용자 범위에 등록 (모든 프로젝트)
#  3) (선택) VS Code Power Query 편집기 확장 설치
# 다시 실행하면 업데이트된다. 필요한 것: uv, Claude Code. Python 3.12는 없으면 uv가 받는다.
set -euo pipefail

KIT="$(cd "$(dirname "$0")" && pwd)"
WHEELS="${XLKIT_WHEELS:-$KIT/wheels}"
LOCK="${XLKIT_LOCK:-$KIT/requirements.lock}"
HOME_KIT="${XLKIT_HOME:-$HOME/.xlkit}"
VENV="$HOME_KIT/venv"
SCOPE=user; [ "${1:-}" = "--project" ] && SCOPE=project

command -v uv >/dev/null || { echo "uv가 없습니다. 설치: brew install uv"; exit 1; }

echo "▶ 1/3 설치 → $HOME_KIT"
uv venv --quiet --allow-existing --python 3.12 "$VENV"
uv pip install --quiet --python "$VENV/bin/python" --no-index --find-links "$WHEELS" -r "$LOCK"
"$VENV/bin/python" -c "import formulon, excel_mcp, xlide_mcp" || { echo "⚠ 패키지 확인 실패"; exit 1; }
mkdir -p "$HOME_KIT/tools" "$HOME_KIT/bin"
cp "$KIT"/tools/*.py "$KIT"/tools/*.sh "$HOME_KIT/tools/"
cat > "$HOME_KIT/bin/xlkit" <<SH
#!/bin/bash
# Claude Code가 연 프로젝트 폴더(CLAUDE_PROJECT_DIR, 없으면 현재 폴더) 안 파일만 다룬다
exec "$VENV/bin/python" "$HOME_KIT/tools/xlkit_mcp.py" "\$@"
SH
chmod +x "$HOME_KIT/bin/xlkit" "$HOME_KIT"/tools/*.sh

echo "▶ 2/3 MCP 등록"
if [ "$SCOPE" = project ]; then
  echo "  프로젝트별 등록 모드: 쓸 프로젝트 폴더에서 아래 명령 실행"
  echo "    claude mcp add -s project xlkit -- \"$HOME_KIT/bin/xlkit\""
elif command -v claude >/dev/null; then
  claude mcp remove -s user xlkit >/dev/null 2>&1 || true
  claude mcp add -s user xlkit -- "$HOME_KIT/bin/xlkit" >/dev/null && echo "  xlkit 등록됨 (사용자 범위: 모든 프로젝트)"
else
  echo "  ⚠ claude 명령이 없습니다. 나중에 실행: claude mcp add -s user xlkit -- \"$HOME_KIT/bin/xlkit\""
fi

echo "▶ 3/3 VS Code 확장 (선택: 사람이 M 코드를 편집할 때)"
if command -v code >/dev/null; then
  code --install-extension ewc3labs.excel-power-query-editor >/dev/null && echo "  설치됨" || echo "  ⚠ 설치 실패"
else
  echo "  건너뜀 ('code' 명령 없음)"
fi

echo
if [ "$SCOPE" = project ]; then
  echo "✅ 완료. 쓸 프로젝트에서 위 'claude mcp add' 명령 실행 → /mcp 에서 xlkit 확인"
else
  echo "✅ 완료. 아무 프로젝트에서 claude 실행 → /mcp 에서 xlkit 연결 확인"
fi
echo "   예전에 haris-musa 등을 따로 등록했다면 중복 방지: claude mcp list 확인 후 claude mcp remove <이름>"
