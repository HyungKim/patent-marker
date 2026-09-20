#!/usr/bin/env bash
# =============================================================================
#  mark.sh ─ 브라우저 없이 마킹하기 (macOS / Linux)
#    bash mark.sh 보고서.pptx [다른.pptx ...]
#    결과는 원본 옆에  파일이름_특허마킹.pptx  로 저장됩니다.
#    (Windows 에서는 mark.bat 위에 파일을 끌어다 놓으면 됩니다)
# =============================================================================
set -u
cd "$(dirname "$0")"
OLLAMA_URL="${PM_OLLAMA_HOST:-http://127.0.0.1:11434}"

if [ $# -eq 0 ]; then
  echo "사용법: bash mark.sh 보고서.pptx [다른.pptx ...]"
  exit 1
fi
if [ ! -x .venv/bin/python ]; then
  echo "먼저 'bash setup.sh' (인터넷 있음) 또는 'bash setup_offline.sh' (인터넷 없음) 를 실행하세요."
  exit 1
fi

OLLAMA="$(command -v ollama || true)"
[ -z "$OLLAMA" ] && [ -x /Applications/Ollama.app/Contents/Resources/ollama ] && \
  OLLAMA=/Applications/Ollama.app/Contents/Resources/ollama

if ! curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1; then
  if [ -n "$OLLAMA" ]; then
    echo "Ollama 를 시작합니다…"
    nohup "$OLLAMA" serve >/tmp/ollama.log 2>&1 &
    for _ in $(seq 1 30); do
      curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1 && break
      sleep 1
    done
  else
    echo "[경고] Ollama 가 설치되어 있지 않습니다. setup.sh 를 먼저 실행하세요."
  fi
fi

exec .venv/bin/python -m app.cli "$@"
