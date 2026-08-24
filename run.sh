#!/usr/bin/env bash
# =============================================================================
#  run.sh ─ 프로그램 실행 (macOS / Linux)
#    · Ollama 가 꺼져 있으면 켜고
#    · 웹 서버를 띄운 뒤
#    · 브라우저를 자동으로 엽니다  →  http://127.0.0.1:8765
#  종료: 이 창에서 Ctrl + C
#  (설치가 안 되어 있으면 먼저 setup.sh 또는 setup_offline.sh)
# =============================================================================
set -u
cd "$(dirname "$0")"
OLLAMA_URL="${PM_OLLAMA_HOST:-http://127.0.0.1:11434}"
PORT="${PM_PORT:-8765}"

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

# 서버가 뜬 뒤(2초 후) 브라우저를 연다
( sleep 2
  if command -v open >/dev/null 2>&1; then open "http://127.0.0.1:$PORT"
  elif command -v xdg-open >/dev/null 2>&1; then xdg-open "http://127.0.0.1:$PORT"
  fi ) >/dev/null 2>&1 &

exec .venv/bin/python -m app.main
