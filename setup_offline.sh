#!/usr/bin/env bash
# =============================================================================
#  setup_offline.sh ─ macOS 용 "인터넷 없는 PC" 설치
#
#  미리 인터넷 되는 컴퓨터에서  bash tools/make_offline_bundle.sh mac  으로 만든
#  offline_bundle/ 폴더가 이 폴더 안에 있어야 합니다.
#
#    offline_bundle/wheels/   파이썬 라이브러리
#    offline_bundle/python/   python-3.12.x-macos11.pkg   (Python 이 없을 때만 사용)
#    offline_bundle/ollama/   Ollama-darwin.zip           (Ollama 가 없을 때만 사용)
#    offline_bundle/models/   모델 파일 (약 9GB)
#
#  사용법   bash setup_offline.sh
# =============================================================================
set -u
cd "$(dirname "$0")"
B="$PWD/offline_bundle"
MODEL="${PM_MODEL:-qwen3:14b}"
OLLAMA_URL="http://127.0.0.1:11434"

say()  { printf '\n\033[1;36m▶ %s\033[0m\n' "$*"; }
ok()   { printf '   \033[1;32m✔ %s\033[0m\n' "$*"; }
warn() { printf '   \033[1;33m! %s\033[0m\n' "$*"; }
die()  { printf '\n\033[1;31m✖ %s\033[0m\n\n' "$*"; exit 1; }

[ -d "$B/wheels" ] || die "offline_bundle/wheels 가 없습니다. 인터넷 되는 PC 에서 tools/make_offline_bundle.sh mac 으로 만들어 함께 복사하세요."

# ── 1. Python ─────────────────────────────────────────────────
say "1/5  Python 확인"
PY=""
for c in python3.12 python3.13 python3.11 python3 /Library/Frameworks/Python.framework/Versions/3.12/bin/python3.12; do
  if command -v "$c" >/dev/null 2>&1 && \
     "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3,11) else 1)' 2>/dev/null; then
    PY="$c"; break
  fi
done
if [ -z "$PY" ]; then
  PKG="$(ls "$B"/python/python-*.pkg 2>/dev/null | head -1)"
  [ -n "$PKG" ] || die "Python 3.11+ 가 없고 offline_bundle/python 에 설치 파일도 없습니다."
  warn "Python 을 설치합니다 (관리자 비밀번호를 묻습니다): $PKG"
  sudo installer -pkg "$PKG" -target / || die "Python 설치 실패"
  PY=/Library/Frameworks/Python.framework/Versions/3.12/bin/python3.12
fi
ok "$("$PY" --version) 사용"

# ── 2. 가상환경 + 라이브러리 (오프라인) ───────────────────────
say "2/5  가상환경 및 라이브러리 (오프라인 휠)"
[ -x .venv/bin/python ] || "$PY" -m venv .venv || die "가상환경 생성 실패"
.venv/bin/python -m pip install --no-index --find-links "$B/wheels" -r requirements.txt \
  || die "휠 설치 실패 ─ offline_bundle 이 이 PC 의 OS/Python 버전에 맞는지 확인하세요"
.venv/bin/python -c "import pptx, fastapi, uvicorn, multipart" || die "라이브러리 확인 실패"
ok "라이브러리 설치 완료"

# ── 3. Ollama ─────────────────────────────────────────────────
say "3/5  Ollama"
OLLAMA="$(command -v ollama || true)"
[ -z "$OLLAMA" ] && [ -x /Applications/Ollama.app/Contents/Resources/ollama ] && \
  OLLAMA=/Applications/Ollama.app/Contents/Resources/ollama
if [ -z "$OLLAMA" ]; then
  [ -f "$B/ollama/Ollama-darwin.zip" ] || die "Ollama 가 없고 offline_bundle/ollama/Ollama-darwin.zip 도 없습니다."
  unzip -q -o "$B/ollama/Ollama-darwin.zip" -d /Applications || die "Ollama 압축 해제 실패"
  OLLAMA=/Applications/Ollama.app/Contents/Resources/ollama
fi
ok "Ollama: $OLLAMA"

# ── 4. 모델 파일 복사 ─────────────────────────────────────────
say "4/5  모델 파일 복사 (약 9GB)"
if [ -d "$B/models/manifests" ]; then
  mkdir -p "$HOME/.ollama/models"
  rsync -a "$B/models/" "$HOME/.ollama/models/" || cp -R "$B/models/." "$HOME/.ollama/models/"
  ok "복사 완료 → ~/.ollama/models"
else
  warn "offline_bundle/models 가 없습니다. 모델을 따로 준비해야 합니다."
fi

# ── 5. 서버 기동 + 확인 ───────────────────────────────────────
say "5/5  Ollama 서버 기동 및 모델 확인"
if ! curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1; then
  nohup "$OLLAMA" serve >/tmp/ollama.log 2>&1 &
  for _ in $(seq 1 30); do
    curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1 && break
    sleep 1
  done
fi
"$OLLAMA" list 2>/dev/null | grep -q "^$MODEL" && ok "모델 $MODEL 준비됨" || warn "$MODEL 이 목록에 없습니다"

printf '\n\033[1;32m설치가 끝났습니다.  실행:  bash run.sh\033[0m\n\n'
