#!/usr/bin/env bash
# =============================================================================
#  setup.sh ─ macOS / Linux 용 "처음 한 번" 설치 (인터넷 필요)
#
#  하는 일
#    1) Python 3.11+ 확인 (없으면 uv 로 3.12 자동 설치)
#    2) .venv 가상환경 만들고 라이브러리 설치
#    3) Ollama 설치
#    4) Ollama 서버 기동
#    5) 모델(qwen3:14b) 내려받기   ← 약 9GB. 가장 오래 걸리는 단계
#
#  사용법      bash setup.sh
#  모델 변경   PM_MODEL=qwen3:8b bash setup.sh
#  오프라인    인터넷이 없는 PC 에서는 setup_offline.sh 를 쓰세요.
# =============================================================================
set -u
cd "$(dirname "$0")"
MODEL="${PM_MODEL:-qwen3:14b}"
OLLAMA_URL="http://127.0.0.1:11434"

say()  { printf '\n\033[1;36m▶ %s\033[0m\n' "$*"; }
ok()   { printf '   \033[1;32m✔ %s\033[0m\n' "$*"; }
warn() { printf '   \033[1;33m! %s\033[0m\n' "$*"; }
die()  { printf '\n\033[1;31m✖ %s\033[0m\n\n' "$*"; exit 1; }

# ── 1. Python ─────────────────────────────────────────────────
say "1/5  Python 확인"
PY=""; USE_UV=0
for c in python3.12 python3.13 python3.11 python3; do
  if command -v "$c" >/dev/null 2>&1 && \
     "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3,11) else 1)' 2>/dev/null; then
    PY="$c"; break
  fi
done
if [ -n "$PY" ]; then
  ok "$("$PY" --version) 사용"
else
  warn "Python 3.11 이상이 없습니다. uv 로 Python 3.12 를 설치합니다 (관리자 권한 불필요)."
  if ! command -v uv >/dev/null 2>&1; then
    curl -LsSf https://astral.sh/uv/install.sh | sh || die "uv 설치 실패 ─ 인터넷 연결을 확인하세요"
    export PATH="$HOME/.local/bin:$PATH"
  fi
  uv python install 3.12 || die "Python 설치 실패"
  USE_UV=1
  ok "uv 로 Python 3.12 준비"
fi

# ── 2. 가상환경 + 라이브러리 ──────────────────────────────────
say "2/5  가상환경(.venv) 및 라이브러리 설치"
if [ ! -x .venv/bin/python ]; then
  if [ "$USE_UV" = 1 ]; then
    uv venv --python 3.12 .venv || die "가상환경 생성 실패"
  else
    "$PY" -m venv .venv || die "가상환경 생성 실패"
  fi
fi
if [ "$USE_UV" = 1 ]; then
  uv pip install --python .venv/bin/python -r requirements.txt || die "라이브러리 설치 실패"
else
  .venv/bin/python -m pip install --upgrade pip >/dev/null 2>&1
  .venv/bin/python -m pip install -r requirements.txt || die "라이브러리 설치 실패 ─ 인터넷 연결을 확인하세요"
fi
.venv/bin/python -c "import pptx, fastapi, uvicorn, multipart" || die "라이브러리 확인 실패"
ok "라이브러리 설치 완료"

# ── 3. Ollama ─────────────────────────────────────────────────
say "3/5  Ollama 설치 확인"
OLLAMA="$(command -v ollama || true)"
[ -z "$OLLAMA" ] && [ -x /Applications/Ollama.app/Contents/Resources/ollama ] && \
  OLLAMA=/Applications/Ollama.app/Contents/Resources/ollama
if [ -z "$OLLAMA" ]; then
  case "$(uname -s)" in
    Darwin)
      if command -v brew >/dev/null 2>&1; then
        brew install ollama || die "brew 로 Ollama 설치 실패"
        OLLAMA="$(command -v ollama)"
      else
        warn "Homebrew 가 없어 Ollama 앱을 직접 내려받습니다."
        curl -L --fail -o /tmp/Ollama-darwin.zip https://ollama.com/download/Ollama-darwin.zip \
          || die "Ollama 다운로드 실패"
        unzip -q -o /tmp/Ollama-darwin.zip -d /Applications || die "압축 해제 실패"
        OLLAMA=/Applications/Ollama.app/Contents/Resources/ollama
      fi ;;
    Linux)
      curl -fsSL https://ollama.com/install.sh | sh || die "Ollama 설치 실패"
      OLLAMA="$(command -v ollama)" ;;
    *) die "지원하지 않는 OS 입니다. Windows 는 setup.bat 을 쓰세요." ;;
  esac
fi
ok "Ollama: $OLLAMA"

# ── 4. Ollama 서버 기동 ───────────────────────────────────────
say "4/5  Ollama 서버 기동"
if curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1; then
  ok "이미 실행 중"
else
  nohup "$OLLAMA" serve >/tmp/ollama.log 2>&1 &
  for _ in $(seq 1 30); do
    curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1 && break
    sleep 1
  done
  curl -sf "$OLLAMA_URL/api/tags" >/dev/null 2>&1 || die "Ollama 서버가 응답하지 않습니다 (/tmp/ollama.log 확인)"
  ok "기동 완료"
fi

# ── 5. 모델 ───────────────────────────────────────────────────
say "5/5  모델 내려받기 ($MODEL)  ─ 약 9GB, 수 분 ~ 수십 분"
if "$OLLAMA" list 2>/dev/null | grep -q "^$MODEL"; then
  ok "이미 있음"
else
  "$OLLAMA" pull "$MODEL" || die "모델 다운로드 실패 ─ 인터넷 연결을 확인하세요"
  ok "다운로드 완료"
fi

printf '\n\033[1;32m설치가 끝났습니다.  실행:  bash run.sh\033[0m\n\n'
