#!/usr/bin/env bash
# =============================================================================
#  tools/make_offline_bundle.sh ─ "인터넷 없는 PC" 용 설치 꾸러미 만들기
#
#  ★ 인터넷이 되는 컴퓨터(지금 이 Mac)에서 실행합니다.
#     결과물 offline_bundle/ 폴더를 프로젝트와 함께 USB 등으로 옮긴 뒤,
#     대상 PC 에서 setup_offline.bat (Windows) / setup_offline.sh (macOS) 를 실행하면 됩니다.
#
#  사용법   bash tools/make_offline_bundle.sh windows     ← Windows 64bit 용
#           bash tools/make_offline_bundle.sh mac         ← Apple Silicon Mac 용
#
#  만들어지는 것 (약 10GB)
#    offline_bundle/wheels/   파이썬 라이브러리 (대상 OS 용)
#    offline_bundle/python/   Python 3.12 설치 파일
#    offline_bundle/ollama/   Ollama 설치 파일
#    offline_bundle/models/   ~/.ollama/models 복사본 (qwen3:14b + qwen3:8b, 약 14GB)
#    offline_bundle/README.txt
# =============================================================================
set -euo pipefail
cd "$(dirname "$0")/.."
TARGET="${1:-windows}"
PYVER="3.12.10"                      # python.org 에 설치 파일이 있는 마지막 3.12 버전
OUT="offline_bundle"
# pip 이 있는 Python 찾기 (uv 로 만든 .venv 에는 pip 이 없을 수 있어 ensurepip 으로 보강)
PY=""
for c in .venv/bin/python python3.12 python3; do
  { [ -x "$c" ] || command -v "$c" >/dev/null 2>&1; } || continue
  if "$c" -m pip --version >/dev/null 2>&1; then PY="$c"; break; fi
  if "$c" -m ensurepip --upgrade >/dev/null 2>&1 && "$c" -m pip --version >/dev/null 2>&1; then PY="$c"; break; fi
done

say() { printf '\n\033[1;36m▶ %s\033[0m\n' "$*"; }
die() { printf '\n\033[1;31m✖ %s\033[0m\n\n' "$*"; exit 1; }

case "$TARGET" in
  windows|mac) ;;
  *) die "대상 OS 를 지정하세요:  windows  또는  mac" ;;
esac
[ -n "$PY" ] || die "pip 이 있는 Python 을 찾지 못했습니다. python.org 에서 Python 3.12 를 설치하세요."
mkdir -p "$OUT"/{wheels,python,ollama,models}

# ── 1. 파이썬 라이브러리 ──────────────────────────────────────
say "1/4  파이썬 라이브러리(wheel) 내려받기 → $OUT/wheels   [$TARGET]"
if [ "$TARGET" = windows ]; then
  # 이 Mac 에서 Windows 용 휠을 받는다 (--only-binary 가 있어야 교차 플랫폼 다운로드 가능)
  "$PY" -m pip download -r requirements.txt -d "$OUT/wheels" \
      --python-version 3.12 --implementation cp --abi cp312 --only-binary=:all: \
      --platform win_amd64 --quiet
else
  # 같은 종류의 Mac(Apple Silicon) 이 대상이면 이 컴퓨터용 휠을 그대로 쓴다
  "$PY" -m pip download -r requirements.txt -d "$OUT/wheels" --only-binary=:all: --quiet
fi
echo "   $(ls "$OUT/wheels" | wc -l | tr -d ' ') 개 파일"

# ── 2. Python 설치 파일 ───────────────────────────────────────
say "2/4  Python $PYVER 설치 파일"
if [ "$TARGET" = windows ]; then f="python-$PYVER-amd64.exe"; else f="python-$PYVER-macos11.pkg"; fi
if [ -s "$OUT/python/$f" ]; then echo "   이미 있음: $f"
else curl -L --fail --progress-bar -o "$OUT/python/$f" "https://www.python.org/ftp/python/$PYVER/$f" || die "Python 설치 파일 다운로드 실패"; fi

# ── 3. Ollama 설치 파일 ───────────────────────────────────────
say "3/4  Ollama 설치 파일"
if [ "$TARGET" = windows ]; then
  [ -s "$OUT/ollama/OllamaSetup.exe" ] && echo "   이미 있음" || \
    curl -L --fail --progress-bar -o "$OUT/ollama/OllamaSetup.exe" https://ollama.com/download/OllamaSetup.exe || die "Ollama 다운로드 실패"
else
  [ -s "$OUT/ollama/Ollama-darwin.zip" ] && echo "   이미 있음" || \
    curl -L --fail --progress-bar -o "$OUT/ollama/Ollama-darwin.zip" https://ollama.com/download/Ollama-darwin.zip || die "Ollama 다운로드 실패"
fi

# ── 4. 모델 파일 ──────────────────────────────────────────────
say "4/4  모델 파일 복사  (~/.ollama/models → $OUT/models, 약 9GB)"
SRC="${OLLAMA_MODELS:-$HOME/.ollama/models}"
[ -d "$SRC/manifests" ] || die "$SRC 에 모델이 없습니다. 먼저 'ollama pull qwen3:14b' 를 하세요."
rsync -a "$SRC/" "$OUT/models/" || cp -R "$SRC/." "$OUT/models/"

cat > "$OUT/README.txt" <<TXT
특허 마킹 도구 ─ 오프라인 설치 꾸러미  (대상: $TARGET, 생성일: $(date +%Y-%m-%d))

이 폴더(offline_bundle)를 patent_marker 프로젝트 폴더 안에 둔 채로
  Windows : setup_offline.bat 더블클릭
  macOS   : bash setup_offline.sh
를 실행하면 인터넷 없이 설치가 끝납니다.

wheels/  파이썬 라이브러리      python/  Python $PYVER 설치 파일
ollama/  Ollama 설치 파일       models/  qwen3:14b 모델 (약 9GB)
TXT

printf '\n\033[1;32m완료.  크기: %s\033[0m\n' "$(du -sh "$OUT" | cut -f1)"
printf '이제 프로젝트 폴더 전체(offline_bundle 포함)를 대상 PC 로 복사하세요.\n\n'
