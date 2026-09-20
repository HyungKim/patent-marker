"""
local.py ─ 브라우저 업로드 없이 파일을 받는 통로
=====================================================================

[왜 필요한가]
  회사 PC 는 보안 정책으로 브라우저의 '파일 올리기' 가 막혀 있는 경우가 많습니다.
  그런데 이 프로그램의 서버(파이썬)는 브라우저와 **같은 PC** 에서 돌고 있으므로,
  파일을 브라우저로 올릴 필요 없이 서버가 디스크에서 직접 읽으면 됩니다.
  이 파일은 "어느 파일인지" 를 서버에 알려 주는 세 가지 방법을 담당합니다.

      ① input 폴더  : patent_marker\\input\\ 에 복사해 두면 목록에 나타남   → list_files()
      ② 경로 붙여넣기: 탐색기 '경로로 복사' 로 얻은 전체 경로를 입력      → resolve_pptx()
      ③ 파일 선택창  : 파이썬이 Windows 파일 열기 창을 띄움 (브라우저 아님)  → pick_files()

  결과는 patent_marker\\output\\ 에 바로 저장됩니다 (다운로드가 막혀 있어도 됨).

[안전장치]
  - 경로로 받는 기능은 .pptx/.potx 파일만 허용합니다.
  - 서버는 127.0.0.1 전용이고, 이 기능들은 화면이 발급한 토큰(main.py)이 있어야 부를 수 있어
    브라우저의 다른 탭이 몰래 호출하지 못합니다.
"""
from __future__ import annotations

import importlib.util
import json
import os
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

from . import config

PPTX_SUFFIXES = (".pptx", ".potx")

# 파일 선택창은 한 번에 하나만 (두 개가 겹쳐 뜨면 사용자가 헷갈린다)
_PICK_LOCK = threading.Lock()
_PICK_TIMEOUT = 900          # 초. 사용자가 창을 열어 둔 채 자리를 비워도 이만큼은 기다린다


# ═════════════════════════════════════════════════════════════════
# 폴더
# ═════════════════════════════════════════════════════════════════
def ensure_dirs() -> None:
    """input / output 폴더가 없으면 만든다. 서버·명령행 시작 때 부른다."""
    config.INPUT_DIR.mkdir(parents=True, exist_ok=True)
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _is_pptx(p: Path) -> bool:
    return p.suffix.lower() in PPTX_SUFFIXES


def list_files(folder: Path) -> list[dict]:
    """폴더 안의 PPTX 목록 (최근 수정 순). 하위 폴더는 보지 않는다.

    PowerPoint 가 파일을 열어 둔 동안 만드는 잠금 파일(~$이름.pptx)과 숨김 파일은 뺀다.
    """
    if not folder.is_dir():
        return []
    rows: list[dict] = []
    for p in folder.iterdir():
        if not p.is_file() or not _is_pptx(p):
            continue
        if p.name.startswith(("~$", ".")):
            continue
        st = p.stat()
        if st.st_size == 0:              # 분석 중인 작업이 이름만 잡아 둔 빈 파일
            continue
        rows.append({
            "name": p.name,
            "path": str(p),
            "size": st.st_size,
            "mtime": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"),
            "_ts": st.st_mtime,
        })
    rows.sort(key=lambda r: (-r["_ts"], r["name"]))
    for r in rows:
        del r["_ts"]
    return rows


# ═════════════════════════════════════════════════════════════════
# 경로 입력
# ═════════════════════════════════════════════════════════════════
def clean_path(raw: str) -> str:
    """사람이 붙여 넣은 경로를 정리한다.

    - 탐색기의 '경로로 복사' 는 양쪽에 따옴표를 붙인다 → 벗긴다
    - 앞뒤 공백·줄바꿈 제거
    - file:///C:/... 형태(브라우저 주소창에서 복사한 경우)도 보통 경로로 바꾼다
    """
    s = (raw or "").strip()
    pairs = {('"', '"'), ("'", "'"), ("“", "”"), ("‘", "’")}
    while len(s) >= 2 and (s[0], s[-1]) in pairs:
        s = s[1:-1].strip()
    if s.lower().startswith("file:"):
        body = unquote(s[5:]).lstrip("/")
        # file:///C:/a/b.pptx → C:/a/b.pptx (Windows) · file:///Users/a/b.pptx → /Users/a/b.pptx (macOS)
        s = body if re.match(r"^[A-Za-z]:", body) else "/" + body
    return s


def resolve_pptx(raw: str) -> Path:
    """입력받은 경로를 실제 파일로 확정한다. 문제가 있으면 ValueError(한글 메시지).

    파일 이름만 적으면 input 폴더에서 찾는다.
    """
    s = clean_path(raw)
    if not s:
        raise ValueError("경로가 비어 있습니다.")
    p = Path(s).expanduser()
    if not p.is_absolute() and (config.INPUT_DIR / s).is_file():
        p = config.INPUT_DIR / s
    p = p.resolve()
    if p.is_dir():
        raise ValueError(f"폴더가 아니라 파일 경로를 입력하세요: {p}")
    if not p.is_file():
        raise ValueError(f"파일을 찾을 수 없습니다: {p}")
    if not _is_pptx(p):
        raise ValueError("PPTX 파일만 지원합니다. (.ppt 는 PowerPoint 에서 .pptx 로 저장한 뒤 사용하세요)")
    return p


# ═════════════════════════════════════════════════════════════════
# 결과 파일 이름
# ═════════════════════════════════════════════════════════════════
def output_path_for(src_name: str, out_dir: Path | None = None) -> Path:
    """결과 파일 경로. 이름이 겹치면 (2), (3) … 을 붙여 덮어쓰지 않는다.

        보고서.pptx → output\\보고서_특허마킹.pptx
                    → output\\보고서_특허마킹(2).pptx   (이미 있으면)
    """
    folder = out_dir or config.OUTPUT_DIR
    folder.mkdir(parents=True, exist_ok=True)
    stem = Path(src_name).stem or "deck"
    cand = folder / f"{stem}_특허마킹.pptx"
    n = 2
    while cand.exists():
        cand = folder / f"{stem}_특허마킹({n}).pptx"
        n += 1
    return cand


# ═════════════════════════════════════════════════════════════════
# 파일 선택창 (파이썬이 띄우는 네이티브 창)
# ═════════════════════════════════════════════════════════════════
def pick_available() -> bool:
    """이 Python 에 tkinter(파일 선택창) 가 들어 있는지."""
    return importlib.util.find_spec("_tkinter") is not None


def _pick_result(returncode: int, stdout: str, stderr: str) -> list[str]:
    """선택창 프로세스의 출력을 경로 목록으로 바꾼다. (테스트하기 쉽게 따로 뗌)"""
    if returncode != 0:
        if "tkinter" in stderr.lower():
            raise RuntimeError(
                "이 PC 의 Python 에는 파일 선택창(tkinter)이 없습니다. "
                "input 폴더에 복사하거나 경로를 직접 입력해 주세요.")
        tail = stderr.strip().splitlines()[-1] if stderr.strip() else f"코드 {returncode}"
        raise RuntimeError(f"파일 선택창을 열지 못했습니다: {tail}")
    line = stdout.strip().splitlines()[-1] if stdout.strip() else "[]"
    try:
        paths = json.loads(line)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"파일 선택창 응답을 읽지 못했습니다: {e}") from e
    return [str(p) for p in paths if p]


def pick_files(multiple: bool, initialdir: Path | None = None) -> list[str]:
    """Windows/macOS 의 파일 열기 창을 띄우고 고른 경로들을 돌려준다. 취소하면 빈 목록.

    브라우저가 아니라 파이썬 프로세스가 여는 창이므로 웹 업로드 차단과 무관하다.
    tkinter 가 없는 Python 이면 RuntimeError (화면이 안내문을 보여 준다).
    """
    if not _PICK_LOCK.acquire(blocking=False):
        raise RuntimeError("파일 선택창이 이미 열려 있습니다. 그 창을 먼저 닫아 주세요.")
    try:
        script = Path(__file__).with_name("pickdialog.py")
        start = str(initialdir or config.INPUT_DIR)
        kwargs: dict = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.run(
            [sys.executable, str(script), "multi" if multiple else "single", start],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            stdin=subprocess.DEVNULL, timeout=_PICK_TIMEOUT, **kwargs,
        )
        return _pick_result(proc.returncode, proc.stdout, proc.stderr)
    except subprocess.TimeoutExpired:
        raise RuntimeError("파일 선택창이 너무 오래 열려 있어 닫았습니다. 다시 시도하세요.")
    finally:
        _PICK_LOCK.release()


# ═════════════════════════════════════════════════════════════════
# 폴더 열기 (탐색기 / Finder)
# ═════════════════════════════════════════════════════════════════
def open_command(folder: Path, select: Path | None = None) -> list[str] | None:
    """운영체제별 '폴더 열기' 명령. select 가 있으면 그 파일이 선택된 채로 연다.

    Windows 는 os.startfile 을 써야 하므로 None 을 돌려주고 open_folder() 가 따로 처리한다.
    """
    if sys.platform == "win32":
        if select is not None and select.exists():
            return ["explorer.exe", "/select,", str(select)]
        return None
    if sys.platform == "darwin":
        if select is not None and select.exists():
            return ["open", "-R", str(select)]
        return ["open", str(folder)]
    return ["xdg-open", str(folder)]


def open_folder(folder: Path, select: Path | None = None) -> None:
    """탐색기(Finder)로 폴더를 연다. 실패하면 OSError — 화면은 대신 경로를 보여 준다."""
    folder.mkdir(parents=True, exist_ok=True)
    cmd = open_command(folder, select)
    if cmd is None:                      # Windows, 파일 지정 없음
        os.startfile(str(folder))        # type: ignore[attr-defined]  (Windows 전용 함수)
        return
    subprocess.Popen(cmd, stdin=subprocess.DEVNULL,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
