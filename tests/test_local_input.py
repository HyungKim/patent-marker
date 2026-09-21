"""
tests/test_local_input.py ─ 브라우저 업로드 없이 파일 넣기 경로 점검 (LLM 없이 몇 초)
=====================================================================
input 폴더 목록 · 경로 정리 · 결과 파일 이름 · 화면 토큰 · 파일 선택창 응답 해석 ·
파이프라인(pipeline.py) · 명령행(app/cli.py) 이 끊기지 않았는지 확인합니다.
모델 자리는 tests/mock_llm.py 의 가짜 Ollama 가 대신합니다.

    실행:  python tests/test_local_input.py
"""
from __future__ import annotations

import os
import subprocess
import sys
import json
import tempfile
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import mock_llm  # noqa: E402  (가짜 Ollama)

from app import config  # noqa: E402

MOCK_HOST = mock_llm.start()
config.OLLAMA_HOST = MOCK_HOST

from app import analyze, cli, local, pipeline  # noqa: E402

analyze.config.OLLAMA_HOST = MOCK_HOST

SAMPLE = ROOT / "samples" / "회사보고자료_예시.pptx"
FAILS = 0


def check(label: str, ok: bool, detail: str = "") -> None:
    global FAILS
    print(f"{'PASS' if ok else 'FAIL'}  {label}" + (f"  ({detail})" if detail else ""))
    if not ok:
        FAILS += 1


# ── 1. 경로 정리 ─────────────────────────────────────────────────
check("따옴표 벗기기 (탐색기 '경로로 복사')",
      local.clean_path('"C:\\보고\\파일.pptx"') == "C:\\보고\\파일.pptx")
check("둥근 따옴표·공백도 벗긴다",
      local.clean_path("  “C:\\a b\\c.pptx”\n") == "C:\\a b\\c.pptx")
check("file:/// (Windows) → 드라이브 경로",
      local.clean_path("file:///C:/Users/%ED%99%8D/a.pptx") == "C:/Users/홍/a.pptx")
check("file:/// (macOS) → 절대 경로",
      local.clean_path("file:///Users/kim/a.pptx") == "/Users/kim/a.pptx")

# ── 2. 파일 목록·경로 확정·결과 이름 (임시 폴더를 input/output 으로) ─
with tempfile.TemporaryDirectory() as td:
    tmp = Path(td)
    config.INPUT_DIR = tmp / "input"
    config.OUTPUT_DIR = tmp / "output"
    config.REVIEW_DIR = tmp / "review_data"        # 실행 기록(run_log.tsv)도 임시 폴더에
    local.ensure_dirs()
    check("ensure_dirs 가 input·output 을 만든다",
          config.INPUT_DIR.is_dir() and config.OUTPUT_DIR.is_dir())

    older = config.INPUT_DIR / "옛날.pptx"
    older.write_bytes(SAMPLE.read_bytes())
    os.utime(older, (time.time() - 100, time.time() - 100))
    newer = config.INPUT_DIR / "최근.PPTX"
    newer.write_bytes(SAMPLE.read_bytes())
    (config.INPUT_DIR / "~$최근.PPTX").write_bytes(b"lock")     # PowerPoint 잠금 파일
    (config.INPUT_DIR / "메모.txt").write_text("x", encoding="utf-8")
    (config.INPUT_DIR / "예약중.pptx").write_bytes(b"")           # 분석 중 예약된 빈 파일
    names = [r["name"] for r in local.list_files(config.INPUT_DIR)]
    check("목록: pptx 만, 최근 수정 순, 잠금·빈 파일 제외", names == ["최근.PPTX", "옛날.pptx"], str(names))
    check("목록 항목에 size·mtime·path 가 있다",
          all(k in local.list_files(config.INPUT_DIR)[0] for k in ("size", "mtime", "path")))
    check("없는 폴더는 빈 목록", local.list_files(tmp / "없음") == [])

    check("경로 확정: 절대 경로", local.resolve_pptx(str(older)) == older.resolve())
    check("경로 확정: 파일 이름만 적으면 input 폴더에서 찾는다",
          local.resolve_pptx("옛날.pptx") == older.resolve())
    check("경로 확정: 따옴표 붙은 경로", local.resolve_pptx(f'"{older}"') == older.resolve())
    for raw, why in [(str(tmp / "없는파일.pptx"), "없는 파일"), (str(config.INPUT_DIR), "폴더"),
                     (str(config.INPUT_DIR / "메모.txt"), "pptx 아님"), ("", "빈 문자열")]:
        try:
            local.resolve_pptx(raw)
            check(f"경로 확정 거부: {why}", False, "예외가 없었다")
        except ValueError as e:
            check(f"경로 확정 거부: {why}", True, str(e)[:40])

    p1 = local.output_path_for("보고서.pptx")
    check("결과 이름: 이름_특허마킹.pptx", p1 == config.OUTPUT_DIR / "보고서_특허마킹.pptx")
    p1.write_bytes(b"x")
    p2 = local.output_path_for("보고서.pptx")
    p2.write_bytes(b"x")
    p3 = local.output_path_for("보고서.pptx")
    check("결과 이름: 겹치면 (2), (3)", (p2.name, p3.name) == ("보고서_특허마킹(2).pptx", "보고서_특허마킹(3).pptx"))
    check("결과 이름: 다른 폴더 지정", local.output_path_for("a.pptx", tmp / "elsewhere").parent == tmp / "elsewhere")

    # ── 3. 파이프라인 (가짜 모델) + 중단 신호 ────────────────────
    stages: list[str] = []
    out = config.OUTPUT_DIR / "파이프라인.pptx"
    resolved, stats = pipeline.run(SAMPLE, out, config.RunOptions(),
                                   progress=lambda p: stages.append(p.stage))
    check("pipeline.run: 결과 파일 저장", out.exists() and out.stat().st_size > 1000)
    check("pipeline.run: 후보와 집계", stats["total"] == len(resolved) > 0 and stats["legend"] == 1,
          f"후보 {stats['total']} · 범례 {stats['legend']}")
    check("pipeline.run: 단계 보고 (읽기→…→완료)",
          stages[0] == "문서 읽는 중" and stages[-1] == "완료" and any("슬라이드 1 " in s for s in stages))
    # ── 3-1. 속도 성적표: 슬라이드 줄 · 집계 · 실행 기록 파일 ────
    check("슬라이드 성적표 줄 (토큰·토큰/초)",
          any("분석 완료 ·" in st and "토큰/초" in st for st in stages),
          next((st for st in stages if "분석 완료 ·" in st), ""))
    check("집계에 토큰 수·속도·버전", stats["prompt_tokens"] > 0 and stats["output_tokens"] > 0
          and stats["tokens_per_sec"] > 0 and stats["version"] == config.VERSION and "토큰/초" in stats["speed_text"],
          stats.get("speed_text", ""))
    check("모델 호출 수 = 문단이 있는 슬라이드 수", stats["calls"] == 5, f"calls={stats['calls']}")
    log = config.REVIEW_DIR / "run_log.tsv"
    lines = log.read_text(encoding="utf-8-sig").splitlines() if log.exists() else []
    check("run_log.tsv 에 머리글 + 한 줄", len(lines) == 2 and lines[0].startswith("일시\t버전") and SAMPLE.name in lines[1],
          lines[1][:80] if len(lines) > 1 else "없음")
    check("run_log.tsv 칸 수 = 머리글 칸 수", len(lines) == 2 and lines[1].count("\t") == lines[0].count("\t"))
    llm = [f for f in resolved if f.source == "llm"]
    check("모델 답의 묵시 여부는 분류로 정해진다",
          llm and all(f.implicit == (f.category in config.IMPLICIT_CATEGORIES) for f in llm), f"llm {len(llm)}건")
    check("모델 답의 사유가 짧게 들어온다", llm and all(len(f.reason) <= 12 for f in llm))
    new_item = analyze._norm_item({"i": 3, "q": "x", "g": "B", "c": "효과만기재", "d": True, "r": "근거"})
    old_item = analyze._norm_item({"seg_id": 3, "quote": "x", "grade": "B", "category": "효과만기재",
                                   "implicit": False, "disclosure_risk": True, "reason": "근거"})
    check("답 형식: 새(i·q·g·c·d·r)와 옛(seg_id…) 키를 모두 받는다",
          new_item == {"seg_id": 3, "quote": "x", "grade": "B", "category": "효과만기재", "disclosure_risk": True, "reason": "근거"}
          and old_item == {**new_item, "implicit": False})
    ev = threading.Event()
    ev.set()
    try:
        pipeline.run(SAMPLE, config.OUTPUT_DIR / "중단.pptx", config.RunOptions(), cancel=ev)
        check("pipeline.run: 중단 신호", False, "Cancelled 가 나지 않았다")
    except pipeline.Cancelled:
        check("pipeline.run: 중단 신호 → Cancelled", not (config.OUTPUT_DIR / "중단.pptx").exists())

    # ── 4. 명령행 (같은 프로세스) ─────────────────────────────────
    outdir = tmp / "cli_out"
    rc = cli.main([str(older), "--out", str(outdir)])
    check("cli: 정상 종료 코드 0", rc == 0)
    check("cli: --out 폴더에 이름_특허마킹.pptx", (outdir / "옛날_특허마킹.pptx").exists())
    src_dir = tmp / "원본자리"
    src_dir.mkdir()
    here = src_dir / "현장.pptx"
    here.write_bytes(SAMPLE.read_bytes())
    rc = cli.main([str(here)])
    check("cli: --out 없으면 원본 옆에 저장", rc == 0 and (src_dir / "현장_특허마킹.pptx").exists())
    rc = cli.main([str(older)])
    check("cli: input 폴더의 파일은 output 폴더에 저장 (목록에 결과가 섞이지 않게)",
          rc == 0 and (config.OUTPUT_DIR / "옛날_특허마킹.pptx").exists()
          and not (config.INPUT_DIR / "옛날_특허마킹.pptx").exists())
    rc = cli.main([str(tmp / "없음.pptx"), str(here)])
    check("cli: 틀린 파일은 건너뛰고 나머지는 처리, 종료 코드 1",
          rc == 1 and (src_dir / "현장_특허마킹(2).pptx").exists())

    # ── 5. 명령행 (mark.bat 이 부르는 방식 그대로: 별도 프로세스) ──
    env = {**os.environ, "PM_OLLAMA_HOST": MOCK_HOST, "PM_OUTPUT_DIR": str(config.OUTPUT_DIR),
           "PM_REVIEW_DIR": str(config.REVIEW_DIR)}      # 별도 프로세스도 실행 기록을 임시 폴더에
    proc = subprocess.run([sys.executable, "-m", "app.cli", str(here), "--out", str(tmp / "proc_out")],
                          cwd=str(ROOT), env=env, capture_output=True, text=True, encoding="utf-8")
    check("python -m app.cli (별도 프로세스)", proc.returncode == 0 and (tmp / "proc_out" / "현장_특허마킹.pptx").exists(),
          (proc.stdout + proc.stderr).strip().splitlines()[-1] if (proc.stdout + proc.stderr).strip() else "")

# ── 6. 화면 토큰 ─────────────────────────────────────────────────
from fastapi import HTTPException  # noqa: E402

from app import main as web  # noqa: E402

try:
    web.require_token("틀린값")
    check("토큰 없이는 403", False)
except HTTPException as e:
    check("토큰 없이는 403", e.status_code == 403)
try:
    web.require_token(None)
    check("헤더가 없어도 403", False)
except HTTPException as e:
    check("헤더가 없어도 403", e.status_code == 403)
check("맞는 토큰은 통과", web.require_token(web.TOKEN) is None)
check("화면에 토큰이 심어진다", web.TOKEN in web.index().body.decode("utf-8"))
_page = web.index()
_html = _page.body.decode("utf-8")
check("화면 제목 옆에 버전이 심어진다", config.VERSION in _html and "__PM_VERSION__" not in _html)
check("첫 화면은 캐시하지 않는다 (업데이트 뒤 옛 화면 방지)", _page.headers.get("cache-control") == "no-store")
# ── 정밀/빠름 모델 선택 ──
check("화면에 모델 선택 상자", 'id="optModel"' in _html and "chosenModel" in _html)
_h = json.loads(web.health().body)
_ch = _h.get("choices") or []
check("상태 응답에 선택지 둘(정밀·빠름)과 설치 여부",
      [c["label"] for c in _ch] == ["정밀", "빠름"] and _ch[0]["id"] == config.MODEL
      and _ch[1]["id"] == config.FAST_MODEL and all("installed" in c for c in _ch), str(_ch)[:120])
check("가짜 Ollama 에는 정밀만 설치됨으로 표시", _ch[0]["installed"] is True and _ch[1]["installed"] is False)
check("cli --fast 는 빠름 모델", cli._parse(["x.pptx", "--fast"]).model == config.FAST_MODEL
      and cli._parse(["x.pptx"]).model == config.MODEL)

# ── 7. 파일 선택창 응답 해석 · 열기 명령 ──────────────────────────
check("선택창 응답: 경로 목록", local._pick_result(0, '["C:\\\\a.pptx", "C:\\\\b.pptx"]\n', "") == ["C:\\a.pptx", "C:\\b.pptx"])
check("선택창 응답: 취소 → 빈 목록", local._pick_result(0, "[]\n", "") == [])
try:
    local._pick_result(1, "", "Traceback ...\nModuleNotFoundError: No module named '_tkinter'")
    check("선택창 응답: tkinter 없음 안내", False)
except RuntimeError as e:
    check("선택창 응답: tkinter 없음 안내", "tkinter" in str(e) and "input 폴더" in str(e))
dlg = subprocess.run([sys.executable, "-c", "import tkinter, tkinter.filedialog"], capture_output=True)
check("이 PC 의 Python 에 tkinter 가 있다 (pick_available 과 일치)",
      (dlg.returncode == 0) == local.pick_available())
check("pickdialog.py 문법", subprocess.run([sys.executable, "-m", "py_compile", str(ROOT / "app" / "pickdialog.py")]).returncode == 0)

folder, sel = Path("/tmp/x"), Path(__file__)
cmd = local.open_command(folder, sel)
if sys.platform == "win32":
    check("폴더 열기 명령 (Windows: 파일 선택)", cmd == ["explorer.exe", "/select,", str(sel)])
    check("폴더 열기 명령 (Windows: 폴더만 → os.startfile)", local.open_command(folder) is None)
elif sys.platform == "darwin":
    check("폴더 열기 명령 (macOS)", cmd == ["open", "-R", str(sel)] and local.open_command(folder) == ["open", str(folder)])
else:
    check("폴더 열기 명령 (Linux)", cmd == ["xdg-open", str(folder)])

print("\n" + ("모두 통과" if not FAILS else f"실패 {FAILS}건"))
sys.exit(1 if FAILS else 0)
