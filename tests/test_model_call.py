"""
tests/test_model_call.py ─ 모델 호출 경로 점검: 스트리밍 · 시간 제한 없음 · 중단 · 서버 감지 (몇 초)
=====================================================================
가짜 Ollama(mock_llm.py)를 느리게 만들어 놓고,
  - 답을 조각으로 받아도 온전히 이어 붙이는지
  - 기본값(PM_TIMEOUT=0)에서는 오래 걸려도 끊지 않는지
  - [중단] 신호가 모델 호출 도중에도 즉시 듣는지
  - PM_TIMEOUT 을 주면 '무응답' 만 끊는지
  - Ollama 프로세스가 사라진 것을 알아보는지
확인합니다.

    실행:  python tests/test_model_call.py
"""
from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import mock_llm  # noqa: E402

from app import config  # noqa: E402

MOCK_HOST = mock_llm.start()
config.OLLAMA_HOST = MOCK_HOST

from app import analyze, pipeline  # noqa: E402
from app.extract import Segment  # noqa: E402

FAILS = 0


def check(label: str, ok: bool, detail: str = "") -> None:
    global FAILS
    print(f"{'PASS' if ok else 'FAIL'}  {label}" + (f"  ({detail})" if detail else ""))
    if not ok:
        FAILS += 1


SEGS = [Segment(seg_id=1, slide_no=1, kind="body", addr="t/1", text="자체 제작 지그로 조립 시간 40% 절감"),
        Segment(seg_id=2, slide_no=1, kind="body", addr="t/2", text="450nm 청색광과 660nm 적색광을 200Hz로 교번 점등")]
OPTS = config.RunOptions()


def call(cancel=None, progress=None):
    return analyze.analyze_slide("테스트", 1, 1, SEGS, {}, OPTS, cancel=cancel, progress=progress)


# ── 1. 스트리밍 조각을 이어 붙인다 ─────────────────────────────────
mock_llm.Handler.delay = 0
config.REQUEST_TIMEOUT = 0
seen: list[int] = []
got = call(progress=seen.append)
check("조각으로 온 답을 이어 붙여 후보를 만든다", len(got) >= 2, f"{len(got)}건: {[f.quote for f in got]}")
check("진행 콜백이 받은 글자 수를 알린다 (마지막 값 > 0)", bool(seen) and seen[-1] > 0, str(seen[-3:]))

# ── 2. 시간 제한 없음: 느려도 끊지 않는다 ───────────────────────────
mock_llm.Handler.delay = 2.5
t = time.time()
got = call()
check("PM_TIMEOUT=0: 2.5초 넘게 걸려도 끊지 않고 답을 받는다", len(got) >= 2 and time.time() - t >= 2.4,
      f"{time.time() - t:.1f}초")

# ── 3. 중단 신호가 호출 도중에 즉시 듣는다 ──────────────────────────
mock_llm.Handler.delay = 8
ev = threading.Event()
threading.Timer(0.8, ev.set).start()
t = time.time()
try:
    call(cancel=ev)
    check("중단 → Aborted", False, "예외가 없었다")
except analyze.Aborted:
    took = time.time() - t
    check("중단 신호 0.8초 뒤 → 1.5초 안에 Aborted (8초 지연을 기다리지 않음)", took < 2.5, f"{took:.1f}초")

# pipeline.run 도 Cancelled 로 바꿔 올린다
ev2 = threading.Event()
threading.Timer(0.8, ev2.set).start()
try:
    pipeline.run(ROOT / "samples" / "회사보고자료_예시.pptx", ROOT / "output" / "_취소테스트.pptx",
                 OPTS, cancel=ev2)
    check("pipeline: 호출 도중 중단 → Cancelled", False, "예외가 없었다")
except pipeline.Cancelled:
    check("pipeline: 호출 도중 중단 → Cancelled", not (ROOT / "output" / "_취소테스트.pptx").exists())

# ── 4. PM_TIMEOUT 을 주면 '무응답' 만 끊는다 ────────────────────────
mock_llm.Handler.delay = 3
config.REQUEST_TIMEOUT = 1
t = time.time()
try:
    call()
    check("PM_TIMEOUT=1 · 3초 무응답 → OllamaError", False, "예외가 없었다")
except analyze.OllamaError as e:
    check("PM_TIMEOUT=1 · 3초 무응답 → OllamaError(무응답 안내)", "아무 응답도" in str(e) and time.time() - t < 2.5,
          str(e)[:50])
config.REQUEST_TIMEOUT = 0
mock_llm.Handler.delay = 0

# ── 5. 서버 생존 확인 ──────────────────────────────────────────────
check("살아 있는 서버는 '사라짐' 이 아니다", analyze._server_gone() is False)
saved = config.OLLAMA_HOST
config.OLLAMA_HOST = "http://127.0.0.1:1"          # 아무도 듣지 않는 포트 → 연결 거부
check("연결 거부 = 서버 사라짐", analyze._server_gone() is True)
config.OLLAMA_HOST = saved

# 감시 스레드가 실제로 끊는지: 호출 중에 주소를 죽은 포트로 바꿔 '서버 사라짐' 을 흉내 낸다
old_every, old_fails = analyze._PING_EVERY, analyze._PING_FAILS
analyze._PING_EVERY, analyze._PING_FAILS = 0.3, 2
mock_llm.Handler.delay = 8
t = time.time()
threading.Timer(0.5, lambda: setattr(config, "OLLAMA_HOST", "http://127.0.0.1:1")).start()
try:
    call()
    check("서버 사라짐 감지 → OllamaError", False, "예외가 없었다")
except analyze.OllamaError as e:
    check("서버 사라짐 감지 → 8초 지연을 기다리지 않고 OllamaError", "사라져" in str(e) and time.time() - t < 4,
          f"{time.time() - t:.1f}초")
finally:
    config.OLLAMA_HOST = saved
    analyze._PING_EVERY, analyze._PING_FAILS = old_every, old_fails
    mock_llm.Handler.delay = 0

print("\n" + ("모두 통과" if not FAILS else f"실패 {FAILS}건"))
sys.exit(1 if FAILS else 0)
