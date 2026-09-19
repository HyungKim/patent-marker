"""
test_eval.py ─ 성능 측정 경로 점검 (모델 없이 몇 초 만에)
=====================================================================

evaluate.py 의 전 과정을 가짜 모델로 검사합니다.
  1) dataset.jsonl 의 확정 기록 → 문제지(build_eval_set) 변환과 나중 기록 우선 규칙
  2) 최초/현재 설정을 같은 문제지로 채점했을 때의 재현율·정밀도·F1·등급일치 계산
  3) 항목 단위 개선 집계 (놓치던 것을 잡음 · 오탐 제거 · 퇴행)
  4) 이력 누적과 "직전 대비 변경 내역" (사전 추가가 changes 에 나타나는지)

실행:  .venv/bin/python tests/test_eval.py     (Windows: .venv\\Scripts\\python)
"""
import json
import os
import sys
import tempfile
from pathlib import Path

# Windows 명령 프롬프트는 글자를 cp949 로 내보내는데, 아래 출력에 쓰는 '—' 같은
# 글자가 cp949 에 없어서 결과를 파일로 넘기면(> log.txt) 도중에 죽어 버린다.
# 검사 내용과는 무관한 사고라, 출력만 utf-8 로 고정해 둔다.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:      # noqa: BLE001  (예전 파이썬·특수 환경에서는 그냥 넘어간다)
        pass

# 검토 데이터가 실제 review_data/ 를 건드리지 않도록, import 전에 임시 폴더로 돌린다
_TMP = Path(tempfile.mkdtemp(prefix="pm-evaltest-"))
os.environ["PM_REVIEW_DIR"] = str(_TMP)

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import analyze, config, evaluate, lexicon, review  # noqa: E402

# ── 문제지 재료: 규칙 사전에 걸리지 않는 중립 문장 3개 ──────────────
# (사전이 걸리면 안전망(구제)이 끼어들어 채점 수학을 검증하기 어려워진다)
P_A = "부서 대항 체육 행사 일정과 준비물 안내"
P_B = "신년 하례식 참석 대상 부서 명단"
P_C = "하계 워크숍 조별 편성 결과 공지"
for t in (P_A, P_B, P_C):
    assert lexicon.scan(t) == [], f"중립 문장이 규칙 사전에 걸림: {t}"

# ── 1. 검토 반영이 쌓은 것처럼 dataset.jsonl 을 만든다 ──────────────
records = [
    # pA: 사람이 새로 칠한 누락(A) + 모델이 맞게 잡았던 일치(B)
    {"type": "miss", "text": P_A, "quote": "체육 행사 일정", "grade": "A", "risk": False},
    {"type": "match", "text": P_A, "quote": "준비물 안내", "grade": "B", "risk": False},
    # pB: 사람이 지운 오탐만 있는 문단
    {"type": "fp", "text": P_B, "quote": "참석 대상 부서"},
    # pC: 먼저 오탐으로 기록됐다가 나중 검토에서 확정 — 나중 기록이 이겨야 한다
    {"type": "fp", "text": P_C, "quote": "조별 편성"},
    {"type": "gold", "text": P_C, "quote": "조별 편성", "grade": "B", "risk": False},
    # 원문에 없는 인용구는 건너뛰어야 한다
    {"type": "gold", "text": P_A, "quote": "없는 인용구", "grade": "B", "risk": False},
]
_TMP.mkdir(exist_ok=True)
with (_TMP / "dataset.jsonl").open("w", encoding="utf-8") as fh:
    for r in records:
        fh.write(json.dumps({"date": "2026-09-19", "file": "검토완료.pptx",
                             "mode": "archive", **r}, ensure_ascii=False) + "\n")

es = evaluate.build_eval_set()
assert len(es["paras"]) == 3, es
assert es["golds"] == 3, es                      # pA 2개 + pC 1개 (없는 인용구 제외)
assert es["neg_only"] == 1, es                   # pB
by_text = {p["text"]: p for p in es["paras"]}
assert len(by_text[P_C]["golds"]) == 1 and not by_text[P_C]["negs"], "나중 기록 우선 실패"
assert len(by_text[P_B]["negs"]) == 1 and not by_text[P_B]["golds"]
print("1. 문제지 변환 OK — 문단 3 · 정답 3 · 나중 기록 우선 · 원문에 없는 인용구 제외")

# ── 2. 가짜 모델: 최초 설정은 일부만 맞히고, 현재 설정은 다 맞힌다 ──
# 현재 설정 프롬프트가 최초와 달라지도록 확정 사례 하나를 심는다 (실사용과 동일)
(_TMP / "examples.json").write_text(json.dumps(
    [{"kind": "include", "quote": "조별 편성", "grade": "B", "risk": False,
      "snippet": ""}], ensure_ascii=False), encoding="utf-8")
assert analyze._system_prompt() != analyze.SYSTEM

BEHAVIOR = {
    # 최초 설정: pA 의 정답 하나를 놓치고, pB 에서 오탐을 내고, pC 는 등급을 틀린다
    True: {P_A: [("준비물 안내", "B")],
           P_B: [("참석 대상 부서", "B")],
           P_C: [("조별 편성", "A")]},
    # 현재 설정: 전부 맞힌다
    False: {P_A: [("체육 행사 일정", "A"), ("준비물 안내", "B")],
            P_B: [],
            P_C: [("조별 편성", "B")]},
}


def _fake_batch(deck_title, slide_no, total, segs, hints, opts, system_override=None):
    is_base = system_override == analyze.SYSTEM
    out = []
    for s in segs:
        for quote, grade in BEHAVIOR[is_base].get(s.text, []):
            out.append(analyze.Finding(
                seg_id=s.seg_id, slide_no=s.slide_no, quote=quote, grade=grade,
                category="구성·구조", implicit=False, disclosure_risk=False,
                reason="테스트", source="llm"))
    return out


analyze._analyze_batch = _fake_batch
analyze.health = lambda: {"ok": True, "models": ["qwen3:8b"]}

entry = evaluate.run_sync(note="테스트 1회차")

b, c, d = entry["baseline"], entry["current"], entry["delta"]
assert (b["tp"], b["miss"], b["fp"]) == (2, 1, 1), b
assert b["recall"] == 66.7 and b["precision"] == 66.7 and b["f1"] == 66.7, b
assert b["grade_acc"] == 50.0, b                 # pC 등급을 A 로 틀림 (정답 B)
assert (c["tp"], c["miss"], c["fp"]) == (3, 0, 0), c
assert c["recall"] == 100.0 and c["precision"] == 100.0 and c["f1"] == 100.0, c
assert c["grade_acc"] == 100.0, c
assert d == {"recall": 33.3, "precision": 33.3, "f1": 33.3}, d
print("2. 채점 OK — 최초 66.7/66.7/66.7(등급 50) → 현재 100/100/100 · Δ33.3%p")

fx = entry["fixes"]
assert fx["miss_fixed"] == 1 and fx["miss_fixed_quotes"] == ["체육 행사 일정"], fx
assert fx["fp_fixed"] == 1 and fx["fp_fixed_quotes"] == ["참석 대상 부서"], fx
assert fx["regressed"] == 0 and fx["new_fp"] == 0, fx
assert fx["still_miss_quotes"] == [], fx
assert entry["run_no"] == 1 and entry["changes"] is None
assert entry["baseline"]["fallback"] is False
print("3. 항목 단위 개선 집계 OK — 누락 해결 1 · 오탐 제거 1 · 퇴행 0")

# ── 3. 사전 추가 후 2회차: changes 에 무엇이 바뀌었는지 남아야 한다 ──
review.add_rule("체육 행사")
entry2 = evaluate.run_sync(note="테스트 2회차")
ch = entry2["changes"]
assert entry2["run_no"] == 2
assert ch["rules_added"] == ["체육 행사"], ch
assert ch["model_change"] is None and ch["rules_removed"] == [], ch
assert entry2["current"]["user_rules"] == 1
assert ch["injected_change"] is None, ch          # 주입 칸이 그대로면 표시 없음
hist = evaluate.history()
assert [e["run_no"] for e in hist] == [1, 2]
print("4. 이력 누적 OK — 2회차 changes 에 '사전 +1 (체육 행사)' 기록됨")

# ── 3.5 주입 칸 수 변화 표시 ──────────────────────────────────────
# 회차 사이에 프롬프트 칸 수가 바뀌면 점수 변화의 원인이 '교정이 좋아져서' 가
# 아닐 수 있다. 특히 이 값을 기록하지 않던 예전 회차(도구 업그레이드 직전)는
# 프롬프트가 가장 크게 바뀐 회차인데 '변경 없음' 으로 보이면 안 된다.
assert evaluate._injected_change(8, 12) == "8 → 12"
assert evaluate._injected_change(12, 12) is None
assert evaluate._injected_change(None, 12) == "기록 없음 → 12"
print("5. 주입 칸 변화 표시 OK — 기록이 없던 회차도 '기록 없음 → N' 으로 드러남")

# 사용자 규칙을 되돌린다 (전역 RULES 를 다른 테스트가 이어받아도 안전하도록)
(_TMP / "extra_rules.json").unlink()
lexicon.load_user_rules()

print("\n전부 통과 — 성능 측정 경로 정상")
