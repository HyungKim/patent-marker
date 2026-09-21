"""
tests/test_memory.py ─ 검토 학습(문단 기억 · 제외 사전/자동 규칙 · 유사 사례 검색 · 스위치) 점검 (가짜 Ollama, 몇 초)
=====================================================================
검토 데이터(dataset.jsonl)를 임시 폴더에 직접 써 놓고, 세 장치가 "똑같은 표현" 만이 아니라
숫자·어순·문장이 조금 다른 문단에서도 동작하는지, 그리고 스위치를 끄면 전부 멈추는지 확인합니다.

    실행:  python tests/test_memory.py
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

_TMP = tempfile.mkdtemp()
os.environ["PM_REVIEW_DIR"] = str(Path(_TMP) / "review_data")

import mock_llm  # noqa: E402

from app import config  # noqa: E402

config.OLLAMA_HOST = mock_llm.start()

from app import analyze, evaluate, lexicon, memory, merge  # noqa: E402

analyze.config.OLLAMA_HOST = config.OLLAMA_HOST
FAILS = 0


def check(label: str, ok: bool, detail: str = "") -> None:
    global FAILS
    FAILS += 0 if ok else 1
    print(f"{'PASS' if ok else 'FAIL'}  {label}" + (f"  ({detail})" if detail else ""))


def seg(i: int, text: str):
    return SimpleNamespace(seg_id=i, slide_no=1, kind="body", text=text, markable=True)


def mk(seg_, quote, span, grade, risk, category, reason):
    return analyze.Finding(seg_id=seg_.seg_id, slide_no=seg_.slide_no, quote=quote, grade=grade,
                           category=category, implicit=category in config.IMPLICIT_CATEGORIES,
                           disclosure_risk=risk, reason=reason, source="memory", span=span)


def llm(seg_, quote, grade="B", category="효과만기재", risk=False):
    span = merge._find_span(seg_.text, quote)
    return analyze.Finding(seg_id=seg_.seg_id, slide_no=1, quote=quote, grade=grade, category=category,
                           implicit=True, disclosure_risk=risk, reason="모델", source="llm", span=span)


# ── 검토 데이터 (사람이 확정한 교정 내역) ──
P1 = "조명 입사각을 22~28° 구간에서 가변한 결과 미세 크랙 대비비가 최대. 24°를 표준 사양으로 확정하고 각도 가변 마운트를 자체 제작."
P2 = "전극 검사장비 매출 비중이 34% → 51%로 상승. 조립·물류 장비 중심 구조에서 검사 솔루션 중심으로 이동 중."
P3 = "코팅 두께 편차 ±0.8㎛ → ±0.3㎛로 축소. 슬롯다이 제어기에 200ms 주기로 반영하는 폐루프 구성."
RECS = [
    {"date": "2026-09-20", "file": "a.pptx", "type": "match", "text": P1, "quote": "22~28° 구간", "grade": "A", "risk": False, "category": "수치·범위한정"},
    {"date": "2026-09-20", "file": "a.pptx", "type": "miss", "text": P1, "quote": "각도 가변 마운트를 자체 제작", "grade": "B", "risk": False, "category": "독자성주장"},
    {"date": "2026-09-20", "file": "a.pptx", "type": "fp", "text": P2, "quote": "매출 비중이 34% → 51%로 상승", "grade": "B", "risk": False, "category": "비교우위"},
    {"date": "2026-09-20", "file": "a.pptx", "type": "gold", "text": P3, "quote": "편차 ±0.8㎛ → ±0.3㎛로 축소", "grade": "A", "risk": False, "category": "수치·범위한정"},
]
Path(config.REVIEW_DIR).mkdir(parents=True, exist_ok=True)
with (config.REVIEW_DIR / "dataset.jsonl").open("w", encoding="utf-8") as fh:
    for r in RECS:
        fh.write(json.dumps(r, ensure_ascii=False) + "\n")

ix = memory.index()
st = memory.stats()
check("색인: 문단 3 · 자동 규칙 3 · 제외 1 · 사례 4", (st["paras"], st["rules"], st["excludes"], st["examples"]) == (3, 3, 1, 4), str(st))

# ── 글자 다루기 ──
check("핵심어 추출: 조사 떼고 두 글자 이상", memory.keywords("각도 가변 마운트를 자체 제작") == ["마운트", "각도", "가변", "자체"][:4]
      or set(memory.keywords("각도 가변 마운트를 자체 제작")) >= {"마운트", "자체", "제작"}, str(memory.keywords("각도 가변 마운트를 자체 제작")))
pat = memory.phrase_pattern("원가 40% 절감")
check("표현 패턴: 숫자·공백이 달라도 맞음", bool(pat.search("원가 27 % 절감")) and not pat.search("원가 절감"))
check("유사도: 숫자만 바뀐 문단은 0.9 이상", memory.similarity(P3, P3.replace("±0.3㎛", "±0.2㎛").replace("200ms", "150ms")) >= 0.9)
check("유사도: 전혀 다른 문단은 낮음", memory.similarity(P1, "감사합니다. 질문 있으시면 말씀해 주세요.") < 0.3)
check("찾기: 숫자가 바뀐 인용구도 자리를 찾음", memory.locate("코팅 두께 편차 ±0.9㎛ → ±0.2㎛로 축소함.", "편차 ±0.8㎛ → ±0.3㎛로 축소") is not None)

check("찾기: 문단 전체 인용구는 단어가 몇 개 바뀌어도 조각의 합으로 자리를 잡음",
      memory.locate("조명 입사각을 20~25° 구간에서 가변한 결과 미세 크랙 대비비가 최대. 22°를 표준 사양으로 결정하고 각도 가변 마운트를 자체 설계·제작.",
                    P1) == (0, 78) or memory.locate("조명 입사각을 20~25° 구간에서 가변한 결과 미세 크랙 대비비가 최대. 22°를 표준 사양으로 결정하고 각도 가변 마운트를 자체 설계·제작.", P1) is not None)

# ── ① 문단 기억: 같은 문단 / 숫자가 바뀐 문단 ──
memory.set_context(True)
S_same = seg(1, P1)
S_var = seg(2, P3.replace("±0.3㎛", "±0.2㎛").replace("200ms", "150ms"))
S_new = seg(3, "2026년 3분기 InterBattery 부스에서 공개 시연 예정")
deck = SimpleNamespace(segments=[S_same, S_var, S_new])
found = [llm(S_same, "표준 사양으로 확정", "B", "최적화·조건확립"),      # 사람은 이 문단에서 다른 두 구간만 확정했다
         llm(S_var, "폐루프 구성", "A", "구성·구조"),
         llm(S_new, "공개 시연", "B", "공개이력", risk=True)]
out, mst = memory.apply_memory(deck, found, mk)
q_same = sorted(f.quote for f in out if f.seg_id == 1)
q_var = [f for f in out if f.seg_id == 2]
check("기억: 같은 문단은 사람 판정으로 통째로 교체", q_same == ["22~28° 구간", "각도 가변 마운트를 자체 제작"], str(q_same))
check("기억: 숫자가 바뀐 문단도 기억이 적용되고 등급 A 유지", len(q_var) == 1 and q_var[0].grade == "A" and q_var[0].source == "memory"
      and "±0.2㎛" in q_var[0].quote, str([(f.quote, f.grade) for f in q_var]))
check("기억: 처음 보는 문단은 모델 답 유지", [f.quote for f in out if f.seg_id == 3] == ["공개 시연"])
check("기억: 집계", (mst["paras"], mst["added"], mst["dropped"]) == (2, 3, 2), str(mst))
# 정답 자리를 하나도 못 찾는 경우(문단은 비슷하지만 표현이 크게 바뀜)엔 모델 답을 지우지 않는다
S_far = seg(9, "코팅 두께 편차를 크게 줄였다. 슬롯다이 제어기에 150ms 주기로 반영하는 폐루프 구성을 유지함.")
if memory.match_para(S_far.text) is not None:
    out9, st9 = memory.apply_memory(SimpleNamespace(segments=[S_far]), [llm(S_far, "폐루프 구성", "A", "구성·구조")], mk)
    check("기억 안전장치: 정답 자리 못 찾으면 모델 답 유지", [f.quote for f in out9] == ["폐루프 구성"] or st9["added"] > 0, str(st9))

# ── ② 제외 사전: 지운 표현을 또 잡으면 뺀다 (숫자가 달라도) ──
S_biz = seg(4, "전극 검사장비 매출 비중이 38% → 55%로 상승했다.")
kept, n_ex = memory.apply_excludes([llm(S_biz, "매출 비중이 38% → 55%로 상승", "B", "비교우위"),
                                    llm(S_new, "공개 시연", "B", "공개이력", risk=True)])
check("제외: 지운 표현(숫자 변형)은 빠지고 공개 항목은 남음", n_ex == 1 and [f.quote for f in kept] == ["공개 시연"])

# ── ② 자동 규칙: 추가한 표현이 어순·숫자가 바뀌어도 규칙 사전에 걸린다 ──
h1 = [h for h in lexicon.scan("이번에는 자체 제작한 각도 가변 마운트를 라인에 적용했다.") if h.rid.startswith("AUTO_")]
h2 = [h for h in lexicon.scan("두께 편차 ±1.1㎛ → ±0.4㎛로 축소") if h.rid.startswith("AUTO_")]
h3 = [h for h in lexicon.scan("감사합니다. 질문 있으시면 말씀해 주세요.") if h.rid.startswith("AUTO_")]
check("자동 규칙: 어순이 바뀐 문장에 걸림 (핵심어 동시 출현)", len(h1) == 1 and h1[0].category == "독자성주장", str([(h.rid, h.matched) for h in h1]))
check("자동 규칙: 숫자가 바뀐 표현에 걸림 (패턴)", len(h2) == 1 and h2[0].category == "수치·범위한정", str([(h.rid, h.matched) for h in h2]))
check("자동 규칙: 무관한 문장에는 안 걸림", h3 == [])
check("자동 규칙: 지워진 표현(매출 비중)은 규칙이 되지 않음", not any("매출" in r.quote for r in ix.rules))

# ── ②-보강: 확정 표현이 걸린 구간이 아무 후보로도 안 덮였으면 구간 단위로 넣는다 ──
S_mix = seg(5, "이번 분기에는 각도 가변 마운트를 자체 제작해 적용했고, 조명 입사각을 20~25° 구간에서 가변했다.")
hits5 = lexicon.scan(S_mix.text)
have = [llm(S_mix, "자체 제작", "B", "독자성주장")]                 # 모델은 자체 제작만 잡았다
out5, n_rs = memory.rescue_confirmed(SimpleNamespace(segments=[S_mix]), list(have), {5: hits5}, mk)
check("확정 표현 구제: 덮이지 않은 '20~25° 구간' 이 A 등급으로 추가됨",
      n_rs == 1 and any("20~25°" in f.quote and f.grade == "A" for f in out5), str([(f.quote, f.grade) for f in out5]))
out5b, n_rs2 = memory.rescue_confirmed(SimpleNamespace(segments=[S_mix]), list(out5), {5: hits5}, mk)
check("확정 표현 구제: 이미 덮인 구간은 다시 넣지 않음", n_rs2 == 0)

# ── ③ 유사 사례: 비슷한 문단의 사례가 먼저 뽑힌다 (글자 겹침 / 가짜 임베딩) ──
blk, info = memory.examples_for(["두께 편차 ±1.1㎛ → ±0.4㎛로 축소한 공정 조건"], use_embed=False, k_pos=1, k_neg=1)
check("유사 사례(글자): 편차 사례가 1순위, 반례도 포함", info["mode"] == "ngram" and "±0.8㎛" in blk and "매출 비중" in blk, blk.replace("\n", " | ")[:160])
blk2, info2 = memory.examples_for(["두께 편차 ±1.1㎛ → ±0.4㎛로 축소한 공정 조건"], use_embed=True, k_pos=1, k_neg=1)
check("유사 사례(임베딩): 가짜 Ollama 의 /api/embed 를 거쳐 같은 결과", info2["mode"] == "embed" and "±0.8㎛" in blk2, blk2.replace("\n", " | ")[:120])
check("임베딩 캐시 파일 생성", (config.REVIEW_DIR / "embed_cache.json").exists())
check("bge-m3 설치 여부 판정 (가짜 서버는 있음)", memory.embed_available() is True)

# ── leave-one-out: 채점 때 자기 자신에서 나온 기억·규칙·사례는 뺀다 ──
memory.set_context(True, exclude_texts={P1})
check("leave-one-out: 자기 문단은 기억되지 않음", memory.match_para(P1) is None)
check("leave-one-out: 자기 문단에서 나온 규칙은 안 걸림", not any(h.rid.startswith("AUTO_") for h in lexicon.scan(P1)))
blk3, _ = memory.examples_for([P1], use_embed=False)
check("leave-one-out: 자기 문단의 사례는 주입되지 않음", "각도 가변 마운트" not in blk3)

# ── 스위치: 끄면 전부 멈춘다 ──
memory.set_context(False)
out_off, mst_off = memory.apply_memory(deck, found, mk)
check("끔: 문단 기억 없음", out_off == found and mst_off["paras"] == 0)
check("끔: 제외 사전 없음", memory.apply_excludes([llm(S_biz, "매출 비중이 38% → 55%로 상승")])[1] == 0)
check("끔: 자동 규칙 없음", not any(h.rid.startswith("AUTO_") for h in lexicon.scan("자체 제작한 각도 가변 마운트")))
check("끔: 유사 사례 없음", memory.examples_for([P3])[0] == "")
memory.clear_context()

# ── 파이프라인·성능 측정과의 연결 ──
from app import pipeline  # noqa: E402

check("RunOptions 에 learn/learn_embed 기본 켬", config.RunOptions().learn is True and config.RunOptions().learn_embed is True)
check("학습 요약 문구", pipeline.learn_text({"on": True, "memory_paras": 2, "memory_added": 3, "memory_dropped": 2,
                                          "excluded": 1, "rescued": 4, "examples": 6, "mode": "embed"})
      == "기억 2문단(+3/−2) · 제외 1 · 확정 표현 구제 4 · 사례 6(뜻 기준)"
      and pipeline.learn_text({"on": False}) == "끔")
es = evaluate.build_eval_set()
check("성능 측정 문제지: 검토 문단 3 · 정답 3", (len(es["paras"]), es["golds"]) == (3, 3), str((len(es["paras"]), es["golds"])))
snap = evaluate._config_snapshot()
check("성능 측정 스냅샷에 기억 크기 기록", snap.get("memory_paras") == 3 and snap.get("auto_rules") == 3 and "embed" in snap)

print(f"\n{'모두 통과' if not FAILS else f'{FAILS}건 실패'}")
sys.exit(1 if FAILS else 0)
