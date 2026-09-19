"""
tests/test_review.py ─ 검토 반영(사전·예시 보강) 경로 점검 (LLM 없이, 몇 초)
=====================================================================
가짜 문서를 만들어 마킹까지 한 뒤, "사람이 검토한 것처럼" 형광펜을 고치고
review.py 가 교정 내역(누락·오탐)을 제대로 읽어 내는지 확인합니다.

    실행:  python tests/test_review.py
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# review 데이터가 실제 review_data/ 를 건드리지 않도록, import 전에 임시 폴더로 돌린다
TMP = tempfile.mkdtemp(prefix="pm-review-test-")
os.environ["PM_REVIEW_DIR"] = TMP

from lxml import etree                              # noqa: E402
from pptx import Presentation                      # noqa: E402
from pptx.util import Inches                       # noqa: E402

from app import config, extract, lexicon, mark, merge, review  # noqa: E402
from app.analyze import Finding                    # noqa: E402

WORK = Path(tempfile.mkdtemp(prefix="pm-review-work-"))

# "크로노백형합금" 은 규칙 사전 어디에도 없는 낱말 → '사전 미포착' 검증용
P1 = "자체 개발 크로노백형합금을 적용해 불량률을 크게 낮춤"
P2 = "보고 절차를 간소화하여 결재 소요 시간 단축"      # 모델 오탐 역할 (사람이 지울 것)

fails = 0


def check(name: str, ok: bool, detail: str = ""):
    global fails
    fails += 0 if ok else 1
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail else ""))


# ── 1. 원본 문서 생성 ────────────────────────────────────────────
src = WORK / "src.pptx"
prs = Presentation()
layout = min(prs.slide_layouts, key=lambda l: len(l.placeholders))
slide = prs.slides.add_slide(layout)
for shp in list(slide.shapes):
    if shp.is_placeholder:
        shp._element.getparent().remove(shp._element)
tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(2))
tb.text_frame.text = P1
tb.text_frame.add_paragraph().text = P2
prs.save(str(src))

# ── 2. 도구가 분석한 것처럼: 판정 → 기록 저장 → 마킹 ────────────
deck = extract.extract(str(src))
seg1 = next(s for s in deck.segments if s.text == P1)
seg2 = next(s for s in deck.segments if s.text == P2)
model_findings = [
    Finding(seg1.seg_id, 1, "자체 개발", "B", "독자성주장", True, False, "독자 개발 주장"),
    Finding(seg2.seg_id, 1, "결재 소요 시간 단축", "B", "효과만기재", True, False, "효과만 기재"),
]
hits = {s.seg_id: lexicon.scan(s.text) for s in deck.segments}
resolved, marks = merge.resolve(deck, model_findings, hits)
fp = review.save_archive(deck, resolved)
# 문구 옵션을 켜서 마킹 — 6절의 '고아 문구로 오탐 추정' 경로까지 검사하기 위해
st_mark = mark.apply(deck, resolved, marks, add_summary=True, tag_marks=True)
marked = WORK / "marked.pptx"
deck.prs.save(str(marked))
check("분석 기록 저장", review.load_archive(fp) is not None)
check("첫 슬라이드에 색상 범례 상자", st_mark["legend"] == 1 and any(
    s.name == config.LEGEND_NAME for s in Presentation(str(marked)).slides[0].shapes))
check("배지 문구 = 출원검토필요", any(
    s.name == config.BADGE_NAME and s.text_frame.text.startswith("출원검토필요")
    for s in Presentation(str(marked)).slides[0].shapes))

# 기본 옵션(문구 없음)으로도 한 번: 범례의 색 견본이 검토본 형광펜으로 오인되지 않아야 한다
deck_b = extract.extract(str(src))
res_b, marks_b = merge.resolve(deck_b, [
    Finding(next(s.seg_id for s in deck_b.segments if s.text == P1), 1, "자체 개발", "B",
            "독자성주장", True, False, "독자 개발 주장"),
    Finding(next(s.seg_id for s in deck_b.segments if s.text == P2), 1, "결재 소요 시간 단축", "B",
            "효과만기재", True, False, "효과만 기재"),
], hits)
st_b = mark.apply(deck_b, res_b, marks_b, add_summary=True)
marked_b = WORK / "marked_default.pptx"
deck_b.prs.save(str(marked_b))
rv_b = review.read_reviewed(str(marked_b))
check("기본 옵션: 문구 없음", st_b["tags"] == 0 and not any(
    p["tags"] for p in rv_b["paras"]))
check("기본 옵션: 범례 견본은 형광펜으로 세지 않음 (본문 2구간만)",
      sum(len(p["spans"]) for p in rv_b["paras"]) == 2,
      str([p["spans"] for p in rv_b["paras"]]))
check("기본 옵션: 지문 일치", rv_b["fingerprint"] == fp)

# 예전 버전이 붙인 【특허검토필요】 문구도 도구 문구로 알아본다
_old = etree.fromstring(
    '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:r><a:t>자체 개발 구간</a:t></a:r><a:r><a:t> 【특허검토필요】</a:t></a:r></a:p>')
_txt, _sp, _tg = review._para_review(_old)
check("예전 문구 호환 (글자에서 제외·위치 기록)", _txt == "자체 개발 구간" and _tg == [8],
      f"{_txt!r} {_tg}")

# ── 3. "사람의 검토" 흉내: P2 형광펜 제거(오탐), P1 에 새 형광펜(누락) ──
from pptx.oxml.ns import qn                         # noqa: E402
from app.extract import _para_text                  # noqa: E402

rprs = Presentation(str(marked))
for shp in rprs.slides[0].shapes:
    if not getattr(shp, "has_text_frame", False) or not shp.has_text_frame:
        continue
    for para in shp.text_frame.paragraphs:
        text = _para_text(para._p)
        if "결재 소요" in text:                     # 오탐: 형광펜만 지운다 (문구는 남김)
            for rPr in para._p.iter(qn("a:rPr")):
                for hl in rPr.findall(qn("a:highlight")):
                    rPr.remove(hl)
        if "크로노백형합금" in text:                # 누락: 사람이 새로 칠한다
            i = text.find("크로노백형합금")
            mark.highlight(para._p, (i, i + len("크로노백형합금")),
                           config.GRADE_COLOR["B"], False)   # 문구는 붙이지 않음
reviewed_path = WORK / "reviewed.pptx"
rprs.save(str(reviewed_path))

# ── 4. 검토본 읽기 → 비교 ────────────────────────────────────────
rv = review.read_reviewed(str(reviewed_path))
check("지문 일치 (문구·배지·요약 제거 후)", rv["fingerprint"] == fp,
      f"{rv['fingerprint']} vs {fp}")

d = review.diff(rv)
c = d["counts"]
check("기록 대조 모드", d["mode"] == "archive")
check("일치 1건 (자체 개발)", c["match"] == 1, str(c))
check("누락 1건 (크로노백형합금)", c["miss"] == 1, str(c))
check("오탐 1건 (결재 단축)", c["fp"] == 1, str(c))
miss = next(i for i in d["items"] if i["type"] == "miss")
check("누락 인용구", miss["quote"] == "크로노백형합금", miss["quote"])
check("사전 미포착 표시", miss.get("lexicon_gap") is True)
fpi = next(i for i in d["items"] if i["type"] == "fp")
check("오탐 인용구에 원문 포함", "단축" in fpi["quote"], fpi["quote"])

# ── 5. 저장 → 예시 주입 → 사전 추가 ──────────────────────────────
st = review.save_dataset("test.pptx", d["mode"], d["items"])
check("데이터셋 누적", st["dataset"] == len(d["items"]), str(st))
block = review.examples_block()
check("예시 블록에 오탐 반례", "단축" in block)
check("예시 블록에 누락 정답", "크로노백형합금" in block)

review.add_rule("크로노백형합금")
hits2 = lexicon.scan("신규 크로노백형합금 적용 결과")
check("추가한 표현을 사전이 잡음", any(h.rid.startswith("USER_") for h in hits2))
check("가중치=안전망 기준(단독 구제 가능)",
      any(h.weight >= merge.RESCUE_SCORE for h in hits2 if h.rid.startswith("USER_")))

# ── 6. 분석 기록이 없을 때(축소 모드) ────────────────────────────
(Path(TMP) / "analyses" / f"{fp}.json").unlink()
d2 = review.diff(review.read_reviewed(str(reviewed_path)))
check("축소 모드 동작", d2["mode"] == "fallback")
check("축소 모드: 확정 2건", d2["counts"]["gold"] == 2, str(d2["counts"]))
check("축소 모드: 고아 문구로 오탐 추정", d2["counts"]["fp"] == 1, str(d2["counts"]))

# ── 7. 여러 검토완료본을 한 번에 (preview_many) ──────────────────
# 두 번째 문서: 다른 내용의 검토본 (형광펜 A색 하나만 칠해 저장)
P3 = "신형 지그 고정 방식 적용 결과 공유"
reviewed2 = WORK / "reviewed2.pptx"
prs3 = Presentation()
layout3 = min(prs3.slide_layouts, key=lambda l: len(l.placeholders))
s3 = prs3.slides.add_slide(layout3)
for shp in list(s3.shapes):
    if shp.is_placeholder:
        shp._element.getparent().remove(shp._element)
tb3 = s3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1))
tb3.text_frame.text = P3
mark.highlight(tb3.text_frame.paragraphs[0]._p, (0, 8), config.GRADE_COLOR["A"], False)
prs3.save(str(reviewed2))

broken = WORK / "broken.pptx"
broken.write_text("이건 PPTX 가 아님", encoding="utf-8")

many = review.preview_many([
    ("reviewed.pptx", str(reviewed_path)),
    ("reviewed_복사본.pptx", str(reviewed_path)),   # 같은 문서를 두 번 (중복)
    ("second.pptx", str(reviewed2)),
    ("broken.pptx", str(broken)),
    ("note.txt", str(reviewed2)),                   # 확장자가 다른 파일
])
check("복수: 첫 파일 정상 처리", many[0].get("filename") == "reviewed.pptx"
      and "counts" in many[0])
check("복수: 같은 문서는 중복 표시(제외)", many[1].get("duplicate_of") == "reviewed.pptx")
check("복수: 두 번째 문서도 함께 처리", many[2].get("mode") == "fallback"
      and many[2]["counts"]["gold"] == 1, str(many[2].get("counts")))
check("복수: 깨진 파일은 그 파일만 오류", bool(many[3].get("error")))
check("복수: PPTX 아닌 파일 거부", bool(many[4].get("error")))
check("복수: 깨진 파일이 있어도 나머지는 살아 있음",
      sum(1 for m in many if "counts" in m) == 2)

# ── 8. 변경 일지 ─ 반영 전후로 무엇이 바뀌었는지 기록·리포트 ──────
log0 = review.change_log()
check("변경 일지: [사전에 추가] 이벤트 기록",
      any(e["kind"] == "rule_add" and e["keyword"] == "크로노백형합금" for e in log0))

res = review.commit_with_log([
    {"filename": "second.pptx", "mode": many[2]["mode"], "items": many[2]["items"]},
])
ch = res["change"]
check("변경 일지: 반영 이벤트 추가", len(review.change_log()) == len(log0) + 1)
check("변경 일지: dataset 전후 수치", ch["dataset"]["after"] == ch["dataset"]["before"] + 1,
      str(ch["dataset"]))
check("변경 일지: 유형별 집계", ch["added"]["gold"] == 1 and ch["added"]["total"] == 1,
      str(ch["added"]))
check("변경 일지: 새로 주입된 예시 추적",
      any(e["quote"] == "신형 지그 고정" for e in ch["examples"]["entered"]),
      str(ch["examples"]["entered"]))
check("변경 일지: 저장 결과 요약", res["saved_files"] == 1 and res["saved_items"] == 1)
try:
    review.commit_with_log([{"filename": "빈파일.pptx", "mode": "archive", "items": []}])
    check("변경 일지: 빈 반영 거부", False)
except ValueError:
    check("변경 일지: 빈 반영 거부", True)

# ── 9. 예시 선별 ─ 보관은 무제한, 주입은 잘 고른 8건 ─────────────
check("교정 항목에 카테고리가 붙는다 (사전에서 유추)",
      next(i for i in d["items"] if i["type"] == "miss").get("category") is None
      or isinstance(next(i for i in d["items"] if i["type"] == "miss").get("category"), str))
fp_item = next(i for i in d["items"] if i["type"] == "fp")
check("오탐 항목은 판정 기록의 카테고리를 그대로 쓴다",
      fp_item.get("category") == "효과만기재", str(fp_item.get("category")))

CATS = config.CATEGORIES
before_pool = len(review._examples_pool())

# 오탐이 몰린 보고서 한 건 — 이것만으로 반례 4칸을 독점하면 안 된다
review.commit_with_log([{"filename": "경영실적.pptx", "mode": "archive", "items": [
    {"type": "fp", "quote": f"매출 지표 {i}", "grade": "B", "risk": False,
     "category": "효과만기재" if i % 2 else "비교우위", "snippet": "…"} for i in range(6)]}])
# 이어서 기술 보고서 — 카테고리가 다양하다
review.commit_with_log([{"filename": "기술개발.pptx", "mode": "archive", "items": [
    {"type": "fp", "quote": f"공정 오탐 {i}", "grade": "B", "risk": False,
     "category": CATS[i], "snippet": "…"} for i in range(4)] + [
    {"type": "miss", "quote": f"기술 수단 {i}", "grade": "A", "risk": False,
     "category": CATS[i], "snippet": "…"} for i in range(4)]}])

pool = review._examples_pool()
check("보관 상한 없음 — 넣은 만큼 그대로 쌓인다",
      len(pool) == before_pool + 14, f"{before_pool} → {len(pool)}")

excl = review._pick(pool, "exclude")
incl = review._pick(pool, "include")
n_inc_cand = len([e for e in pool if e.get("kind") == "include"])
check("주입 칸은 종류별 상수를 따른다 (반례 4 · 정답 8)",
      len(excl) == review.EXAMPLE_INJECT_EXCLUDE
      and len(incl) == min(review.EXAMPLE_INJECT_INCLUDE, n_inc_cand),
      f"반례 {len(excl)} · 정답 {len(incl)} (정답 후보 {n_inc_cand}건)")
check("반례와 정답에 칸을 다르게 줄 수 있다",
      review.EXAMPLE_INJECT_EXCLUDE != review.EXAMPLE_INJECT_INCLUDE,
      f"{review.EXAMPLE_INJECT_EXCLUDE} + {review.EXAMPLE_INJECT_INCLUDE}")
check("한 번의 반영이 반례 4칸을 독점하지 못한다",
      len({e.get("batch") for e in excl}) >= 2
      and sum(1 for e in excl if e.get("batch") == max(p.get("batch") or 0
                                                       for p in pool))
      <= review.EXAMPLE_PER_BATCH,
      str([(e["quote"], e.get("batch")) for e in excl]))
check("카테고리가 겹치지 않게 고른다",
      len({e.get("category") for e in excl}) >= 3,
      str([e.get("category") for e in excl]))

# 축소 모드의 확정 라벨(gold)은 '모델이 실제로 놓친 것'(miss)보다 뒤로 밀린다
review.commit_with_log([{"filename": "다른PC분석본.pptx", "mode": "fallback", "items": [
    {"type": "gold", "quote": f"이미 잡던 표현 {i}", "grade": "A", "risk": False,
     "category": CATS[i], "snippet": "…"} for i in range(6)]}])
incl2 = review._pick(review._examples_pool(), "include")
kinds2 = [e.get("type") for e in incl2]
first_gold = kinds2.index("gold") if "gold" in kinds2 else len(kinds2)
check("gold 는 확증된 교정을 다 채운 뒤에만 들어간다",
      all(k == "gold" for k in kinds2[first_gold:])
      and not any(k == "gold" for k in kinds2[:first_gold]),
      str([(e["quote"], e.get("type")) for e in incl2]))
n_conf = len([e for e in review._examples_pool()
              if e.get("kind") == "include" and e.get("type") != "gold"])
check("확증된 교정이 칸보다 많으면 gold 는 하나도 안 들어간다",
      first_gold == min(len(incl2), n_conf),
      f"확증 후보 {n_conf}건 · 앞쪽 확증 {first_gold}칸")

# 그래도 gold 밖에 없으면 정상 주입되어야 한다 (칸이 비면 안 됨)
gold_only = [{"kind": "include", "type": "gold", "quote": f"확정 {i}", "grade": "A",
              "risk": False, "category": CATS[i], "date": "2026-01-01", "batch": 99,
              "snippet": "…"} for i in range(6)]
check("정답 후보가 gold 뿐이어도 칸이 비지 않는다",
      len(review._pick(gold_only, "include")) == min(6, review.EXAMPLE_INJECT_INCLUDE))

st2 = review.stats()
pool2 = review._examples_pool()
want = (len(review._pick(pool2, "exclude")) + len(review._pick(pool2, "include")))
check("화면용 집계에 보관/주입이 따로 나온다",
      st2["examples"] > st2["injected"] and st2["injected"] == want,
      f"보관 {st2['examples']} · 주입 {st2['injected']}")


# ── 10. 쌓인 것을 잃지 않는가 ─ 사용하다 보면 실제로 겪는 상황들 ──
# 아래 네 가지는 모두 "조용히 잘못되는" 부류라, 한 번 깨지면 사용자가 알아채기
# 어렵습니다. 그래서 상황을 그대로 만들어 놓고 확인합니다.
def _fresh(sub: str) -> Path:
    """검사마다 빈 review_data 폴더에서 시작한다 (앞 검사의 이력과 섞이지 않게)."""
    d = Path(TMP) / sub
    d.mkdir(parents=True, exist_ok=True)
    os.environ["PM_REVIEW_DIR"] = str(d)
    config.REVIEW_DIR = d
    return d


def _commit(name: str, items: list[dict]):
    return review.commit_with_log([{"filename": name, "mode": "archive",
                                    "items": items}])


_fresh("t10a")
# (가) [반영 저장] 한 번 = batch 한 개. 검토완료본을 5개 같이 올려도 마찬가지다.
for i in range(6):                                   # 지난 반영 이력을 쌓아 둔다
    _commit(f"과거{i}.pptx", [{"type": "miss", "quote": f"과거{i}-{j}", "grade": "B",
                              "category": f"과거{i}{j}"} for j in range(2)])
review.commit_with_log([{"filename": f"신규{k}.pptx", "mode": "archive", "items": [
    {"type": "miss", "quote": f"신규{k}-{j}", "grade": "B",
     "category": f"신규{k}{j}"} for j in range(2)]} for k in range(5)])
pool10 = review._examples_pool()
new_batches = {p["batch"] for p in pool10 if p["quote"].startswith("신규")}
check("파일을 여러 개 올려도 [반영 저장] 한 번은 batch 하나",
      len(new_batches) == 1, f"batch {sorted(new_batches)}")
inc10 = [e["quote"] for e in review._pick(pool10, "include")]
check("파일을 여러 개 올린 반영도 주입 칸을 독점하지 못한다",
      sum(1 for q in inc10 if q.startswith("신규")) <= review.EXAMPLE_PER_BATCH,
      str(inc10))

_fresh("t10b")
# (나) 사전에 없던 새 표현은 category 가 비어 있다. 이런 예시가 아무리 쌓여도
#     한 칸으로 뭉개지면, 도구가 가장 배워야 할 재료가 프롬프트에 안 들어간다.
for i in range(10):
    _commit(f"기존{i}.pptx", [{"type": "miss", "quote": f"기존{i}", "grade": "B",
                              "category": f"기존분류{i}"}])
_commit("새표현.pptx", [{"type": "miss", "quote": f"사전밖표현{j}", "grade": "B",
                       "category": None} for j in range(12)])
inc_b = [e["quote"] for e in review._pick(review._examples_pool(), "include")]
check("카테고리 없는 예시가 한 칸으로 뭉개지지 않는다",
      sum(1 for q in inc_b if q.startswith("사전밖")) == review.EXAMPLE_PER_BATCH,
      str(inc_b))

_fresh("t10c")
# (다) 같은 표현을 나중에 반대로 고쳤다면, 프롬프트에 "빼라"와 "잡아라"가 같이
#     들어가면 안 된다. 검토가 늘 그렇듯 나중 기록이 이긴다.
_commit("먼저.pptx", [{"type": "miss", "quote": "자체 개발한 공법", "grade": "A",
                      "category": "공정·방법"}])
_commit("나중.pptx", [{"type": "fp", "quote": "자체 개발한 공법",
                      "category": "공정·방법"}])
kinds_c = [p["kind"] for p in review._examples_pool() if p["quote"] == "자체 개발한 공법"]
check("같은 표현이 반례·정답 양쪽에 동시에 남지 않는다",
      kinds_c == ["exclude"], str(kinds_c))

d10 = _fresh("t10d")
# (라) examples.json 이 깨졌을 때(전원이 갑자기 꺼지는 등) 다음 [반영 저장] 이
#     빈 풀을 덮어써 버리면 그동안 쌓은 교정이 통째로 사라진다.
for i in range(5):
    _commit(f"보존{i}.pptx", [{"type": "miss", "quote": f"보존{i}", "grade": "B",
                              "category": f"보존분류{i}"}])
review._examples_path().write_text('[{"kind": "include", "quo', encoding="utf-8")
_commit("그다음.pptx", [{"type": "miss", "quote": "그다음", "grade": "B",
                       "category": "공정·방법"}])
kept = [p.name for p in d10.iterdir() if p.name.startswith("examples.broken-")]
check("깨진 examples.json 은 지우지 않고 옆에 보존한다", len(kept) == 1, str(kept))
check("깨진 파일을 치운 사실이 변경 일지에 남는다",
      any(e.get("kind") == "examples_recovered" for e in review.change_log()))
review._rules_path().write_text("{이것도 깨짐", encoding="utf-8")
review.add_rule("복구확인표현")
kept_r = [p.name for p in d10.iterdir() if p.name.startswith("extra_rules.broken-")]
check("깨진 extra_rules.json 도 마찬가지로 보존한다", len(kept_r) == 1, str(kept_r))

print(f"\n{'모두 통과' if fails == 0 else f'{fails}건 실패'}")
sys.exit(1 if fails else 0)
