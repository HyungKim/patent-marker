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
mark.apply(deck, resolved, marks, add_summary=True, tag_marks=True)
marked = WORK / "marked.pptx"
deck.prs.save(str(marked))
check("분석 기록 저장", review.load_archive(fp) is not None)

# ── 3. "사람의 검토" 흉내: P2 형광펜 제거(오탐), P1 에 새 형광펜(누락) ──
from lxml import etree                              # noqa: E402
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

print(f"\n{'모두 통과' if fails == 0 else f'{fails}건 실패'}")
sys.exit(1 if fails else 0)
