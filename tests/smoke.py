"""
tests/smoke.py ─ 가장 가벼운 점검 (LLM 없이, 몇 초)
=====================================================================
Ollama 나 모델이 없어도 돌아갑니다. "PPTX 읽기 → 규칙 사전 → 마킹 → 저장" 경로가
끊기지 않았는지만 확인합니다. 모델 자리는 규칙 사전 결과로 대신 채웁니다.

    실행:  python tests/smoke.py
           python tests/smoke.py samples/회사보고자료_예시.pptx out_smoke.pptx
"""
from __future__ import annotations

import sys
from pathlib import Path

# 프로젝트 루트를 import 경로에 추가 (어느 폴더에서 실행해도 app 패키지를 찾도록)
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import config, extract, lexicon, mark, merge  # noqa: E402
from app.analyze import Finding, prescreen  # noqa: E402

src = sys.argv[1] if len(sys.argv) > 1 else "samples/회사보고자료_예시.pptx"
dst = sys.argv[2] if len(sys.argv) > 2 else "out_smoke.pptx"

deck = extract.extract(src)
print(f"슬라이드 {deck.slide_count}장 · 문단 {len(deck.segments)}개")

opts = config.RunOptions()
hits_by_seg, targets = prescreen(deck, opts)
print(f"1차 스캔 대상 문단 {len(targets)}개")

# 모델 자리를 규칙 사전 결과로 대신 채운다
fake: list[Finding] = []
for seg in deck.segments:
    hits = hits_by_seg.get(seg.seg_id) or []
    if not hits:
        continue
    sc = lexicon.score(seg.text, hits)
    if sc < 8 and not lexicon.has_disclosure_risk(hits):
        continue
    top = max(hits, key=lambda h: h.weight)
    fake.append(
        Finding(
            seg_id=seg.seg_id, slide_no=seg.slide_no,
            quote=seg.text[top.span[0]:top.span[1]],
            grade="A" if sc >= 16 else "B",
            category=top.category,
            implicit=top.category in ("효과만기재", "독자성주장", "최적화·조건확립"),
            disclosure_risk=lexicon.has_disclosure_risk(hits),
            reason=top.hint, source="lexicon",
        )
    )

resolved, marks = merge.resolve(deck, fake, hits_by_seg)
print(f"후보 {len(resolved)}건 · 하이라이트 구간 {len(marks)}개")

stats = mark.apply(deck, resolved, marks, add_summary=True, tag_marks=True)
deck.prs.save(dst)
print(f"저장: {dst}  {stats}")

for f in resolved[:12]:
    flag = "⚠" if f.disclosure_risk else " "
    print(f"  p{f.slide_no} [{f.grade}]{flag} {f.category:<12} {f.quote[:44]!r}")
