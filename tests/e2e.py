"""
tests/e2e.py ─ 진짜 온디바이스 모델로 끝까지 돌려 보기 (수 분)
=====================================================================
Ollama 가 떠 있고 모델(qwen3:14b)이 내려받아져 있어야 합니다.
슬라이드마다 걸린 시간과 최종 후보 목록을 출력합니다.

    실행:  python tests/e2e.py
           python tests/e2e.py samples/회사보고자료_예시.pptx out_e2e.pptx
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import analyze, config, extract, mark, merge  # noqa: E402

src = sys.argv[1] if len(sys.argv) > 1 else "samples/회사보고자료_예시.pptx"
dst = sys.argv[2] if len(sys.argv) > 2 else "out_e2e.pptx"

print("health:", analyze.health())
t0 = time.time()

deck = extract.extract(src)
opts = config.RunOptions()
hits_by_seg, targets = analyze.prescreen(deck, opts)
print(f"슬라이드 {deck.slide_count}장 · 문단 {len(deck.segments)}개 · 분석 대상 {len(targets)}개\n")

by_slide: dict[int, list] = {}
for seg in targets:
    by_slide.setdefault(seg.slide_no, []).append(seg)

title = next((s.text.strip() for s in deck.segments if s.kind == "title"), "")

all_f = []
for n in range(1, deck.slide_count + 1):
    segs = by_slide.get(n, [])
    if not segs:
        continue
    hints = {s.seg_id: sorted({h.category for h in hits_by_seg.get(s.seg_id, [])})
             for s in segs}
    hints = {k: v for k, v in hints.items() if v}
    t = time.time()
    got = analyze.analyze_slide(title, n, deck.slide_count, segs, hints, opts)
    print(f"  슬라이드 {n}: {len(got)}건  ({time.time() - t:.1f}s)")
    all_f += got

resolved, marks = merge.resolve(deck, all_f, hits_by_seg)
located = sum(1 for f in resolved if f.span is not None)
llm = sum(1 for f in resolved if f.source == "llm")
print(f"\n후보 {len(resolved)}건 (모델 {llm} · 규칙 안전망 {len(resolved) - llm}) · "
      f"인용구 위치 확정 {located}/{len(resolved)} · 하이라이트 {len(marks)}구간")

stats = mark.apply(deck, resolved, marks, add_summary=True, tag_marks=True)
deck.prs.save(dst)
print(f"저장: {dst}  {stats}  총 {time.time() - t0:.1f}s\n")

for f in resolved:
    flag = "⚠" if f.disclosure_risk else " "
    imp = "묵시" if f.implicit else "명시"
    src_ = "M" if f.source == "llm" else "R"
    print(f"p{f.slide_no} [{f.grade}]{flag}{src_} {imp} {f.category:<12} {f.quote[:34]!r}")
    if f.reason:
        print(f"        {f.reason[:96]}")
