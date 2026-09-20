"""
pipeline.py ─ 파일 하나를 끝까지 처리하는 공통 흐름
=====================================================================

[이 파일이 하는 일]
  PPTX 한 개를  읽기 → 규칙 사전 → 모델 판정 → 병합 → 마킹 → 저장  까지 돌립니다.
  웹 서버(main.py)와 명령행 도구(cli.py)가 같은 함수를 부르므로, 어느 쪽으로
  실행하든 결과가 똑같습니다.

[초보자를 위한 설명]
  - progress : "지금 어느 단계인지" 를 알려 주는 콜백 함수. 웹 화면은 이 값으로
               진행 막대를 그리고, 명령행은 한 줄씩 출력합니다. 없어도 됩니다.
  - cancel   : '중단' 신호(threading.Event). 슬라이드를 넘어갈 때마다 확인해서
               켜져 있으면 Cancelled 예외로 멈춥니다.
"""
from __future__ import annotations

import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from . import analyze, config, extract, mark, merge, review


class Cancelled(Exception):
    """사용자가 중단을 눌러 멈춘 경우."""


@dataclass
class Progress:
    """progress 콜백에 넘겨 주는 현재 상태 한 묶음."""

    stage: str                  # 화면에 보여 줄 단계 설명
    slide_done: int = 0
    slide_total: int = 0
    findings: list | None = None    # 지금까지 찾은 후보 (to_public() 결과)


ProgressFn = Callable[[Progress], None]


def run(src: Path, dst: Path, opts: config.RunOptions,
        progress: ProgressFn | None = None,
        cancel: threading.Event | None = None) -> tuple[list, dict]:
    """src 를 분석해 dst 에 마킹본을 저장한다. (후보 목록, 집계) 를 돌려준다.

    오류는 그대로 올려 보낸다 — 부르는 쪽(main.py / cli.py)이 화면에 맞게 보여 준다.
    """
    def report(stage: str, done: int = 0, total: int = 0, findings: list | None = None) -> None:
        if progress is not None:
            progress(Progress(stage, done, total, findings))

    # ── 1단계: PPTX 읽기 ──────────────────────────────────────
    report("문서 읽는 중")
    deck = extract.extract(str(src))
    total = deck.slide_count

    # ── 2단계: 규칙 사전으로 1차 스캔 ─────────────────────────
    report("규칙 사전 1차 스캔", 0, total)
    hits_by_seg, targets = analyze.prescreen(deck, opts)

    by_slide: dict[int, list] = {}
    for seg in targets:
        by_slide.setdefault(seg.slide_no, []).append(seg)

    title = ""
    for seg in deck.segments:
        if seg.kind == "title":
            title = seg.text.strip()
            break

    # ── 3단계: 슬라이드마다 온디바이스 모델에게 판정 요청 ──────
    all_findings: list[analyze.Finding] = []
    for slide_no in range(1, total + 1):
        if cancel is not None and cancel.is_set():
            raise Cancelled()
        segs = by_slide.get(slide_no, [])
        report(f"슬라이드 {slide_no} 분석 중 (온디바이스 모델)", slide_no - 1, total,
               [f.to_public() for f in all_findings])
        if segs:
            # 규칙 사전이 찾은 카테고리를 힌트로 함께 보낸다
            hints = {
                s.seg_id: [h.category for h in hits_by_seg.get(s.seg_id, [])]
                for s in segs
            }
            hints = {k: sorted(set(v)) for k, v in hints.items() if v}
            all_findings += analyze.analyze_slide(title, slide_no, total, segs, hints, opts)
        report(f"슬라이드 {slide_no} 분석 완료", slide_no, total,
               [f.to_public() for f in all_findings])

    # ── 4단계: 규칙 결과와 모델 결과 병합, 등급 확정 ──────────
    report("결과 병합 및 등급 산정", total, total)
    resolved, marks = merge.resolve(deck, all_findings, hits_by_seg)

    # 판정 기록을 자동 저장해 둔다 — [검토 반영] 탭이 검토완료본과 비교할 기준.
    # 실패해도 분석 자체는 계속한다.
    try:
        review.save_archive(deck, resolved)
    except Exception:  # noqa: BLE001
        traceback.print_exc()

    # ── 5단계: PPTX 에 마킹하고 저장 ──────────────────────────
    report("PPTX 마킹 중", total, total, [f.to_public() for f in resolved])
    stats = mark.apply(deck, resolved, marks, tag_marks=opts.tag_marks)
    dst.parent.mkdir(parents=True, exist_ok=True)
    deck.prs.save(str(dst))

    # ── 집계 ──────────────────────────────────────────────────
    counts = {"A": 0, "B": 0, "C": 0}
    risk = implicit = 0
    for f in resolved:
        counts[f.grade] = counts.get(f.grade, 0) + 1
        risk += 1 if f.disclosure_risk else 0
        implicit += 1 if f.implicit else 0
    summary = {
        **stats,
        "total": len(resolved),
        "grades": counts,
        "disclosure_risk": risk,
        "implicit": implicit,
        "segments": len(deck.segments),
    }
    report("완료", total, total, [f.to_public() for f in resolved])
    return resolved, summary
