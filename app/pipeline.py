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
               모델을 기다리는 동안에는 2초마다 "N분 NN초 경과 · 문단 읽는 중 / 답변 작성 중 N자" 로 갱신됩니다.
  - cancel   : '중단' 신호(threading.Event). 슬라이드 사이에서 확인할 뿐 아니라,
               모델 호출 도중에도 analyze._chat 이 소켓을 끊어 즉시 멈춥니다 (Cancelled 예외).
  - 시간 제한 : 없습니다 (config.REQUEST_TIMEOUT=0). 느린 PC 에서 슬라이드 하나가 10분을 넘겨도
               끊지 않습니다. 대신 Ollama 프로세스가 사라지면 analyze 가 감지해 OllamaError 로 멈춥니다.
  - 속도 기록 : 슬라이드마다 "입력/출력 토큰 · 초당 토큰" 을 progress 로 알리고, 파일 하나가 끝나면
               review_data/run_log.tsv 에 한 줄을 남깁니다 (엑셀로 열림). 개선 전후를 숫자로 비교하는 근거.
"""
from __future__ import annotations

import datetime as _dt
import threading
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from . import analyze, config, extract, mark, memory, merge, review

RUN_LOG_COLUMNS = ["일시", "버전", "모델", "파일", "슬라이드", "모델호출", "입력토큰", "출력토큰",
                   "읽기초", "쓰기초", "쓰기토큰/초", "총소요초", "후보", "A", "B", "C", "인용일치", "학습", "결과파일"]


def learn_text(lt: dict) -> str:
    """검토 학습 적용 결과 한 줄. 예: '기억 3문단(+5/−4) · 제외 1 · 사례 12(뜻 기준)'  /  꺼져 있으면 '끔'"""
    if not lt or not lt.get("on"):
        return "끔"
    mode = {"embed": "뜻 기준", "ngram": "글자 겹침", "off": "없음"}.get(lt.get("mode", "off"), "없음")
    return (f"기억 {lt['memory_paras']}문단(+{lt['memory_added']}/−{lt['memory_dropped']}) · "
            f"제외 {lt['excluded']} · 확정 표현 구제 {lt.get('rescued', 0)} · 사례 {lt['examples']}({mode})")


def speed_text(st: dict) -> str:
    """집계 상자를 사람이 읽는 한 줄로. 예: '입력 2,187 / 출력 1,870 토큰 · 쓰기 9.3 토큰/초'"""
    tps = st["output_tokens"] / st["output_sec"] if st.get("output_sec") else 0.0
    return (f"입력 {st['prompt_tokens']:,} / 출력 {st['output_tokens']:,} 토큰"
            + (f" · 쓰기 {tps:.1f} 토큰/초" if tps else ""))


def _append_run_log(row: list) -> None:
    """review_data/run_log.tsv 에 한 줄 추가. 실패해도 분석 결과에는 영향을 주지 않는다."""
    try:
        p = config.REVIEW_DIR / "run_log.tsv"
        p.parent.mkdir(parents=True, exist_ok=True)
        header = "\t".join(RUN_LOG_COLUMNS)
        new = not p.exists()
        # 열이 늘어난 새 버전이면(예전 머리글과 다르면) 빈 줄 뒤에 새 머리글을 한 번 더 적는다 — 옛 줄과 섞이지 않게
        stale = False
        if not new:
            with p.open("r", encoding="utf-8-sig") as f:
                lines = f.read().splitlines()
            last_header = next((ln for ln in reversed(lines) if ln.startswith("일시\t")), "")
            stale = last_header != header
        with p.open("a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
            if new or stale:
                f.write(("" if new else "\n") + header + "\n")      # BOM 을 붙여 엑셀이 한글을 바로 읽게
            f.write("\t".join(str(v) for v in row) + "\n")
    except Exception:  # noqa: BLE001
        traceback.print_exc()


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

    t_run = time.monotonic()
    # 검토 학습 스위치를 이 스레드에 건다 — 규칙 사전(자동 규칙)·기억·사례 검색이 모두 이 값을 본다
    memory.set_context(opts.learn)
    try:
        return _run(src, dst, opts, report, cancel, t_run)
    finally:
        memory.clear_context()


def _run(src: Path, dst: Path, opts: config.RunOptions, report, cancel, t_run: float) -> tuple[list, dict]:
    learn = {"on": bool(opts.learn), "memory_paras": 0, "memory_added": 0, "memory_dropped": 0,
             "excluded": 0, "rescued": 0, "examples": 0, "mode": "off"}
    use_embed = bool(opts.learn and opts.learn_embed and memory.embed_available())

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
    run_stats = analyze.new_stats()                 # 파일 전체의 토큰·시간 집계 (속도 기록용)
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

            # 시간 제한이 없는 대신 "살아 있음" 을 보여 준다: 경과 시간 + 모델이 읽는 중인지 쓰는 중인지
            t_slide = time.monotonic()
            snapshot = [f.to_public() for f in all_findings]

            def alive(nchars: int, _n=slide_no, _t=t_slide, _snap=snapshot) -> None:
                sec = int(time.monotonic() - _t)
                phase = f"답변 작성 중 {nchars}자" if nchars else "문단 읽는 중"
                report(f"슬라이드 {_n} 분석 중 (온디바이스 모델) · {sec // 60}분 {sec % 60:02d}초 경과 · {phase}",
                       _n - 1, total, _snap)

            # ③ 이 슬라이드와 비슷한 확정 사례를 골라 지시서에 붙인다 (학습이 꺼져 있거나 데이터가 없으면 빈 문자열)
            block, info = memory.examples_for([s.text for s in segs], use_embed=use_embed)
            system = (analyze.SYSTEM + "\n\n" + block) if block else None
            learn["examples"] += info["pos"] + info["neg"]
            if info["mode"] != "off" and block:
                learn["mode"] = info["mode"]

            slide_stats = analyze.new_stats()
            try:
                all_findings += analyze.analyze_slide(title, slide_no, total, segs, hints, opts,
                                                      cancel=cancel, progress=alive, stats=slide_stats,
                                                      system=system)
            except analyze.Aborted:
                raise Cancelled()
            analyze.add_stats(run_stats, slide_stats)
            sec = int(time.monotonic() - t_slide)
            # 슬라이드 하나의 성적표 — 회사 PC 의 실제 속도가 여기서 숫자로 드러난다
            report(f"슬라이드 {slide_no} 분석 완료 · {sec // 60}분 {sec % 60:02d}초 · {speed_text(slide_stats)}",
                   slide_no, total, [f.to_public() for f in all_findings])
            continue
        report(f"슬라이드 {slide_no} 분석 완료", slide_no, total,
               [f.to_public() for f in all_findings])

    # ── 4단계: 규칙 결과와 모델 결과 병합, 등급 확정 ──────────
    report("결과 병합 및 등급 산정", total, total)
    resolved, marks = merge.resolve(deck, all_findings, hits_by_seg)

    # ── 4-1단계: 검토 학습 — ① 기억된 문단은 사람 판정으로, ② 지운 표현은 제외 ──
    if opts.learn:
        def _mem_finding(seg, quote, span, grade, risk, category, reason):
            return analyze.Finding(seg_id=seg.seg_id, slide_no=seg.slide_no, quote=quote, grade=grade,
                                   category=category, implicit=category in config.IMPLICIT_CATEGORIES,
                                   disclosure_risk=risk, reason=reason, source="memory", span=span)
        def _rule_finding(seg, quote, span, grade, risk, category, reason):
            return analyze.Finding(seg_id=seg.seg_id, slide_no=seg.slide_no, quote=quote, grade=grade,
                                   category=category, implicit=category in config.IMPLICIT_CATEGORIES,
                                   disclosure_risk=risk, reason=reason, source="lexicon", span=span)
        resolved, mst = memory.apply_memory(deck, resolved, _mem_finding)
        resolved, n_ex = memory.apply_excludes(resolved)
        resolved, n_rs = memory.rescue_confirmed(deck, resolved, hits_by_seg, _rule_finding)
        learn.update(memory_paras=mst["paras"], memory_added=mst["added"],
                     memory_dropped=mst["dropped"], excluded=n_ex, rescued=n_rs)
        if mst["paras"] or n_ex or n_rs:
            resolved = merge.sort_findings(resolved)
            marks = merge.build_marks(deck, resolved)

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
    elapsed = time.monotonic() - t_run
    tps = run_stats["output_tokens"] / run_stats["output_sec"] if run_stats["output_sec"] else 0.0
    # 품질 지표 하나: 모델 인용구가 원문과 정확히 일치해 자리를 잡은 비율 (못 잡으면 문단 전체가 칠해진다)
    llm_n = sum(1 for f in resolved if f.source == "llm")
    located = sum(1 for f in resolved if f.source == "llm" and f.span is not None)
    summary = {
        **stats,
        "total": len(resolved),
        "grades": counts,
        "disclosure_risk": risk,
        "implicit": implicit,
        "segments": len(deck.segments),
        # 속도 성적표 (화면 결과 칸·명령행·실행 기록에 쓰임)
        "model": opts.model,
        "version": config.VERSION,
        "seconds": round(elapsed, 1),
        "calls": run_stats["calls"],
        "prompt_tokens": run_stats["prompt_tokens"],
        "output_tokens": run_stats["output_tokens"],
        "prompt_sec": round(run_stats["prompt_sec"], 1),
        "output_sec": round(run_stats["output_sec"], 1),
        "tokens_per_sec": round(tps, 1),
        "quote_located": f"{located}/{llm_n}",
        "speed_text": f"{speed_text(run_stats)} · 총 {int(elapsed) // 60}분 {int(elapsed) % 60:02d}초",
        "learn": learn,
        "learn_text": learn_text(learn),
    }
    _append_run_log([
        _dt.datetime.now().strftime("%Y-%m-%d %H:%M"), config.VERSION, opts.model, src.name, total,
        run_stats["calls"], run_stats["prompt_tokens"], run_stats["output_tokens"],
        round(run_stats["prompt_sec"], 1), round(run_stats["output_sec"], 1), round(tps, 1),
        round(elapsed, 1), len(resolved), counts["A"], counts["B"], counts["C"], f"{located}/{llm_n}",
        learn_text(learn), dst.name,
    ])
    report("완료", total, total, [f.to_public() for f in resolved])
    return resolved, summary
