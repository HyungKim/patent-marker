"""
main.py ─ 웹 서버 (프로그램의 "현관문")
=====================================================================

[이 파일이 하는 일]
  브라우저와 대화하는 창구입니다. 사용자가 화면에서 파일을 올리면
  이 파일이 받아서 extract → analyze → merge → mark 순서로 다른 파일들을
  불러 일을 시키고, 진행 상황과 결과를 브라우저에 돌려줍니다.

[초보자를 위한 설명]
  - FastAPI : "웹 주소(URL)마다 어떤 파이썬 함수를 실행할지" 를 정해 주는 라이브러리.
              아래 @app.get("/...") / @app.post("/...") 가 그 연결 고리입니다.
  - uvicorn : FastAPI 앱을 실제로 띄워 주는 서버 프로그램. serve() 안에서 실행합니다.
  - 127.0.0.1 : "내 컴퓨터 자신". 이 주소로만 열기 때문에 외부에서 접속할 수 없습니다.

  브라우저 화면(static/index.html)과 이 파일 사이의 약속(API):
      GET  /                       화면(HTML) 내려주기
      GET  /api/health             Ollama 와 모델이 준비됐는지
      POST /api/jobs               파일 업로드 → 분석 시작 (job_id 반환)
      GET  /api/jobs/{id}          진행 상황 + 지금까지 찾은 후보
      POST /api/jobs/{id}/cancel   중단
      GET  /api/jobs/{id}/download 마킹된 PPTX 내려받기
      DELETE /api/jobs/{id}        작업과 임시 파일 삭제

[분석이 오래 걸리는데 화면이 멈추지 않는 이유]
  분석은 별도의 '스레드(thread)' 에서 돌립니다. 브라우저는 0.9초마다
  /api/jobs/{id} 를 물어보며 진행률을 갱신합니다. (index.html 의 poll() 참고)
"""
from __future__ import annotations

import shutil
import tempfile
import threading
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from fastapi import FastAPI, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import analyze, config, extract, mark, merge

# 화면 파일(index.html)이 있는 폴더
STATIC = Path(__file__).parent / "static"

# docs_url=None : FastAPI 가 기본으로 만드는 API 문서 화면을 끕니다(불필요).
app = FastAPI(title="특허 마킹 도구", docs_url=None, redoc_url=None)


@dataclass
class Job:
    """분석 작업 한 건의 상태. 메모리(JOBS 딕셔너리)에만 보관한다."""

    job_id: str
    filename: str
    workdir: Path                   # 임시 작업 폴더
    src: Path                       # 업로드된 원본
    out: Path                       # 마킹된 결과 파일
    status: str = "queued"          # queued | running | done | error | cancelled
    stage: str = ""                 # 화면에 보여 줄 현재 단계 설명
    slide_done: int = 0
    slide_total: int = 0
    findings: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    error: str = ""
    cancel: threading.Event = field(default_factory=threading.Event)  # 중단 신호

    def snapshot(self) -> dict:
        """브라우저에 보낼 수 있는 형태(JSON 가능)로 현재 상태를 복사한다."""
        return {
            "job_id": self.job_id,
            "filename": self.filename,
            "status": self.status,
            "stage": self.stage,
            "slide_done": self.slide_done,
            "slide_total": self.slide_total,
            "findings": self.findings,
            "stats": self.stats,
            "error": self.error,
        }


JOBS: dict[str, Job] = {}       # job_id → Job
_LOCK = threading.Lock()


# ═════════════════════════════════════════════════════════════════
# 실제 분석 파이프라인 (별도 스레드에서 실행)
# ═════════════════════════════════════════════════════════════════
def _run(job: Job, opts: config.RunOptions) -> None:
    """파일 하나를 끝까지 처리한다. 진행 상황은 job 객체에 계속 적어 둔다."""
    try:
        # ── 1단계: PPTX 읽기 ──────────────────────────────────
        job.status, job.stage = "running", "문서 읽는 중"
        deck = extract.extract(str(job.src))
        job.slide_total = deck.slide_count

        # ── 2단계: 규칙 사전으로 1차 스캔 ─────────────────────
        job.stage = "규칙 사전 1차 스캔"
        hits_by_seg, targets = analyze.prescreen(deck, opts)

        by_slide: dict[int, list] = {}
        for seg in targets:
            by_slide.setdefault(seg.slide_no, []).append(seg)

        title = ""
        for seg in deck.segments:
            if seg.kind == "title":
                title = seg.text.strip()
                break

        # ── 3단계: 슬라이드마다 온디바이스 모델에게 판정 요청 ──
        all_findings: list[analyze.Finding] = []
        for slide_no in range(1, deck.slide_count + 1):
            if job.cancel.is_set():                       # 사용자가 '중단' 을 눌렀으면
                job.status, job.stage = "cancelled", "사용자 중단"
                return
            segs = by_slide.get(slide_no, [])
            job.stage = f"슬라이드 {slide_no} 분석 중 (온디바이스 모델)"
            if segs:
                # 규칙 사전이 찾은 카테고리를 힌트로 함께 보낸다
                hints = {
                    s.seg_id: [h.category for h in hits_by_seg.get(s.seg_id, [])]
                    for s in segs
                }
                hints = {k: sorted(set(v)) for k, v in hints.items() if v}
                all_findings += analyze.analyze_slide(
                    title, slide_no, deck.slide_count, segs, hints, opts
                )
            job.slide_done = slide_no
            job.findings = [f.to_public() for f in all_findings]   # 중간 결과도 화면에

        # ── 4단계: 규칙 결과와 모델 결과 병합, 등급 확정 ──────
        job.stage = "결과 병합 및 등급 산정"
        resolved, marks = merge.resolve(
            deck, all_findings, hits_by_seg, include_grade_c=opts.include_grade_c
        )
        job.findings = [f.to_public() for f in resolved]

        # ── 5단계: PPTX 에 마킹하고 저장 ──────────────────────
        job.stage = "PPTX 마킹 중"
        stats = mark.apply(deck, resolved, marks,
                           add_summary=opts.add_summary, tag_marks=opts.tag_marks)
        deck.prs.save(str(job.out))

        # ── 집계 ──────────────────────────────────────────────
        counts = {"A": 0, "B": 0, "C": 0}
        risk = 0
        implicit = 0
        for f in resolved:
            counts[f.grade] = counts.get(f.grade, 0) + 1
            risk += 1 if f.disclosure_risk else 0
            implicit += 1 if f.implicit else 0
        job.stats = {
            **stats,
            "total": len(resolved),
            "grades": counts,
            "disclosure_risk": risk,
            "implicit": implicit,
            "segments": len(deck.segments),
        }
        job.stage = "완료"
        job.status = "done"
    except analyze.OllamaError as e:
        job.status, job.error, job.stage = "error", str(e), "온디바이스 모델 연결 실패"
    except Exception as e:  # noqa: BLE001  (어떤 오류든 화면에 보여 주기 위해 전부 잡음)
        job.status = "error"
        job.error = f"{type(e).__name__}: {e}"
        job.stage = "오류"
        traceback.print_exc()


# ═════════════════════════════════════════════════════════════════
# 웹 주소(URL) ↔ 함수 연결
# ═════════════════════════════════════════════════════════════════
@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    """첫 화면. static/index.html 을 그대로 돌려준다."""
    return HTMLResponse((STATIC / "index.html").read_text(encoding="utf-8"))


@app.get("/api/health")
def health() -> JSONResponse:
    """Ollama 가 떠 있는지, 모델이 내려받아져 있는지 확인. 화면 상단 상태 표시에 쓴다."""
    h = analyze.health()
    h["model"] = config.MODEL
    h["model_ready"] = analyze.model_available(config.MODEL, h.get("models", []))
    h["host"] = config.OLLAMA_HOST
    return JSONResponse(h)


@app.post("/api/jobs")
async def create_job(
    file: UploadFile,
    model: str = Form(config.MODEL),
    think: bool = Form(False),
    include_grade_c: bool = Form(False),
    scan_all: bool = Form(True),
    add_summary: bool = Form(True),
    tag_marks: bool = Form(True),
) -> JSONResponse:
    """파일 업로드를 받아 임시 폴더에 저장하고, 분석 스레드를 시작한다."""
    name = file.filename or "deck.pptx"
    if not name.lower().endswith((".pptx", ".potx")):
        raise HTTPException(400, "PPTX 파일만 지원합니다. (.ppt 는 먼저 .pptx 로 변환하세요)")

    job_id = uuid.uuid4().hex[:12]
    workdir = Path(tempfile.mkdtemp(prefix=f"pm-{job_id}-"))   # OS 임시 폴더 아래
    src = workdir / "input.pptx"
    stem = Path(name).stem
    out = workdir / f"{stem}_특허마킹.pptx"

    with src.open("wb") as fh:
        shutil.copyfileobj(file.file, fh)

    job = Job(job_id=job_id, filename=name, workdir=workdir, src=src, out=out)
    with _LOCK:
        JOBS[job_id] = job

    # 화면의 체크박스 값 → RunOptions
    opts = config.RunOptions(
        model=model or config.MODEL,
        think=think,
        scan_all_paragraphs=scan_all,
        include_grade_c=include_grade_c,
        add_summary=add_summary,
        tag_marks=tag_marks,
    )
    # daemon=True : 서버를 끄면 분석 스레드도 같이 종료
    threading.Thread(target=_run, args=(job, opts), daemon=True).start()
    return JSONResponse({"job_id": job_id})


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> JSONResponse:
    """진행 상황 조회. 브라우저가 주기적으로 호출한다."""
    job = JOBS.get(job_id)
    if job is None:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")
    return JSONResponse(job.snapshot())


@app.post("/api/jobs/{job_id}/cancel")
def job_cancel(job_id: str) -> JSONResponse:
    """중단 요청. 분석 스레드가 다음 슬라이드로 넘어갈 때 신호를 확인하고 멈춘다."""
    job = JOBS.get(job_id)
    if job is None:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")
    job.cancel.set()
    return JSONResponse({"ok": True})


@app.get("/api/jobs/{job_id}/download")
def job_download(job_id: str):
    """마킹된 PPTX 파일 내려받기."""
    job = JOBS.get(job_id)
    if job is None or job.status != "done" or not job.out.exists():
        raise HTTPException(404, "다운로드할 결과가 아직 없습니다.")
    return FileResponse(
        str(job.out),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=job.out.name,
    )


@app.delete("/api/jobs/{job_id}")
def job_delete(job_id: str) -> JSONResponse:
    """작업 기록과 임시 파일을 지운다."""
    job = JOBS.pop(job_id, None)
    if job is not None:
        job.cancel.set()
        shutil.rmtree(job.workdir, ignore_errors=True)
    return JSONResponse({"ok": True})


# /static/... 주소로 static 폴더의 파일을 그대로 내어 준다
app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")


def serve() -> None:
    """서버를 띄운다. run.sh / run.bat 이 `python -m app.main` 으로 이 함수를 부른다."""
    import uvicorn

    print(f"\n  특허 마킹 도구  →  http://{config.HOST}:{config.PORT}")
    print(f"  온디바이스 모델 : {config.MODEL} @ {config.OLLAMA_HOST}")
    print("  외부 네트워크로 나가는 통신은 없습니다.")
    print("  종료하려면 이 창에서 Ctrl + C\n")
    uvicorn.run(app, host=config.HOST, port=config.PORT, log_level="warning")


if __name__ == "__main__":
    serve()
