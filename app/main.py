"""
main.py ─ 웹 서버 (프로그램의 "현관문")
=====================================================================

[이 파일이 하는 일]
  브라우저와 대화하는 창구입니다. 사용자가 화면에서 파일을 고르면
  이 파일이 받아서 pipeline.py(읽기 → 판정 → 병합 → 마킹)에 넘기고,
  진행 상황과 결과를 브라우저에 돌려줍니다.

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
      DELETE /api/jobs/{id}        작업 기록 삭제

  브라우저 업로드가 막힌 PC 를 위한 통로 (app/local.py 참고, 화면 토큰 필요):
      GET  /api/local/files            input · output 폴더의 PPTX 목록
      POST /api/local/jobs             {path} 경로의 파일로 분석 시작
      POST /api/local/pick             파이썬이 파일 선택창을 띄움 → {paths}
      POST /api/local/open             탐색기로 폴더 열기
      POST /api/local/review/preview   {paths} 경로의 검토완료본으로 교정 내역 계산

[분석이 오래 걸리는데 화면이 멈추지 않는 이유]
  분석은 별도의 '스레드(thread)' 에서 돌립니다. 브라우저는 0.9초마다
  /api/jobs/{id} 를 물어보며 진행률을 갱신합니다. (index.html 의 poll() 참고)

[결과가 저장되는 곳]
  어느 방법으로 넣었든 결과는 output 폴더에  이름_특허마킹.pptx  로 저장됩니다.
  다운로드 버튼은 같은 파일을 브라우저로 내려 주는 것뿐입니다.
"""
from __future__ import annotations

import asyncio
import secrets
import shutil
import tempfile
import threading
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from fastapi import Body, Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import analyze, config, evaluate, local, pipeline, review

# 화면 파일(index.html)이 있는 폴더
STATIC = Path(__file__).parent / "static"

# docs_url=None : FastAPI 가 기본으로 만드는 API 문서 화면을 끕니다(불필요).
app = FastAPI(title="특허 마킹 도구", docs_url=None, redoc_url=None)

# 화면 토큰 — 서버를 켤 때마다 새로 만들고 index.html 안에 심어 둔다.
# 경로로 파일을 읽는 기능(/api/local/…)은 이 값을 아는 화면만 부를 수 있다.
# (같은 브라우저의 다른 웹페이지는 이 값을 읽을 수 없으므로 몰래 호출하지 못한다)
TOKEN = secrets.token_urlsafe(24)


def require_token(x_pm_token: str | None = Header(default=None)) -> None:
    """/api/local/… 전용 확인. 헤더 X-PM-Token 이 화면에 심어 둔 값과 같아야 한다."""
    if not x_pm_token or not secrets.compare_digest(x_pm_token.encode("utf-8"), TOKEN.encode("utf-8")):
        raise HTTPException(403, "이 기능은 프로그램 화면에서만 쓸 수 있습니다. "
                                 "화면을 새로고침한 뒤 다시 시도하세요.")


@dataclass
class Job:
    """분석 작업 한 건의 상태. 메모리(JOBS 딕셔너리)에만 보관한다."""

    job_id: str
    filename: str
    src: Path                       # 원본 (업로드면 임시 폴더의 복사본, 경로 지정이면 그 파일)
    out: Path                       # 마킹된 결과 파일 (output 폴더)
    source: str = "upload"          # upload | local
    workdir: Path | None = None     # 업로드 임시 폴더 (경로 지정이면 없음)
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
            "source": self.source,
            "status": self.status,
            "stage": self.stage,
            "slide_done": self.slide_done,
            "slide_total": self.slide_total,
            "findings": self.findings,
            "stats": self.stats,
            "error": self.error,
            "out_path": str(self.out),
            "out_name": self.out.name,
            "out_dir": str(self.out.parent),
        }


JOBS: dict[str, Job] = {}       # job_id → Job
_LOCK = threading.Lock()


# ═════════════════════════════════════════════════════════════════
# 분석 스레드
# ═════════════════════════════════════════════════════════════════
def _run(job: Job, opts: config.RunOptions) -> None:
    """pipeline.run 을 돌리며 진행 상황을 job 객체에 계속 적어 둔다."""
    def progress(p: pipeline.Progress) -> None:
        job.stage = p.stage
        job.slide_done, job.slide_total = p.slide_done, p.slide_total
        if p.findings is not None:
            job.findings = p.findings

    try:
        job.status = "running"
        _, job.stats = pipeline.run(job.src, job.out, opts, progress=progress, cancel=job.cancel)
        job.stage, job.status = "완료", "done"
    except pipeline.Cancelled:
        job.status, job.stage = "cancelled", "사용자 중단"
    except analyze.OllamaError as e:
        job.status, job.error, job.stage = "error", f"{e}  (버전 {config.VERSION})", "온디바이스 모델 연결 실패"
    except Exception as e:  # noqa: BLE001  (어떤 오류든 화면에 보여 주기 위해 전부 잡음)
        job.status = "error"
        # 어느 단계에서, 어느 버전이 냈는지 같이 적는다 — 화면 문구만 전달받아도 원인을 좁힐 수 있게
        job.error = f"{type(e).__name__}: {e}  (단계: {job.stage} · 버전 {config.VERSION})"
        job.stage = "오류"
        traceback.print_exc()
    finally:
        if job.status != "done":
            _release_output(job)
        if job.workdir is not None:          # 업로드 복사본은 분석이 끝나면 필요 없다
            shutil.rmtree(job.workdir, ignore_errors=True)


def _reserve_output(name: str) -> Path:
    """output 폴더에 결과 파일 이름을 잡아 둔다 (같은 이름이 동시에 겹치지 않도록 빈 파일로 예약)."""
    out = local.output_path_for(name)
    out.touch()
    return out


def _release_output(job: Job) -> None:
    """실패·중단한 작업의 예약 파일(빈 파일)을 치운다."""
    try:
        if job.out.exists() and job.out.stat().st_size == 0:
            job.out.unlink()
    except OSError:
        pass


def _start(name: str, src: Path, opts: config.RunOptions,
           source: str, workdir: Path | None = None) -> Job:
    """Job 을 만들고 분석 스레드를 시작한다. 업로드·경로 지정 두 통로가 함께 쓴다."""
    job = Job(job_id=uuid.uuid4().hex[:12], filename=name, src=src,
              out=_reserve_output(name), source=source, workdir=workdir)
    with _LOCK:
        JOBS[job.job_id] = job
    # daemon=True : 서버를 끄면 분석 스레드도 같이 종료
    threading.Thread(target=_run, args=(job, opts), daemon=True).start()
    return job


def _options(model: str, think: bool, scan_all: bool, tag_marks: bool) -> config.RunOptions:
    """화면의 체크박스 값 → RunOptions"""
    return config.RunOptions(model=model or config.MODEL, think=think,
                             scan_all_paragraphs=scan_all, tag_marks=tag_marks)


# ═════════════════════════════════════════════════════════════════
# 웹 주소(URL) ↔ 함수 연결
# ═════════════════════════════════════════════════════════════════
@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    """첫 화면. static/index.html 에 화면 토큰과 버전을 심어 돌려준다."""
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    html = html.replace("__PM_TOKEN__", TOKEN).replace("__PM_VERSION__", config.VERSION)
    # no-store: 업데이트 뒤 브라우저가 예전 화면을 캐시에서 꺼내 쓰지 않게 (Ctrl+F5 를 잊어도 새 화면)
    return HTMLResponse(html, headers={"Cache-Control": "no-store"})


@app.get("/api/health")
def health() -> JSONResponse:
    """Ollama 가 떠 있는지, 모델이 내려받아져 있는지 확인. 화면 상단 상태 표시에 쓴다."""
    h = analyze.health()
    h["model"] = config.MODEL
    h["model_ready"] = analyze.model_available(config.MODEL, h.get("models", []))
    h["host"] = config.OLLAMA_HOST
    h["version"] = config.VERSION
    return JSONResponse(h)


@app.post("/api/jobs")
async def create_job(
    file: UploadFile,
    model: str = Form(config.MODEL),
    think: bool = Form(False),
    scan_all: bool = Form(True),
    tag_marks: bool = Form(False),
) -> JSONResponse:
    """파일 업로드를 받아 임시 폴더에 저장하고, 분석 스레드를 시작한다."""
    name = file.filename or "deck.pptx"
    if not name.lower().endswith(local.PPTX_SUFFIXES):
        raise HTTPException(400, "PPTX 파일만 지원합니다. (.ppt 는 먼저 .pptx 로 변환하세요)")

    workdir = Path(tempfile.mkdtemp(prefix="pm-up-"))   # OS 임시 폴더 아래
    src = workdir / "input.pptx"
    with src.open("wb") as fh:
        shutil.copyfileobj(file.file, fh)

    job = _start(name, src, _options(model, think, scan_all, tag_marks), "upload", workdir)
    return JSONResponse({"job_id": job.job_id})


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
    """마킹된 PPTX 파일 내려받기 (output 폴더에 저장된 것과 같은 파일)."""
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
    """작업 기록과 업로드 임시 파일을 지운다. (output 폴더의 결과는 남긴다)"""
    job = JOBS.pop(job_id, None)
    if job is not None:
        job.cancel.set()
        if job.workdir is not None:
            shutil.rmtree(job.workdir, ignore_errors=True)
    return JSONResponse({"ok": True})


# ═════════════════════════════════════════════════════════════════
# 브라우저 업로드 없이 파일 넣기 (input 폴더 · 경로 입력 · 파일 선택창)
# ═════════════════════════════════════════════════════════════════
@app.get("/api/local/files", dependencies=[Depends(require_token)])
def local_files() -> JSONResponse:
    """input · output 폴더의 PPTX 목록과 폴더 위치. 화면의 목록·안내문이 쓴다."""
    local.ensure_dirs()
    return JSONResponse({
        "input_dir": str(config.INPUT_DIR),
        "output_dir": str(config.OUTPUT_DIR),
        "input_files": local.list_files(config.INPUT_DIR),
        "output_files": local.list_files(config.OUTPUT_DIR),
        "pick_available": local.pick_available(),
    })


@app.post("/api/local/jobs", dependencies=[Depends(require_token)])
def local_job(payload: dict = Body(...)) -> JSONResponse:
    """{path: 경로} 의 파일을 그 자리에서 읽어 분석을 시작한다. (복사하지 않는다)"""
    try:
        src = local.resolve_pptx(str(payload.get("path") or ""))
    except ValueError as e:
        raise HTTPException(400, str(e))
    opts = _options(str(payload.get("model") or config.MODEL),
                    bool(payload.get("think", False)),
                    bool(payload.get("scan_all", True)),
                    bool(payload.get("tag_marks", False)))
    job = _start(src.name, src, opts, "local")
    return JSONResponse({"job_id": job.job_id})


@app.post("/api/local/pick", dependencies=[Depends(require_token)])
async def local_pick(payload: dict | None = Body(default=None)) -> JSONResponse:
    """파이썬이 파일 선택창을 띄운다. 고른 경로 목록을 돌려준다 (취소하면 빈 목록)."""
    multiple = bool((payload or {}).get("multiple", False))
    try:
        # 창이 닫힐 때까지 기다리는 동안 다른 요청(진행률 조회 등)이 막히지 않게 스레드에서 돌린다
        paths = await asyncio.to_thread(local.pick_files, multiple)
    except RuntimeError as e:
        raise HTTPException(400, str(e))
    return JSONResponse({"paths": paths})


@app.post("/api/local/open", dependencies=[Depends(require_token)])
def local_open(payload: dict | None = Body(default=None)) -> JSONResponse:
    """탐색기(Finder)로 input 또는 output 폴더를 연다. select 가 있으면 그 파일을 선택한 채로."""
    which = (payload or {}).get("which", "output")
    folder = config.INPUT_DIR if which == "input" else config.OUTPUT_DIR
    sel_raw = (payload or {}).get("select")
    select = Path(str(sel_raw)) if sel_raw else None
    try:
        local.open_folder(folder, select)
    except Exception as e:  # noqa: BLE001  (열기가 막힌 PC 도 있다 — 경로만 알려 준다)
        raise HTTPException(400, f"폴더를 열지 못했습니다 ({type(e).__name__}). "
                                 f"탐색기 주소창에 직접 입력하세요: {folder}")
    return JSONResponse({"ok": True, "path": str(folder)})


@app.post("/api/local/review/preview", dependencies=[Depends(require_token)])
def local_review_preview(payload: dict = Body(...)) -> JSONResponse:
    """{paths: [경로, …]} 의 검토완료본을 읽어 파일별 교정 내역을 계산한다 (저장은 안 함)."""
    raws = [str(p) for p in (payload.get("paths") or []) if str(p).strip()]
    if not raws:
        raise HTTPException(400, "검토완료 PPTX 경로를 하나 이상 넣어 주세요.")
    named: list[tuple[str, str]] = []
    bad: list[dict] = []
    for raw in raws:
        try:
            p = local.resolve_pptx(raw)
            named.append((p.name, str(p)))
        except ValueError as e:          # 틀린 경로 하나가 나머지를 막지 않게 그 항목만 오류로
            bad.append({"filename": local.clean_path(raw) or raw, "error": str(e)})
    results = review.preview_many(named) if named else []
    return JSONResponse({"files": results + bad, "stats": review.stats()})


# ═════════════════════════════════════════════════════════════════
# [검토 반영] 탭 ─ 검토완료본 업로드 → 교정 내역 → 사전·예시 보강
# ═════════════════════════════════════════════════════════════════
@app.post("/api/review/preview")
async def review_preview(files: list[UploadFile] = File(default=[]),
                         file: UploadFile | None = File(default=None)) -> JSONResponse:
    """검토완료 PPTX(여러 개 가능)를 받아 파일별 교정 내역을 계산해 돌려준다.

    아직 아무것도 저장하지 않는다 — 화면에서 [반영 저장] 을 눌러야 기록된다.
    한 파일이 깨져 있거나 같은 문서가 겹쳐도 나머지 파일은 정상 처리된다.
    (file 은 예전 단일 업로드 형식과의 호환용)
    """
    ups = list(files) + ([file] if file is not None else [])
    if not ups:
        raise HTTPException(400, "PPTX 파일을 올려 주세요.")
    workdir = Path(tempfile.mkdtemp(prefix="pm-review-"))
    try:
        named: list[tuple[str, str]] = []
        for n, up in enumerate(ups):
            tmp = workdir / f"r{n}.pptx"
            with tmp.open("wb") as fh:
                shutil.copyfileobj(up.file, fh)
            named.append((up.filename or f"reviewed{n}.pptx", str(tmp)))
        results = review.preview_many(named)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        raise HTTPException(400, f"검토본을 읽지 못했습니다: {type(e).__name__}: {e}")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return JSONResponse({"files": results, "stats": review.stats()})


@app.post("/api/review/commit")
async def review_commit(payload: dict = Body(...)) -> JSONResponse:
    """미리보기에서 확인한 교정 내역을 저장하고, 무엇이 바뀌었는지 리포트를 돌려준다.

    여러 파일을 한 번에 받는다: {"files": [{filename, mode, items}, ...]}.
    예전 단일 형식 {filename, mode, items} 도 그대로 동작한다.
    응답의 change = 반영 전후 변화 요약 (변경 일지에도 같은 내용이 쌓인다).
    """
    batches = payload.get("files")
    if batches is None:                # 예전 단일 파일 형식
        batches = [payload]
    try:
        result = review.commit_with_log(batches)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return JSONResponse({"ok": True, **result})


@app.post("/api/review/rule")
async def review_rule(payload: dict = Body(...)) -> JSONResponse:
    """[사전에 추가] — 놓친 표현을 규칙 사전에 등록한다. 다음 분석부터 반드시 잡힌다."""
    try:
        st = review.add_rule(payload.get("keyword") or "")
    except ValueError as e:
        raise HTTPException(400, str(e))
    return JSONResponse({"ok": True, "stats": st})


@app.get("/api/review/stats")
def review_stats() -> JSONResponse:
    """누적 현황 — 데이터 건수·예시 수·추가된 사전 표현 수."""
    return JSONResponse(review.stats())


# ═════════════════════════════════════════════════════════════════
# [성능 기록] 탭 ─ 최초 설정 vs 현재 설정을 같은 문제지로 채점해 이력 누적
# ═════════════════════════════════════════════════════════════════
@app.post("/api/eval/run")
async def eval_run(payload: dict | None = Body(default=None)) -> JSONResponse:
    """성능 측정 시작. 문단 묶음마다 모델을 두 번(최초·현재) 부르므로 몇 분 걸린다."""
    try:
        evaluate.start(note=(payload or {}).get("note") or "")
    except (ValueError, RuntimeError) as e:
        raise HTTPException(400, str(e))
    return JSONResponse({"ok": True})


@app.get("/api/eval/status")
def eval_status() -> JSONResponse:
    """측정 진행 상황. 화면이 주기적으로 호출한다."""
    return JSONResponse(evaluate.status())


@app.post("/api/eval/cancel")
def eval_cancel() -> JSONResponse:
    """측정 중단 요청. 다음 문단 묶음으로 넘어갈 때 멈춘다."""
    evaluate.cancel()
    return JSONResponse({"ok": True})


@app.get("/api/eval/history")
def eval_history() -> JSONResponse:
    """누적 측정 이력 + 변경 일지 + 현재 평가셋 규모. [성능 기록] 탭이 그린다."""
    return JSONResponse({"entries": evaluate.history(),
                         "changes": review.change_log(),
                         "eval_set": evaluate.eval_set_summary(),
                         "stats": review.stats()})


# /static/... 주소로 static 폴더의 파일을 그대로 내어 준다
app.mount("/static", StaticFiles(directory=str(STATIC)), name="static")


def serve() -> None:
    """서버를 띄운다. run.sh / run.bat 이 `python -m app.main` 으로 이 함수를 부른다."""
    import uvicorn

    local.ensure_dirs()
    print(f"\n  특허 마킹 도구 (버전 {config.VERSION})  →  http://{config.HOST}:{config.PORT}")
    print(f"  온디바이스 모델 : {config.MODEL} @ {config.OLLAMA_HOST}")
    print(f"  파일 넣는 폴더  : {config.INPUT_DIR}")
    print(f"  결과 저장 폴더  : {config.OUTPUT_DIR}")
    print("  외부 네트워크로 나가는 통신은 없습니다.")
    print("  종료하려면 이 창에서 Ctrl + C\n")
    uvicorn.run(app, host=config.HOST, port=config.PORT, log_level="warning")


if __name__ == "__main__":
    serve()
