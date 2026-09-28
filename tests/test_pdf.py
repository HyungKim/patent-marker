"""
tests/test_pdf.py ─ PDF 입력(파워포인트를 PDF 로 변환한 보고서) 경로 점검 (LLM 없이, 30초)
=====================================================================
예시 PPTX 의 글자를 그대로 담은 PDF 를 tools/make_sample_pdf.py 로 만들어(저장소에 PDF 를 두지 않으므로 매번 생성)
  읽기(pdfdoc.extract) → 파이프라인(형광펜 주석 + 범례) → 검토본 읽기 → 검토 시뮬레이션(지움·추가·색 변경)
  → 명령행(app/cli.py) → 웹(업로드·내려받기·검토본 업로드) 이 끊기지 않는지 확인합니다.
같은 문서의 PPTX 결과와도 비교해, PDF 로 넣어도 같은 후보가 나오는지 봅니다.
samples/회사보고자료_예시.pdf 가 따로 있으면(예: PowerPoint 나 Chrome 으로 찍은 진짜 변환본) 그 파일의 읽기도 함께 점검합니다.
모델 자리는 tests/mock_llm.py 의 가짜 Ollama 가 대신합니다.

    실행:  python tests/test_pdf.py
"""
from __future__ import annotations

import json
import os
import re
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

for _s in (sys.stdout, sys.stderr):                    # Windows 콘솔(cp949) 대비
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

# 실제 review_data · input · output 을 건드리지 않도록 import 전에 임시 폴더로 돌린다
TMP = Path(tempfile.mkdtemp(prefix="pm-pdf-test-"))
os.environ["PM_REVIEW_DIR"] = str(TMP / "review_data")
os.environ["PM_INPUT_DIR"] = str(TMP / "input")
os.environ["PM_OUTPUT_DIR"] = str(TMP / "output")

import mock_llm  # noqa: E402

MOCK_HOST = mock_llm.start()
os.environ["PM_OLLAMA_HOST"] = MOCK_HOST

from pypdf import PdfReader, PdfWriter                              # noqa: E402
from pypdf.annotations import Highlight                              # noqa: E402
from pypdf.generic import ArrayObject, FloatObject, NameObject       # noqa: E402

from app import analyze, config, extract, local, merge, pdfdoc, pipeline, review  # noqa: E402

config.OLLAMA_HOST = MOCK_HOST
analyze.config.OLLAMA_HOST = MOCK_HOST

sys.path.insert(0, str(ROOT / "tools"))
import make_sample_pdf  # noqa: E402

PPTX = ROOT / "samples" / "회사보고자료_예시.pptx"
PDF = TMP / "회사보고자료_예시.pdf"                 # 생성한 간이 PDF (글자·좌표만, 5쪽)
REAL_PDF = ROOT / "samples" / "회사보고자료_예시.pdf"   # 진짜 변환본이 있으면 추가 점검
FAILS = 0


def check(label: str, ok: bool, detail: str = "") -> None:
    global FAILS
    print(f"{'PASS' if ok else 'FAIL'}  {label}" + (f"  ({detail})" if detail else ""))
    if not ok:
        FAILS += 1


def norm(t: str) -> str:
    """공백과, PDF 표준 한글 글꼴에 없어 바뀌는 글자(가운뎃점·줄표)를 빼고 비교한다."""
    return re.sub(r"[\s·•—―–]+", "", t)


def annots_of(path: Path) -> list[tuple[int, dict]]:
    """(쪽 번호, 주석 객체) 목록."""
    out = []
    for pno, page in enumerate(PdfReader(str(path)).pages, 1):
        for a in page.get("/Annots") or []:
            out.append((pno, a.get_object()))
    return out


# ── 0. 간이 PDF 생성 (tools/make_sample_pdf.py) ─────────────────────
n_pages, n_paras = make_sample_pdf.build(PPTX, PDF)
check("간이 PDF 생성: 5쪽 · 문단 106개 · 10KB 미만", n_pages == 5 and n_paras == 106 and PDF.stat().st_size < 10_000,
      f"{n_pages}쪽 · {n_paras}문단 · {PDF.stat().st_size:,} bytes")

# ── 1. 읽기: PDF → 문단 + 글자 좌표, PPTX 원본과 대조 ───────────────
deck = extract.extract(str(PDF))
check("extract.extract 가 .pdf 를 PdfDeck 으로 읽는다", isinstance(deck, pdfdoc.PdfDeck) and deck.is_pdf)
check("쪽 수 5 (슬라이드 = 쪽) · 글상자 106개 (문단 = 글상자)", deck.slide_count == 5 and len(deck.segments) == 106,
      f"{deck.slide_count}쪽 · {len(deck.segments)}상자")
check("줄바꿈된 글상자가 있다 (여러 줄을 한 문단으로 잇는 경로)", sum(1 for s in deck.segments if len(s.text) > 70) >= 3)
check("문단마다 글자 수 = 좌표 수", all(len(deck.boxes[s.seg_id]) == len(s.text) for s in deck.segments))
check("한글이 깨지지 않는다", sum(1 for s in deck.segments if re.search("[가-힣]", s.text)) >= len(deck.segments) * 0.6)
inside = all(
    b is None or (-1 <= b[0] <= b[2] <= deck.pages[s.slide_no][0] + 1 and -1 <= b[1] <= b[3] <= deck.pages[s.slide_no][1] + 1)
    for s in deck.segments for b in deck.boxes[s.seg_id])
check("글자 좌표가 쪽 안에 있다", inside)
check("주소 표기 p쪽/글상자n", all(re.fullmatch(r"p\d+/글상자\d+", s.addr) for s in deck.segments))

pdeck = extract.extract(str(PPTX))
pp = [norm(s.text) for s in pdeck.segments if s.kind in ("title", "body", "table") and s.text.strip()]
exact = sum(1 for t in pp if any(t == norm(s.text) for s in deck.segments))
check(f"PPTX 문단 {len(pp)}개가 PDF 문단과 하나하나 같다", exact == len(pp), f"{exact}/{len(pp)}")

if REAL_PDF.exists():      # 진짜 변환본(PowerPoint·Chrome 등)이 있으면: 읽기·문단 포함 여부만 확인
    rdeck = extract.extract(str(REAL_PDF))
    r_all = "\x1f".join(norm(s.text) for s in rdeck.segments)
    found = sum(1 for t in pp if t in r_all)
    check(f"[추가] 진짜 변환본 {REAL_PDF.name}: {rdeck.slide_count}쪽 · 글상자 {len(rdeck.segments)} · PPTX 문단 포함",
          found == len(pp) and all(len(rdeck.boxes[s.seg_id]) == len(s.text) for s in rdeck.segments), f"{found}/{len(pp)}")

# ── 2. 파일 넣기 경로: 목록 · 경로 확정 · 결과 이름 ─────────────────
local.ensure_dirs()
src = config.INPUT_DIR / "변환보고서.pdf"
src.write_bytes(PDF.read_bytes())
(config.INPUT_DIR / "메모.txt").write_text("x", encoding="utf-8")
names = [r["name"] for r in local.list_files(config.INPUT_DIR)]
check("input 목록에 PDF 가 나온다 (txt 는 제외)", names == ["변환보고서.pdf"], str(names))
check("경로 확정: 파일 이름만으로 input 의 PDF 를 찾는다", local.resolve_pptx("변환보고서.pdf") == src.resolve())
p1 = local.output_path_for("변환보고서.pdf")
check("결과 이름: 이름_특허마킹.pdf", p1 == config.OUTPUT_DIR / "변환보고서_특허마킹.pdf", p1.name)
p1.write_bytes(b"x")
check("결과 이름: 겹치면 (2).pdf", local.output_path_for("변환보고서.pdf").name == "변환보고서_특허마킹(2).pdf")
p1.unlink()
check("PPTX 결과 이름은 그대로 .pptx", local.output_path_for("보고서.pptx").name == "보고서_특허마킹.pptx")

# ── 3. 파이프라인: 형광펜 주석 + 1쪽 범례 ───────────────────────────
stages: list[str] = []
dst = config.OUTPUT_DIR / "변환보고서_특허마킹.pdf"
resolved, stats = pipeline.run(src, dst, config.RunOptions(), progress=lambda p: stages.append(p.stage))
check("결과 PDF 저장", dst.exists() and dst.stat().st_size > 10_000, f"{dst.stat().st_size if dst.exists() else 0} bytes")
check("후보·집계 (문구 삽입은 PDF 에 없음 → tags 0)",
      stats["total"] == len(resolved) > 0 and stats["legend"] == 1 and stats["tags"] == 0,
      f"후보 {stats['total']} · 범례 {stats['legend']} · 문구 {stats['tags']}")
check("단계 보고: 'PDF 마킹 중' → 완료", "PDF 마킹 중" in stages and stages[-1] == "완료")
check("run_log.tsv 에 기록", (config.REVIEW_DIR / "run_log.tsv").exists()
      and "변환보고서.pdf" in (config.REVIEW_DIR / "run_log.tsv").read_text(encoding="utf-8-sig"))

annots = annots_of(dst)
hl = [(p, o) for p, o in annots if o.get("/Subtype") == "/Highlight"]
check("쪽 수 유지", len(PdfReader(str(dst)).pages) == 5)
check("형광펜 주석 수 = 칠한 구간 수 = 계획한 구간 수", len(hl) == stats["runs_painted"] == stats["marks"] > 0,
      f"주석 {len(hl)} · 칠함 {stats['runs_painted']} · 계획 {stats['marks']}")
grade_hex = {v.upper() for v in config.GRADE_COLOR.values()}
check("형광펜 색은 등급 색 셋 중 하나", all(pdfdoc._hex(o.get("/C")) in grade_hex for _, o in hl))
check("형광펜 메모에 '출원검토필요' + [등급] 분류",
      all(str(o.get("/Contents", "")).startswith("출원검토필요\n[") for _, o in hl))
check("형광펜 이름표 = 출원검토필요", all(str(o.get("/T")) == "출원검토필요" for _, o in hl))
first_page = [o for p, o in annots if p == 1]
check("1쪽 범례: 메모 1개 + 색 견본 3개",
      sum(o.get("/Subtype") == "/Text" for o in first_page) == 1
      and sum(o.get("/Subtype") == "/Square" for o in first_page) == 3)
check("범례 주석 이름표 = LEGEND_NAME",
      all(str(o.get("/T")) == config.LEGEND_NAME for o in first_page if o.get("/Subtype") in ("/Text", "/Square")))
legend_note = next(o for o in first_page if o.get("/Subtype") == "/Text")
check("범례 메모에 안내 제목·등급 설명", config.LEGEND_TITLE in str(legend_note.get("/Contents"))
      and all(config.GRADE_LABEL[g] in str(legend_note.get("/Contents")) for g in "ABC"))

# 같은 문서의 PPTX 결과와 비교 — PDF 로 넣어도 같은 후보가 나와야 한다 (쪽 = 슬라이드, 문단 글자열도 같음).
# PPTX 쪽은 발표자 노트(PDF 변환본에는 없는 부분)의 후보를 빼고, 문단·등급·공개 단위로 비교한다
# (가운뎃점 · 이 PDF 표준 글꼴에 없어 • 로 바뀌면 규칙 사전 인용구의 끝자리가 한 문단에서 달라질 수 있어 인용구 자체는 참고로만 센다).
pres, pstats = pipeline.run(PPTX, config.OUTPUT_DIR / "원본_특허마킹.pptx", config.RunOptions())
pkind = {s.seg_id: s for s in pdeck.segments}
pres_body = [f for f in pres if pkind[f.seg_id].kind in ("title", "body", "table")]
pkey = lambda f, segs: (f.slide_no, norm(segs[f.seg_id].text), f.grade, bool(f.disclosure_risk))   # noqa: E731
dseg = {s.seg_id: s for s in deck.segments}
kp, kx = {pkey(f, dseg) for f in resolved}, {pkey(f, pkind) for f in pres_body}
qp, qx = {(f.slide_no, norm(f.quote)) for f in resolved}, {(f.slide_no, norm(f.quote)) for f in pres_body}
check("PPTX 결과(노트 제외)와 같은 후보 — 슬라이드·문단·등급·공개 전부 일치, 건수 같음",
      kp == kx and len(resolved) == len(pres_body),
      f"PDF {len(resolved)} · PPTX {len(pres_body)}(노트 {len(pres) - len(pres_body)}건 제외) · 인용구까지 같은 것 {len(qp & qx)}"
      + (f" · PDF만 {sorted(kp - kx)[:2]} · PPTX만 {sorted(kx - kp)[:2]}" if kp != kx else ""))

# ── 4. 검토본 읽기 (고치지 않은 마킹본) ─────────────────────────────
rv = review.read_reviewed(str(dst))
fp = review.deck_fingerprint(deck)
check("지문이 분석 기록과 같다 (기록을 찾을 수 있다)", rv["fingerprint"] == fp and review.load_archive(fp) is not None)
check("쪽 수·문단 수", rv["slide_count"] == 5 and len(rv["paras"]) == len(deck.segments))
nspans = sum(len(p["spans"]) for p in rv["paras"])
check("읽은 형광펜 구간 수 = 칠한 구간 수", nspans == stats["runs_painted"], f"{nspans} vs {stats['runs_painted']}")

marks = merge.build_marks(deck, resolved)
para_of = {(p["slide_no"], p["text"]): p for p in rv["paras"]}
exact_pos = 0
for m in marks:
    seg = deck.get(m.seg_id)
    want = m.span or (0, len(seg.text))
    p = para_of.get((seg.slide_no, seg.text))
    if p and any(abs(sp["start"] - want[0]) <= 1 and abs(sp["end"] - want[1]) <= 1 for sp in p["spans"]):
        exact_pos += 1
check("형광펜 자리가 글자 단위로 되살아난다 (±1글자)", exact_pos == len(marks), f"{exact_pos}/{len(marks)}")
grades_back = {review.color_to_grade(sp["color"]) for p in rv["paras"] for sp in p["spans"]}
check("색 → 등급 복원", grades_back <= {("A", False), ("B", False), ("B", True)}, str(grades_back))

d0 = review.diff(rv)
check("고치지 않은 검토본: 오탐·누락·등급변경 0", d0["mode"] == "archive"
      and d0["counts"]["fp"] == 0 and d0["counts"]["miss"] == 0 and d0["counts"]["regrade"] == 0
      and d0["counts"]["match"] >= 1, json.dumps(d0["counts"], ensure_ascii=False))

# ── 5. 검토 시뮬레이션: 형광펜 하나 지움(오탐) · 새로 칠함(누락, 하늘색=B) · 색 바꿈(노랑→살구=C) ─
reader = PdfReader(str(dst))
writer = PdfWriter(clone_from=reader)
removed_quote = regraded_quote = None
for pno, page in enumerate(writer.pages, 1):
    arr = page.get("/Annots")
    if not arr:
        continue
    for idx, a in enumerate(list(arr)):
        o = a.get_object()
        if o.get("/Subtype") != "/Highlight":
            continue
        if removed_quote is None:
            removed_quote = str(o["/Contents"]).split("\n")[1]     # "[B] 라벨 · 분류 · 사유" — 무엇을 지웠는지 참고용
            del arr[idx]
            break
    if removed_quote is not None:
        break
# 지운 구간의 인용구: 원본 읽기와 수정본 읽기의 차이로 찾는다
mod = TMP / "검토본.pdf"
with open(mod, "wb") as fh:
    writer.write(fh)
rv1 = review.read_reviewed(str(mod))
lost = [(p["slide_no"], p["text"][sp["start"]:sp["end"]].strip())
        for p0, p in zip(rv["paras"], rv1["paras"]) for sp in p0["spans"]
        if sp not in p["spans"]]
check("형광펜 하나를 지웠다", len(lost) == 1, str(lost)[:80])

# 누락: 형광펜이 없는 문단 하나를 골라 앞 12글자를 하늘색으로 칠한다
cand = next(s for s in deck.segments if len(s.text) >= 20 and not any(m.seg_id == s.seg_id for m in marks))
rects = pdfdoc._line_rects(deck.boxes[cand.seg_id], (0, 12))
writer = PdfWriter(clone_from=PdfReader(str(mod)))
new_hl = Highlight(rect=pdfdoc._union(rects), quad_points=pdfdoc._quads(rects), highlight_color="4FC3F7")
writer.add_annotation(page_number=cand.slide_no - 1, annotation=new_hl)
# 등급 변경: 노랑(A) 형광펜 하나를 살구(C)로
for page in writer.pages:
    done = False
    for a in page.get("/Annots") or []:
        o = a.get_object()
        if o.get("/Subtype") == "/Highlight" and pdfdoc._hex(o.get("/C")) == config.GRADE_COLOR["A"]:
            o[NameObject("/C")] = ArrayObject([FloatObject(1.0), FloatObject(0.62), FloatObject(0.50)])
            regraded_quote = str(o["/Contents"]).split("\n")[1]
            done = True
            break
    if done:
        break
with open(mod, "wb") as fh:
    writer.write(fh)

d1 = review.diff(review.read_reviewed(str(mod)))
c = d1["counts"]
miss = [i for i in d1["items"] if i["type"] == "miss"]
fps = [i for i in d1["items"] if i["type"] == "fp"]
regr = [i for i in d1["items"] if i["type"] == "regrade"]
check("교정 내역: 오탐 1 · 누락 1 · 등급변경 1", c["fp"] == 1 and c["miss"] == 1 and c["regrade"] == 1,
      json.dumps(c, ensure_ascii=False))
check("오탐 = 지운 구간", fps and norm(fps[0]["quote"]) == norm(lost[0][1]), fps[0]["quote"][:40] if fps else "")
check("누락 = 새로 칠한 12글자, 하늘색 → B", miss and miss[0]["quote"] == cand.text[:12].strip()
      and miss[0]["grade"] == "B" and not miss[0]["risk"], str(miss[0])[:100] if miss else "")
check("등급변경 = 살구색 → C(공개 관련정보)", regr and regr[0]["risk"] and regr[0]["from_grade"] == "A",
      str({k: regr[0][k] for k in ("from_grade", "grade", "risk")}) if regr else "")
check("검토 반영 미리보기가 PDF 를 받는다 (.txt 는 거절)",
      review.preview_many([("검토본.pdf", str(mod))])[0]["counts"]["miss"] == 1
      and "PDF" in review.preview_many([("메모.txt", str(mod))])[0]["error"])

# ── 6. 명령행 ────────────────────────────────────────────────────
env = {**os.environ, "PM_OLLAMA_HOST": MOCK_HOST, "PM_OUTPUT_DIR": str(config.OUTPUT_DIR), "PM_REVIEW_DIR": str(config.REVIEW_DIR)}
proc = subprocess.run([sys.executable, "-m", "app.cli", str(src), "--out", str(TMP / "cli_out")],
                      cwd=str(ROOT), env=env, capture_output=True, text=True, encoding="utf-8")
cli_out = TMP / "cli_out" / "변환보고서_특허마킹.pdf"
check("app.cli 가 PDF 를 받아 _특허마킹.pdf 를 만든다", proc.returncode == 0 and cli_out.exists()
      and cli_out.read_bytes()[:4] == b"%PDF", (proc.stderr or proc.stdout)[-200:].strip() if proc.returncode else "")

# ── 7. 웹: 업로드 → 진행 → 내려받기 · 검토본 업로드 ──────────────────
import uvicorn  # noqa: E402
from app import main as webmain  # noqa: E402

with socket.socket() as sk:
    sk.bind(("127.0.0.1", 0))
    port = sk.getsockname()[1]
server = uvicorn.Server(uvicorn.Config(webmain.app, host="127.0.0.1", port=port, log_level="warning"))
threading.Thread(target=server.run, daemon=True).start()
for _ in range(100):
    if server.started:
        break
    time.sleep(0.1)
BASE = f"http://127.0.0.1:{port}"


def multipart(fields: dict, files: list[tuple[str, str, bytes]]) -> tuple[bytes, str]:
    b = "----pmtest"
    body = b""
    for k, v in fields.items():
        body += f"--{b}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode()
    for field, name, data in files:
        body += (f"--{b}\r\nContent-Disposition: form-data; name=\"{field}\"; filename=\"{name}\"\r\n"
                 f"Content-Type: application/octet-stream\r\n\r\n").encode() + data + b"\r\n"
    return body + f"--{b}--\r\n".encode(), f"multipart/form-data; boundary={b}"


def post(path: str, body: bytes, ctype: str):
    req = urllib.request.Request(BASE + path, data=body, headers={"Content-Type": ctype}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, r.headers, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers, e.read()


body, ctype = multipart({"model": config.MODEL, "think": "false", "scan_all": "false", "tag_marks": "false"},
                        [("file", "웹업로드.pdf", PDF.read_bytes())])
st, _, raw = post("/api/jobs", body, ctype)
job = json.loads(raw).get("job_id") if st == 200 else None
check("업로드(.pdf) 를 받는다", st == 200 and bool(job), raw[:120].decode("utf-8", "replace"))
snap = {}
for _ in range(600):
    with urllib.request.urlopen(f"{BASE}/api/jobs/{job}", timeout=10) as r:
        snap = json.loads(r.read())
    if snap.get("status") in ("done", "error"):
        break
    time.sleep(0.2)
check("업로드 작업 완료", snap.get("status") == "done", snap.get("error") or snap.get("stage", ""))
with urllib.request.urlopen(f"{BASE}/api/jobs/{job}/download", timeout=30) as r:
    dl_type, dl_disp, dl = r.headers.get("content-type", ""), r.headers.get("content-disposition", ""), r.read()
check("내려받기: application/pdf · 이름 _특허마킹.pdf · %PDF 로 시작",
      dl_type.startswith("application/pdf") and ".pdf" in dl_disp and dl[:4] == b"%PDF"
      and "웹업로드_특허마킹" in urllib.parse.unquote(dl_disp), f"{dl_type} · {dl_disp[:80]}")
(TMP / "dl.pdf").write_bytes(dl)
check("내려받은 파일에도 형광펜 주석", any(o.get("/Subtype") == "/Highlight" for _, o in annots_of(TMP / "dl.pdf")))
st, _, raw = post("/api/jobs", *multipart({}, [("file", "메모.txt", b"x")]))
check("업로드(.txt) 는 거절하며 PDF 를 안내", st == 400 and "PDF" in raw.decode("utf-8", "replace"))
st, _, raw = post("/api/review/preview", *multipart({}, [("files", "검토본.pdf", mod.read_bytes())]))
res = (json.loads(raw).get("files") or [{}]) if st == 200 else [{}]
check("검토본(.pdf) 업로드 → 교정 내역", st == 200 and res[0].get("counts", {}).get("miss") == 1
      and res[0]["counts"]["fp"] == 1, raw[:120].decode("utf-8", "replace"))
server.should_exit = True

print(f"\n{'모두 통과' if not FAILS else f'실패 {FAILS}건'}  (임시 폴더: {TMP})")
sys.exit(1 if FAILS else 0)
