"""
tests/ci_windows.py ─ Windows 에서 설치·실행·점검이 끝까지 되는지 자동 확인
=====================================================================
회사 PC 와 같은 Windows 에서 실제로 돌려 보기 위한 스크립트입니다. GitHub 의 Windows 서버(Actions)가
저장소를 내려받아 이 파일을 실행합니다 (.github/workflows/windows-check.yml).
회사 PC 에서 직접 돌릴 일은 없습니다.

  진짜 Ollama 와 모델(5GB) 대신 가짜 Ollama(tests/mock_llm.py)를 11434 포트에 띄워 놓고
    1) setup.bat      설치 — 가상환경, 라이브러리, Ollama·모델 확인까지
    2) 점검 스크립트   tests/test_*.py 전부 + smoke.py  (출력 글자표는 한국어 Windows 처럼 cp949)
    3) mark.bat       파일을 아이콘에 끌어다 놓은 것처럼 (공백·한글 경로, PPTX 와 PDF)
    4) run.bat        웹 서버를 띄워 경로 입력·업로드로 분석, 결과 내려받기
  를 차례로 돌리고 결과 표를 남깁니다.

  여기서 확인하지 못하는 것: 진짜 모델의 판정과 속도, winget 으로 Ollama 를 설치하는 부분,
  브라우저 화면, PowerPoint·PDF 뷰어에서 보이는 모습, 회사 보안 프로그램의 영향.

  Windows 가 아닌 곳(Mac)에서 실행하면 배치 파일 대신 그 안에서 부르는 파이썬 명령을 바로 돌려
  이 스크립트 자체에 틀린 곳이 없는지만 봅니다 (설치 단계는 건너뜀, 폴더·포트는 따로 씀).
"""
from __future__ import annotations

import json
import locale
import os
import platform
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass

ROOT = Path(__file__).resolve().parents[1]
WIN = os.name == "nt"
VPY = ROOT / ".venv" / ("Scripts/python.exe" if WIN else "bin/python")
MOCK_PORT = 11434 if WIN else 11601          # 배치 파일은 11434 만 본다. Mac 에서는 진짜 Ollama 와 겹치지 않게
WEB_PORT = 8765 if WIN else 8791
TMP = ROOT / "ci_tmp"
TESTS = ["test_filter", "smoke", "test_review", "test_eval", "test_memory", "test_local_input",
         "test_model_call", "test_pdf", "test_prehl", "test_cp949"]
RESULTS: list[tuple[str, bool, str]] = []

# `where ollama` 와 `ollama list` 에 답하는 가짜 명령. 서버 자리는 아래에서 파이썬으로 따로 띄운다.
STUB = """@echo off
if /i "%~1"=="list" goto list
exit /b 0
:list
echo NAME                 ID      SIZE    MODIFIED
echo qwen3:8b             stub    5.2 GB  now
echo qwen3:4b-instruct    stub    2.5 GB  now
echo bge-m3:latest        stub    1.2 GB  now
exit /b 0
"""


def note(name: str, ok: bool, detail: str = "") -> bool:
    RESULTS.append((name, bool(ok), detail))
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  ({detail})" if detail else ""), flush=True)
    return bool(ok)


def decode(raw: bytes) -> str:
    """배치 파일의 echo 는 UTF-8(chcp 65001)로, 그 안에서 부른 파이썬은 cp949 로 적는다 — 줄마다 맞는 쪽으로 읽는다."""
    lines = []
    for line in raw.split(b"\n"):
        for enc in ("utf-8", "cp949"):
            try:
                lines.append(line.decode(enc))
                break
            except UnicodeDecodeError:
                continue
        else:
            lines.append(line.decode("utf-8", "replace"))
    return "\n".join(lines)


def run(cmd: list[str], env: dict, timeout: int, label: str) -> tuple[int, str]:
    print(f"\n{'─' * 20} {label} {'─' * 20}", flush=True)
    try:
        p = subprocess.run(cmd, cwd=str(ROOT), env=env, stdin=subprocess.DEVNULL,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout)
        rc, raw = p.returncode, p.stdout
    except subprocess.TimeoutExpired as e:
        rc, raw = -9, (e.stdout or b"") + "\n[시간 초과]".encode("utf-8")
    out = decode(raw)
    print(out[-5000:], flush=True)
    return rc, out


def get(url: str, timeout: float = 10, headers: dict | None = None, data: bytes | None = None):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method="POST" if data is not None else "GET")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.status, r.headers, r.read()


def wait_http(url: str, seconds: float) -> bool:
    end = time.time() + seconds
    while time.time() < end:
        try:
            get(url, 3)
            return True
        except Exception:  # noqa: BLE001
            time.sleep(0.5)
    return False


def stop(proc: subprocess.Popen) -> None:
    if WIN:                                      # run.bat 이 띄운 파이썬까지 함께 끝낸다
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True)
    else:
        proc.terminate()
    try:
        proc.wait(timeout=15)
    except subprocess.TimeoutExpired:
        proc.kill()


def wait_job(base: str, job: str, seconds: float = 180) -> dict:
    snap: dict = {}
    end = time.time() + seconds
    while time.time() < end:
        snap = json.loads(get(f"{base}/api/jobs/{job}")[2])
        if snap.get("status") in ("done", "error"):
            break
        time.sleep(0.3)
    return snap


def main() -> int:
    if not WIN:
        print("Windows 가 아닙니다 — 배치 파일은 건너뛰고 같은 파이썬 명령으로 이 스크립트만 점검합니다.")
    version = re.search(r'^VERSION = "([^"]+)"', (ROOT / "app" / "config.py").read_text(encoding="utf-8"), re.M).group(1)
    env = dict(os.environ)
    for k in ("PYTHONIOENCODING", "PYTHONUTF8", "PM_MOCK_PORT"):
        env.pop(k, None)
    shutil.rmtree(TMP, ignore_errors=True)
    TMP.mkdir()
    out_dir = ROOT / "output"
    if not WIN:                                  # 이 PC 의 output·review_data 를 건드리지 않게
        out_dir = TMP / "output"
        env.update({"PM_OLLAMA_HOST": f"http://127.0.0.1:{MOCK_PORT}", "PM_PORT": str(WEB_PORT),
                    "PM_OUTPUT_DIR": str(out_dir), "PM_INPUT_DIR": str(TMP / "input"),
                    "PM_REVIEW_DIR": str(TMP / "review_data")})
    # 배치 파일 안의 파이썬은 한국어 Windows 에서 출력이 파일·파이프로 갈 때처럼 cp949 로 적게 한다
    # (모아 두지 않고 바로 적게 해서, 서버를 강제로 끝내도 찍힌 줄이 남게 한다)
    benv = {**env, "PYTHONIOENCODING": "cp949", "PYTHONUNBUFFERED": "1"}
    info = [f"도구 버전 {version}", platform.platform(), f"Python {platform.python_version()}",
            f"시스템 글자표 {locale.getpreferredencoding(False)}"]
    print(" · ".join(info), flush=True)

    stub = ROOT / "ci_stub"
    stub.mkdir(exist_ok=True)
    (stub / "ollama.cmd").write_bytes(STUB.replace("\n", "\r\n").encode("ascii"))
    env["PATH"] = str(stub) + os.pathsep + env.get("PATH", "")

    mock = subprocess.Popen([sys.executable, "-u", str(ROOT / "tests" / "mock_llm.py"), "--serve"], cwd=str(ROOT),
                            env={**env, "PM_MOCK_PORT": str(MOCK_PORT)}, stdin=subprocess.DEVNULL,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    libs = ""
    srv = None
    try:
        if not note(f"준비: 가짜 Ollama 가 {MOCK_PORT} 에서 응답",
                    wait_http(f"http://127.0.0.1:{MOCK_PORT}/api/tags", 30)):
            return finish(info, libs)

        # ── 1. setup.bat ────────────────────────────────────────────
        vpy = VPY
        if WIN:
            shutil.rmtree(ROOT / ".venv", ignore_errors=True)
            rc, out = run(["cmd", "/c", "setup.bat"], env, 1800, "setup.bat")
            note("setup.bat: 끝까지 실행 (종료 코드 0)", rc == 0, f"종료 코드 {rc}")
            note("setup.bat: 가상환경(.venv) 생성", vpy.exists())
            steps = [k for k in ("[1/5]", "[2/5]", "[3/5]", "[4/5]", "[5/5]", "[5/5-2]", "[5/5-3]", "설치 완료") if k in out]
            note("setup.bat: 다섯 단계와 '설치 완료' 가 한글로 찍힘", len(steps) == 8 and "가상환경" in out, " ".join(steps))
            note("setup.bat: 모델 셋을 '이미 있음' 으로 인식", out.count("이미 있음") >= 3, f"{out.count('이미 있음')}회")
        if not vpy.exists():
            note("준비: 가상환경의 파이썬", False, str(vpy))
            return finish(info, libs)
        rc, out = run([str(vpy), "-c",
                       "import importlib.metadata as m; print(' · '.join(f'{n} {m.version(n)}' for n in "
                       "('python-pptx','lxml','fastapi','uvicorn','python-multipart','pdfminer.six','pypdf','cryptography')))"],
                      env, 120, "설치된 라이브러리")
        libs = out.strip().splitlines()[-1] if rc == 0 and out.strip() else ""
        note("라이브러리: PDF 용 둘(pdfminer.six · pypdf)까지 설치됨", "pdfminer.six 20250506" in libs and "pypdf 6.1.3" in libs, libs)

        # ── 2. 점검 스크립트 (출력 글자표를 한국어 Windows 처럼) ──────────
        tenv = {**env, "PYTHONIOENCODING": "cp949"}
        for t in TESTS:
            rc, out = run([str(vpy), f"tests/{t}.py"], tenv, 900, f"tests/{t}.py")
            n_pass, n_fail = len(re.findall(r"^PASS", out, re.M)), len(re.findall(r"^FAIL", out, re.M))
            last = next((ln.strip() for ln in reversed(out.splitlines()) if ln.strip()), "")
            note(f"점검 {t}", rc == 0 and n_fail == 0, f"통과 {n_pass} · 실패 {n_fail} · {last[:70]}")

        # ── 3. mark.bat (끌어다 놓기) ─────────────────────────────────
        work = TMP / "끌어다 놓기 시험 (공백 있음)"
        work.mkdir(parents=True)
        pptx, pdf = work / "보고서 초안.pptx", work / "변환본.pdf"
        shutil.copyfile(ROOT / "samples" / "회사보고자료_예시.pptx", pptx)
        rc, out = run([str(vpy), "tools/make_sample_pdf.py", str(pptx), str(pdf)], env, 120, "점검용 PDF 만들기")
        note("준비: 점검용 PDF 생성", rc == 0 and pdf.exists())
        mark = ["cmd", "/c", "mark.bat"] if WIN else [str(vpy), "-m", "app.cli"]
        rc, out = run(mark + [str(pptx), str(pdf)], benv, 900, "mark.bat (파일 두 개)")
        o1, o2 = work / "보고서 초안_특허마킹.pptx", work / "변환본_특허마킹.pdf"
        note("mark.bat: 끝까지 실행 (종료 코드 0)", rc == 0, f"종료 코드 {rc}")
        note("mark.bat: 원본 옆에 결과 두 개 저장 (공백·한글 경로)", o1.exists() and o2.exists(),
             ", ".join(p.name for p in work.iterdir()))
        note("mark.bat: 검은 창에 버전·완료·저장 줄이 한글로 찍힘", version in out and "완료" in out and "저장" in out)
        chk = ("import sys, zipfile; from pypdf import PdfReader; "
               "z = zipfile.ZipFile(sys.argv[1]); n = sum(z.read(i).count(b'<a:highlight') for i in z.namelist() if i.startswith('ppt/slides/slide')); "
               "a = sum(1 for pg in PdfReader(sys.argv[2]).pages for x in (pg.get('/Annots') or []) if x.get_object().get('/Subtype') == '/Highlight'); "
               "print(n, a)")
        rc, out = run([str(vpy), "-c", chk, str(o1), str(o2)], env, 120, "결과 파일의 형광펜 세기")
        nums = re.findall(r"\d+", out.strip().splitlines()[-1]) if rc == 0 and out.strip() else []
        note("mark.bat 결과: PPTX 와 PDF 에 형광펜이 들어 있음", len(nums) == 2 and int(nums[0]) > 0 and int(nums[1]) > 0,
             f"PPTX 형광펜 런 {nums[0] if nums else '?'} · PDF 주석 {nums[1] if len(nums) > 1 else '?'}")

        # ── 4. run.bat (웹 서버) ─────────────────────────────────────
        log = open(TMP / "run_bat.log", "wb")
        srv = subprocess.Popen(["cmd", "/c", "run.bat"] if WIN else [str(vpy), "-m", "app.main"], cwd=str(ROOT),
                               env=benv, stdin=subprocess.DEVNULL, stdout=log, stderr=subprocess.STDOUT)
        base = f"http://127.0.0.1:{WEB_PORT}"
        up = wait_http(base + "/api/health", 120)
        note(f"run.bat: 서버가 {WEB_PORT} 에서 응답", up)
        if up:
            health = json.loads(get(base + "/api/health")[2])
            note("run.bat: 상태 확인 — 모델 준비됨 · 버전 일치", bool(health.get("ok")) and health.get("version") == version
                 and bool(health.get("model_ready")), f"version {health.get('version')} · model_ready {health.get('model_ready')}")
            page = get(base + "/")[2].decode("utf-8")
            token = page.split('const TOKEN = "', 1)[1].split('"', 1)[0] if 'const TOKEN = "' in page else ""
            note("run.bat: 첫 화면에 버전·새 체크박스", version in page and 'id="optStrip"' in page and len(token) > 10)
            st, _, raw = get(base + "/api/local/jobs", 60, {"Content-Type": "application/json", "X-PM-Token": token},
                             json.dumps({"path": str(pptx), "model": "qwen3:8b"}).encode("utf-8"))
            snap = wait_job(base, json.loads(raw)["job_id"])
            note("run.bat: 경로로 넣은 PPTX 분석 완료", snap.get("status") == "done", str(snap.get("error") or snap.get("stage")))
            outs = sorted(p.name for p in out_dir.glob("*특허마킹*"))
            note("run.bat: output 폴더에 결과 저장", any(n.endswith(".pptx") for n in outs), ", ".join(outs))
            b = "----pmci"
            body = (f"--{b}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"업로드 시험.pdf\"\r\n"
                    f"Content-Type: application/pdf\r\n\r\n").encode("utf-8") + pdf.read_bytes() + f"\r\n--{b}--\r\n".encode()
            st, _, raw = get(base + "/api/jobs", 60, {"Content-Type": f"multipart/form-data; boundary={b}"}, body)
            job = json.loads(raw)["job_id"]
            snap = wait_job(base, job)
            note("run.bat: 업로드한 PDF 분석 완료", snap.get("status") == "done", str(snap.get("error") or snap.get("stage")))
            st, hd, raw = get(f"{base}/api/jobs/{job}/download", 60)
            note("run.bat: 결과 내려받기 (PDF)", raw[:4] == b"%PDF" and "application/pdf" in hd.get("content-type", ""),
                 f"{len(raw):,} bytes")
        stop(srv)
        srv = None
        log.close()
        text = decode((TMP / "run_bat.log").read_bytes())
        print(f"\n{'─' * 20} run.bat 검은 창 {'─' * 20}\n{text[-3000:]}", flush=True)
        note("run.bat: 검은 창에 안내가 한글로 찍히고 오류가 없음",
             "특허 마킹 도구" in text and "통신은 없습니다" in text and "Traceback" not in text)
    except Exception as e:  # noqa: BLE001
        note("점검 스크립트 자체 오류", False, f"{type(e).__name__}: {e}")
    finally:
        if srv is not None:
            stop(srv)
        mock.terminate()
        if not WIN:
            shutil.rmtree(TMP, ignore_errors=True)
            shutil.rmtree(ROOT / "ci_stub", ignore_errors=True)
    return finish(info, libs)


def finish(info: list[str], libs: str) -> int:
    bad = [r for r in RESULTS if not r[1]]
    lines = [f"## Windows 점검 — {'모두 통과' if not bad else f'실패 {len(bad)}건'} ({len(RESULTS) - len(bad)}/{len(RESULTS)})",
             "", " · ".join(info), "", f"설치된 라이브러리: {libs}" if libs else "", "",
             "| 결과 | 항목 | 내용 |", "|---|---|---|"]
    for name, ok, detail in RESULTS:
        lines.append(f"| {'통과' if ok else '**실패**'} | {name} | {detail.replace('|', '/')} |")
    text = "\n".join(lines) + "\n"
    print("\n" + text, flush=True)
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if path:
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(text)
    return 1 if bad or not RESULTS else 0


if __name__ == "__main__":
    sys.exit(main())
