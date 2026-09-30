"""
tests/ci_windows.py ─ Windows 에서 설치·실행·점검이 끝까지 되는지 자동 확인
=====================================================================
회사 PC 와 같은 Windows 에서 실제로 돌려 보기 위한 스크립트입니다. GitHub 의 Windows 서버(Actions)가
회사 PC 와 같은 방법(Download ZIP → 압축 풀기 → C:\\patent_marker)으로 받아 이 파일을 실행합니다
(.github/workflows/windows-check.yml). 회사 PC 에서 직접 돌릴 일은 없습니다.

  진짜 Ollama 와 모델(5GB) 대신 가짜 Ollama(tests/mock_llm.py)를 11434 포트에 띄워 놓고
    0) 받은 파일      한글 이름 파일이 그대로 풀렸는지, 배치 파일 줄바꿈이 Windows 식(CRLF)인지
    1) setup.bat      설치 — 가상환경, 라이브러리, Ollama·모델 확인까지.
                      이어서 PDF 라이브러리 둘을 지운 뒤 다시 실행 (예전 판에서 올라오는 PC 의 업데이트 경로)
    2) 점검 스크립트   tests/test_*.py 전부 + smoke.py  (출력 글자표는 한국어 Windows 처럼 cp949,
                      임시 폴더는 한글 사용자 이름처럼 한글·공백이 든 경로)
    3) mark.bat       탐색기가 파일을 끌어다 놓을 때 만드는 명령줄 그대로 (공백·한글·괄호·& 가 든 이름, PPTX 와 PDF)
    4) run.bat        웹 서버를 띄워 경로 입력·업로드로 분석, 결과 내려받기
  를 차례로 돌리고 결과 표를 남깁니다.

  여기서 확인하지 못하는 것: 진짜 모델의 판정과 속도, winget 으로 Ollama 를 설치하는 부분,
  브라우저 화면, PowerPoint·PDF 뷰어에서 보이는 모습, 회사 보안 프로그램의 영향.

  `--moved` 를 붙이면 "설치한 폴더를 다른 드라이브로 옮긴 뒤" 를 점검합니다 (회사 PC 에서 C: 가 차서 D: 로 옮기는 경우).
  가상환경을 지우지 않고 그대로 쓰며, 예전 자리(PM_OLD_ROOT)가 비었는지와 가상환경이 새 자리를 가리키는지를 더 봅니다.

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
MOVED = "--moved" in sys.argv[1:]              # 다른 드라이브로 옮긴 뒤의 점검
VPY = ROOT / ".venv" / ("Scripts/python.exe" if WIN else "bin/python")
MOCK_PORT = 11434 if WIN else 11601          # 배치 파일은 11434 만 본다. Mac 에서는 진짜 Ollama 와 겹치지 않게
WEB_PORT = 8765 if WIN else 8791
TMP = ROOT / "ci_tmp"
TESTS = ["test_filter", "smoke", "test_review", "test_eval", "test_memory", "test_local_input",
         "test_model_call", "test_pdf", "test_prehl", "test_cp949"]
RESULTS: list[tuple[str, bool | None, str]] = []     # 결과가 None 이면 '참고' (통과·실패로 세지 않음)

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


def note(name: str, ok: bool | None, detail: str = "") -> bool:
    RESULTS.append((name, None if ok is None else bool(ok), detail))
    mark = "INFO" if ok is None else ("PASS" if ok else "FAIL")
    print(f"{mark}  {name}" + (f"  ({detail})" if detail else ""), flush=True)
    return bool(ok)


def drop(bat: str, paths: list[Path], vpy: Path) -> list[str] | str:
    """탐색기가 파일을 배치 파일 위에 끌어다 놓을 때 만드는 명령줄. 공백이 든 경로만 따옴표로 감싼다.

    Windows 가 아니면 배치 파일이 부르는 파이썬 명령으로 바꾼다.
    """
    if not WIN:
        return [str(vpy), "-m", "app.cli", *map(str, paths)]
    args = " ".join(f'"{p}"' if " " in str(p) else str(p) for p in paths)
    return f'cmd /c ""{ROOT / bat}" {args}"'


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


def run(cmd: list[str] | str, env: dict, timeout: int, label: str) -> tuple[int, str]:
    print(f"\n{'─' * 20} {label} {'─' * 20}", flush=True)
    if isinstance(cmd, str):
        print(cmd, flush=True)
    try:
        p = subprocess.run(cmd, cwd=str(ROOT), env=env, stdin=subprocess.DEVNULL,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout)
        rc, raw = p.returncode, p.stdout
    except subprocess.TimeoutExpired as e:
        rc, raw = -9, (e.stdout or b"") + "\n[시간 초과]".encode("utf-8")
    out = decode(raw)
    if len(out) > 5000:                          # 잘린 앞부분에 실패 줄이 있으면 그것만은 보여 준다
        for ln in out[:-5000].splitlines():
            if ln.startswith("FAIL"):
                print(ln, flush=True)
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
    if MOVED:
        print(f"옮긴 자리에서 다시 점검합니다: {ROOT} (가상환경은 옮겨 온 것을 그대로 씀)")
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
            f"시스템 글자표 {locale.getpreferredencoding(False)}", f"폴더 {ROOT}" + (" (옮긴 뒤)" if MOVED else "")]
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

        # ── 0. 받은 파일 상태 ────────────────────────────────────────
        named = ["samples/회사보고자료_예시.pptx", "docs/00_Windows_따라하기_가이드.md", "★먼저읽기_Windows_설치순서.txt"]
        lost = [n for n in named if not (ROOT / n).is_file()]
        note("받은 파일: 한글 이름 파일이 그대로 있음", not lost, "없음: " + ", ".join(lost) if lost else f"{len(named)}개 확인")
        bats = sorted(ROOT.glob("*.bat"))
        bare = [b.name for b in bats if b.read_bytes().replace(b"\r\n", b"").count(b"\n")]
        note("받은 파일: 배치 파일 줄바꿈이 모두 CRLF", len(bats) >= 4 and not bare,
             "LF 섞임: " + ", ".join(bare) if bare else ", ".join(b.name for b in bats))
        if MOVED:
            old = os.environ.get("PM_OLD_ROOT", "")
            note("옮긴 뒤: 예전 자리가 비어 있음", bool(old) and not Path(old).exists(), old or "PM_OLD_ROOT 없음")
            if WIN:
                note("옮긴 뒤: 다른 드라이브에서 실행 중", bool(old) and Path(old).drive.lower() != ROOT.drive.lower(),
                     f"{Path(old).drive or '?'} → {ROOT.drive or '?'}")

        # ── 1. setup.bat ────────────────────────────────────────────
        vpy = VPY
        if WIN:
            if not MOVED:
                shutil.rmtree(ROOT / ".venv", ignore_errors=True)
            rc, out = run(["cmd", "/c", "setup.bat"], env, 1800, "setup.bat" + (" (옮긴 뒤 다시 실행)" if MOVED else ""))
            note("setup.bat: 끝까지 실행 (종료 코드 0)", rc == 0, f"종료 코드 {rc}")
            note("setup.bat: 가상환경(.venv) " + ("그대로 씀" if MOVED else "생성"), vpy.exists())
            steps = [k for k in ("[1/5]", "[2/5]", "[3/5]", "[4/5]", "[5/5]", "[5/5-2]", "[5/5-3]", "설치 완료") if k in out]
            note("setup.bat: 다섯 단계와 '설치 완료' 가 한글로 찍힘", len(steps) == 8 and "가상환경" in out, " ".join(steps))
            note("setup.bat: 모델 셋을 '이미 있음' 으로 인식", out.count("이미 있음") >= 3, f"{out.count('이미 있음')}회")
            # 2026-09-28 이전 판에서 올라오는 PC: 가상환경은 있고 PDF 라이브러리 둘만 없다 → setup.bat 재실행으로 채워지는지
            if vpy.exists() and not MOVED:
                run([str(vpy), "-m", "pip", "uninstall", "-y", "pdfminer.six", "pypdf"], env, 300, "업데이트 흉내: PDF 라이브러리 지우기")
                rc, out = run([str(vpy), "-c", "import pdfminer, pypdf"], env, 60, "지워졌는지")
                gone = rc != 0
                t0 = time.time()
                rc, out = run(["cmd", "/c", "setup.bat"], env, 1800, "setup.bat (다시 실행)")
                took = time.time() - t0
                rc2, _ = run([str(vpy), "-c", "import pdfminer, pypdf"], env, 60, "다시 들어왔는지")
                note("setup.bat 재실행: 빠진 PDF 라이브러리 둘을 채움 (업데이트 경로)",
                     gone and rc == 0 and rc2 == 0 and "설치 완료" in out, f"종료 코드 {rc} · {took:.0f}초")
        if not vpy.exists():
            note("준비: 가상환경의 파이썬", False, str(vpy))
            return finish(info, libs)
        rc, out = run([str(vpy), "-c",
                       "import importlib.metadata as m; print(' · '.join(f'{n} {m.version(n)}' for n in "
                       "('python-pptx','lxml','fastapi','uvicorn','python-multipart','pdfminer.six','pypdf','cryptography')))"],
                      {**env, "PYTHONIOENCODING": "utf-8"}, 120, "설치된 라이브러리")
        libs = out.strip().splitlines()[-1] if rc == 0 and out.strip() else ""
        note("라이브러리: PDF 용 둘(pdfminer.six · pypdf)까지 설치됨", "pdfminer.six 20250506" in libs and "pypdf 6.1.3" in libs, libs)
        rc, out = run([str(vpy), "-c", "import sys; print(sys.prefix)"], {**env, "PYTHONIOENCODING": "utf-8"}, 60, "가상환경 위치")
        prefix = out.strip().splitlines()[-1] if rc == 0 and out.strip() else ""
        note("가상환경이 이 폴더를 가리킴" + (" (옮긴 뒤에도)" if MOVED else ""),
             rc == 0 and Path(prefix).resolve() == (ROOT / ".venv").resolve(), prefix)

        # ── 2. 점검 스크립트 (출력 글자표를 한국어 Windows 처럼) ──────────
        # 임시 폴더: 사용자 이름이 한글인 PC 의 C:\Users\홍길동\AppData\Local\Temp 처럼 한글·공백이 든 경로
        ktmp = TMP / "홍길동 임시"
        ktmp.mkdir()
        for e in (env, benv):
            e.update({"TEMP": str(ktmp), "TMP": str(ktmp), "TMPDIR": str(ktmp)})
        benv["PATH"] = env["PATH"]
        tenv = {**env, "PYTHONIOENCODING": "cp949"}
        for t in TESTS:
            rc, out = run([str(vpy), f"tests/{t}.py"], tenv, 900, f"tests/{t}.py")
            n_pass, n_fail = len(re.findall(r"^PASS", out, re.M)), len(re.findall(r"^FAIL", out, re.M))
            last = next((ln.strip() for ln in reversed(out.splitlines()) if ln.strip()), "")
            note(f"점검 {t}", rc == 0 and n_fail == 0, f"통과 {n_pass} · 실패 {n_fail} · {last[:40]}")

        # ── 3. mark.bat (끌어다 놓기) ─────────────────────────────────
        work = TMP / "끌어다 놓기 시험 (공백 있음)"
        work.mkdir(parents=True)
        pptx, pdf = work / "보고서 초안.pptx", work / "변환본.pdf"
        shutil.copyfile(ROOT / "samples" / "회사보고자료_예시.pptx", pptx)
        rc, out = run([str(vpy), "tools/make_sample_pdf.py", str(pptx), str(pdf)], env, 120, "점검용 PDF 만들기")
        note("준비: 점검용 PDF 생성", rc == 0 and pdf.exists())
        rc, out = run(drop("mark.bat", [pptx, pdf], vpy), benv, 900, "mark.bat (파일 두 개)")
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

        # 이름에 괄호·&·% 가 든 파일 (회사 문서에 흔한 이름). 공백이 없으면 탐색기는 따옴표 없이 넘긴다
        names = TMP / "이름시험"
        names.mkdir()
        paren, amp_sp, amp = names / "보고서(최종).pptx", names / "R&D 현황 50%.pptx", names / "R&D현황.pptx"
        for f in (paren, amp_sp, amp):
            shutil.copyfile(pptx, f)
        rc, out = run(drop("mark.bat", [paren], vpy), benv, 900, "mark.bat 보고서(최종).pptx — 따옴표 없이")
        note("mark.bat: 괄호가 든 이름 (따옴표 없이 넘어옴)", (names / "보고서(최종)_특허마킹.pptx").exists())
        rc, out = run(drop("mark.bat", [amp_sp], vpy), benv, 900, "mark.bat \"R&D 현황 50%.pptx\"")
        note("mark.bat: & 와 % 와 공백이 든 이름", (names / "R&D 현황 50%_특허마킹.pptx").exists())
        # & 가 있고 공백이 없는 이름: cmd 가 & 앞에서 이름을 자르고 뒷부분을 명령으로 실행하려 든다.
        # mark.bat 이 원래 명령줄을 넘겨 주어 cli 가 이름을 되살리고, 끝에서 창을 닫아 뒷부분이 실행되지 않게 한다.
        cut_cmd, aenv = drop("mark.bat", [amp], vpy), benv
        if not WIN:                                  # Mac: cmd 가 자른 모습과 원래 명령줄을 흉내 낸다
            cut_cmd = [str(vpy), "-m", "app.cli", str(names / "R")]
            aenv = {**benv, "PM_CMDLINE": f'cmd.exe /c ""C:\\patent_marker\\mark.bat" {amp}"'}
        rc, out = run(cut_cmd, aenv, 900, "mark.bat R&D현황.pptx — 따옴표 없이")
        stray = "recognized" in out or "배치 파일이 아닙니다" in out
        note("mark.bat: & 가 있고 공백이 없는 이름 — 잘린 이름을 되살려 처리",
             (names / "R&D현황_특허마킹.pptx").exists() and "되살렸습니다" in out and not stray,
             "잘린 뒷부분이 명령으로 실행됨" if stray else "")
        if WIN:
            rc, out = run("cmd /c mark.bat", benv, 120, "mark.bat (파일 없이 더블클릭)")
            note("mark.bat: 파일 없이 실행하면 사용법을 알리고 끝남", rc == 1 and "끌어다" in out, f"종료 코드 {rc}")

        # ── 4. run.bat (웹 서버) ─────────────────────────────────────
        for old_result in out_dir.glob("*특허마킹*"):    # 앞선 실행의 결과가 남아 있으면 지우고 새로 확인
            old_result.unlink()
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
            st, _, raw = get(base + "/api/local/jobs", 60, {"Content-Type": "application/json", "X-PM-Token": token},
                             json.dumps({"path": f'"{amp}"', "model": "qwen3:8b"}).encode("utf-8"))
            snap = wait_job(base, json.loads(raw)["job_id"])
            note("run.bat: 경로 입력은 & 가 든 이름도 됨 (따옴표째 붙여넣기)", snap.get("status") == "done"
                 and (out_dir / "R&D현황_특허마킹.pptx").exists(), str(snap.get("error") or snap.get("stage")))
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
    bad = [r for r in RESULTS if r[1] is False]
    counted = [r for r in RESULTS if r[1] is not None]
    lines = [f"## Windows 점검 — {'모두 통과' if not bad else f'실패 {len(bad)}건'} ({len(counted) - len(bad)}/{len(counted)})",
             "", " · ".join(info), "", f"설치된 라이브러리: {libs}" if libs else "", "",
             "| 결과 | 항목 | 내용 |", "|---|---|---|"]
    for name, ok, detail in RESULTS:
        lines.append(f"| {'참고' if ok is None else '통과' if ok else '**실패**'} | {name} | {detail.replace('|', '/')} |")
    text = "\n".join(lines) + "\n"
    print("\n" + text, flush=True)
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if path:
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(text)
    return 1 if bad or not counted else 0


if __name__ == "__main__":
    sys.exit(main())
