"""
cli.py ─ 브라우저 없이 명령행(또는 mark.bat 끌어다 놓기)으로 마킹하기
=====================================================================

    python -m app.cli 보고서.pptx                    → 보고서_특허마킹.pptx (원본 옆)
    python -m app.cli 보고서.pdf                     → 보고서_특허마킹.pdf  (PPT 를 PDF 로 변환한 자료도 그대로)
    python -m app.cli a.pptx b.pptx --out C:\\결과    → 지정한 폴더에 저장
    python -m app.cli 보고서.pptx --tag              → 【출원검토필요】 문구도 표시 (흑백 인쇄용)
    python -m app.cli 보고서.pptx --fast             → 빠름 모델(qwen3:4b-instruct)로 초벌 — 약 3배 빠름, 후보는 덜 잡힘
    python -m app.cli 보고서.pptx --keep-highlights  → 원본에 있던 형광펜을 지우지 않고 그대로 둠 (기본은 지우고 시작)
    python -m app.cli 보고서.pptx --mode broad       → 판정 기준 '빠짐없이' (기본은 '확실한 것만')
    python -m app.cli 보고서.pptx --no-disclosure   → 공개 관련 표시 끔 (전시·논문·출시 신호를 보지 않음, 조금 빠름)

  Windows 에서는 mark.bat 아이콘 위에 PPTX·PDF 를 끌어다 놓으면 이 파일이 실행됩니다.
  Ollama 가 떠 있어야 합니다 (run.bat / mark.bat 이 자동으로 켭니다).

[결과 저장 위치]
  기본은 원본 파일 옆입니다. 그 폴더에 쓸 수 없으면(읽기 전용·네트워크 드라이브 등)
  output 폴더로 저장하고 그렇게 알려 줍니다. 같은 이름이 있으면 (2), (3) 을 붙입니다.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path

from . import analyze, config, local, pipeline


def _parse(argv: list[str]) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        prog="python -m app.cli",
        description="PPTX·PDF 를 읽어 특허 출원 검토가 필요한 구간을 형광펜으로 표시합니다.")
    ap.add_argument("files", nargs="+", help="분석할 PPTX 또는 PDF 파일 (여러 개 가능)")
    ap.add_argument("--out", metavar="폴더", help="결과를 저장할 폴더 (기본: 원본 옆)")
    ap.add_argument("--tag", action="store_true", help="형광펜 뒤에 【출원검토필요】 문구도 표시")
    ap.add_argument("--think", action="store_true", help="추론 모드 (정확도↑ 속도↓)")
    ap.add_argument("--mode", choices=("strict", "broad"), default=config.MODE,
                    help="판정 기준: strict = 확실한 것만(기본, 빠름) / broad = 빠짐없이 (2026-09-30b 까지의 동작)")
    ap.add_argument("--scan-all", dest="scan_all", action="store_const", const=True, default=None,
                    help="규칙 사전에 걸리지 않은 문단도 모델에 보냄 (확실한 것만 에서는 기본 끔)")
    ap.add_argument("--no-scan-all", dest="scan_all", action="store_const", const=False,
                    help="규칙 사전에 걸린 문단만 모델에 보냄 (빠름 · 빠짐없이 에서는 기본 켬)")
    ap.add_argument("--model", default=config.MODEL, help=f"Ollama 모델 이름 (기본 {config.MODEL})")
    ap.add_argument("--fast", action="store_true",
                    help=f"빠름 모델({config.FAST_MODEL}) 사용 — 약 3배 빠르지만 후보를 덜 잡음 (초벌용)")
    ap.add_argument("--no-learn", action="store_true", help="검토 학습(문단 기억·제외 사전·유사 사례) 끄기")
    ap.add_argument("--no-embed", action="store_true", help=f"유사 사례를 뜻 기준({config.EMBED_MODEL})으로 찾지 않고 글자 겹침만")
    ap.add_argument("--keep-highlights", action="store_true",
                    help="원본에 있던 형광펜을 지우지 않고 그대로 둠 (기본은 지우고 시작 — 검토 반영 때 섞이지 않게)")
    ap.add_argument("--no-disclosure", action="store_true",
                    help="공개 관련 표시 끔 — 전시·논문·출시 같은 공개 신호를 보지 않고 공개 후보를 내지 않음 (조금 빠름)")
    args = ap.parse_args(argv)
    if args.fast:
        args.model = config.FAST_MODEL
    return args


def recover_dropped(files: list[str], cmdline: str) -> list[str]:
    """mark.bat 에 끌어다 놓은 파일 이름이 Windows 명령줄 규칙 때문에 잘렸으면 원래 이름을 되살린다.

    탐색기는 공백이 든 경로만 따옴표로 감싸 넘긴다. 그래서 C:\\문서\\R&D현황.pptx 처럼 & 가 있고 공백이 없는
    이름은 cmd 가 & 앞(C:\\문서\\R)까지만 넘기고 뒷부분은 다른 명령으로 실행하려 든다 (^ 도 사라진다).
    mark.bat 이 잘리기 전의 명령줄 전체(%cmdcmdline%)를 PM_CMDLINE 으로 넘겨 주므로 거기서 경로를 다시 읽는다.
    넘어온 파일이 전부 있으면 아무것도 하지 않는다 (2026-09-30 Windows 점검에서 확인한 문제).
    """
    n_ok = sum(Path(local.clean_path(f)).is_file() for f in files)
    if not cmdline or n_ok == len(files):
        return files
    m = re.search(r'mark\.bat"?\s+(.*)$', cmdline, re.I | re.S)
    if not m:
        return files
    rest = m.group(1).strip()
    if rest.endswith('"') and rest.count('"') % 2 == 1:      # cmd /c ""...mark.bat" 파일들" 의 바깥 따옴표
        rest = rest[:-1]
    found = [a or b for a, b in re.findall(r'"([^"]*)"|(\S+)', rest)]
    found = [f for f in found if Path(f).is_file()]
    return found if len(found) > n_ok else files


def _dest_for(src: Path, out_dir: Path | None) -> tuple[Path, str]:
    """결과 경로와, 원본 옆에 못 써서 output 폴더로 바꿨을 때의 안내문."""
    if out_dir is not None:
        return local.output_path_for(src.name, out_dir), ""
    if src.parent == config.INPUT_DIR.resolve():
        # input 폴더의 파일은 결과를 옆에 두면 '분석할 파일' 목록에 결과가 섞여 보인다 → output 폴더로
        return local.output_path_for(src.name), f"  (input 폴더의 파일이라 {config.OUTPUT_DIR} 에 저장합니다)"
    try:
        probe = src.parent / f".pm-write-test-{src.stem}"
        probe.touch()
        probe.unlink()
        return local.output_path_for(src.name, src.parent), ""
    except OSError:
        note = f"  (원본 폴더에 쓸 수 없어 {config.OUTPUT_DIR} 에 저장합니다)"
        return local.output_path_for(src.name), note


def run_one(src: Path, out_dir: Path | None, opts: config.RunOptions) -> Path:
    """파일 하나를 처리하고 결과 경로를 돌려준다. 진행 상황은 한 줄씩 출력."""
    dst, note = _dest_for(src, out_dir)
    print(f"\n▶ {src.name}")
    if note:
        print(note)
    t0 = time.time()
    last = {"slide": 0, "beat": time.time()}

    def progress(p: pipeline.Progress) -> None:
        if p.slide_total and p.slide_done != last["slide"]:
            last["slide"] = p.slide_done
            last["beat"] = time.time()
            # "슬라이드 3 분석 완료 · 2분 10초 · 입력 2,187 / 출력 1,870 토큰 · 쓰기 9.3 토큰/초" 같은 성적표 줄
            print(f"  {p.stage} · 후보 {len(p.findings or [])}건 ({p.slide_done}/{p.slide_total})")
        elif "경과" in p.stage and time.time() - last["beat"] >= 60:
            last["beat"] = time.time()          # 시간 제한이 없으므로 1분마다 살아 있음을 보여 준다
            print(f"    … {p.stage}")

    resolved, stats = pipeline.run(src, dst, opts, progress=progress)
    g = stats.get("grades", {})
    print(f"  완료 — 후보 {stats.get('total', 0)}건 "
          f"(A {g.get('A', 0)} · B {g.get('B', 0)} · C {g.get('C', 0)}) · {time.time() - t0:.0f}초")
    if stats.get("speed_text"):
        print(f"  속도: {stats['speed_text']} · 모델 호출 {stats.get('calls', 0)}회 · "
              f"인용구 일치 {stats.get('quote_located', '')}  (review_data\\run_log.tsv 에 기록됨)")
        print(f"  검토 학습: {stats.get('learn_text', '')}")
    if stats.get("prehl_text"):
        print(f"  {stats['prehl_text']}")
    print(f"  저장: {dst}")
    return dst


def main(argv: list[str] | None = None) -> int:
    # 출력이 파일로 돌려질 때(mark.bat > log.txt) cp949 로 못 적는 글자(✗ ▶)가 있어도 멈추지 않게
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except Exception:  # noqa: BLE001
                pass
    args = _parse(sys.argv[1:] if argv is None else argv)
    local.ensure_dirs()

    out_dir: Path | None = None
    if args.out:
        out_dir = Path(local.clean_path(args.out)).expanduser().resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

    opts = config.RunOptions(model=args.model, think=args.think, mode=args.mode,
                             scan_all_paragraphs=args.scan_all, tag_marks=args.tag,
                             learn=config.LEARN and not args.no_learn,
                             learn_embed=config.LEARN_EMBED and not args.no_embed,
                             strip_highlights=config.STRIP_HIGHLIGHTS and not args.keep_highlights,
                             disclosure=config.DISCLOSURE and not args.no_disclosure)

    files = recover_dropped(args.files, os.environ.get("PM_CMDLINE", ""))
    if files != args.files:
        print("  (파일 이름의 & 때문에 Windows 가 자른 경로를 되살렸습니다)")

    # 먼저 파일들을 전부 확인한다 — 하나가 틀렸다고 나머지까지 못 돌리지 않게, 틀린 것만 알려 준다
    srcs: list[Path] = []
    bad = 0
    for raw in files:
        try:
            srcs.append(local.resolve_pptx(raw))
        except ValueError as e:
            print(f"✗ {raw}: {e}")
            bad += 1
    if not srcs:
        return 1

    h = analyze.health()
    if not h.get("ok"):
        print(f"\n✗ Ollama 에 연결할 수 없습니다 ({config.OLLAMA_HOST}). "
              "run.bat 을 먼저 실행하거나 `ollama serve` 를 켠 뒤 다시 시도하세요.")
        return 1
    if not analyze.model_available(opts.model, h.get("models", [])):
        print(f"\n✗ 모델 {opts.model} 이 없습니다. `ollama pull {opts.model}` 로 먼저 내려받으세요.")
        return 1

    print(f"온디바이스 모델 {opts.model} · 판정 기준 {pipeline.MODE_LABEL.get(opts.mode, opts.mode)}{'' if opts.disclosure else ' · 공개 표시 끔'} · "
          f"파일 {len(srcs)}개 · 도구 버전 {config.VERSION}")
    for src in srcs:
        try:
            run_one(src, out_dir, opts)
        except analyze.OllamaError as e:
            print(f"  ✗ 온디바이스 모델 연결 실패: {e}")
            bad += 1
        except Exception as e:  # noqa: BLE001  (한 파일의 오류가 다음 파일을 막지 않도록)
            print(f"  ✗ 실패: {type(e).__name__}: {e}")
            bad += 1
    print()
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
