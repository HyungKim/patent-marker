"""
tests/test_mode.py ─ 판정 기준 '확실한 것만'(strict) / '빠짐없이'(broad) 와 슬라이드 묶음 호출 점검 (가짜 모델, 30초)
=====================================================================
2026-10-01 (docs/05 23회차). 회사 PC 에서 "너무 느리고, 애매한 것까지 칠한다" 는 사용자 판단에 따라
기본 판정 기준을 '확실한 것만' 으로 바꾸고, 지시서를 반으로 줄이고, 슬라이드 몇 장을 한 호출로 묶었다.

    실행:  .venv/bin/python tests/test_mode.py     (Windows: .venv\\Scripts\\python tests\\test_mode.py)

확인하는 것
  1. 설정: 기본 판정 기준, 기준별 '전체 문단 검사' 기본값, 지시서 선택, 지시서 길이
  2. 규칙 사전: strict 규칙 집합, 부정 표현 뒤집기, 분수 표기
  3. 1차 스캔: strict 는 명시형·공개 규칙만 보고 짧은 문단을 보내지 않는다
  4. 병합: strict 는 자리 못 찾은 후보·수단 없는 B·표 조각을 버리고, 확실한 공개 신호만 살린다 / 너그러운 자리 찾기
  5. 묶음 호출: 글자 예산대로 묶고, 프롬프트에 슬라이드 머리글이 들어간다
  6. 파이프라인(가짜 모델): strict 가 broad 보다 후보·출력 토큰이 적고 문단 전체 칠하기가 없다, 실행 기록에 판정 기준 칸
  7. 명령행·서버·화면: --mode, /api/health 의 선택지, 화면의 판정 기준 선택
"""
from __future__ import annotations

import copy
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

TMP = Path(tempfile.mkdtemp(prefix="pm-mode-test-"))
os.environ["PM_REVIEW_DIR"] = str(TMP / "review_data")
os.environ["PM_OUTPUT_DIR"] = str(TMP / "output")
os.environ["PM_INPUT_DIR"] = str(TMP / "input")

import mock_llm  # noqa: E402

MOCK_HOST = mock_llm.start()
os.environ["PM_OLLAMA_HOST"] = MOCK_HOST

from app import analyze, cli, config, extract, lexicon, merge, pipeline  # noqa: E402
from app.extract import Segment  # noqa: E402

config.OLLAMA_HOST = MOCK_HOST
analyze.config.OLLAMA_HOST = MOCK_HOST
config.REVIEW_DIR = TMP / "review_data"
config.OUTPUT_DIR = TMP / "output"
config.INPUT_DIR = TMP / "input"

SAMPLE = ROOT / "samples" / "회사보고자료_예시.pptx"
FAILS = 0


def check(label: str, ok: bool, detail: str = "") -> None:
    global FAILS
    print(f"{'PASS' if ok else 'FAIL'}  {label}" + (f"  ({detail})" if detail else ""))
    if not ok:
        FAILS += 1


# ── 1. 설정 ────────────────────────────────────────────────────────
check("기본 판정 기준은 '확실한 것만'", config.MODE == "strict" and config.RunOptions().mode == "strict")
check("판정 기준별 전체 문단 검사 기본값: strict 끔 · broad 켬",
      config.RunOptions().scan_all_paragraphs is False and config.RunOptions(mode="broad").scan_all_paragraphs is True
      and config.RunOptions(mode="strict", scan_all_paragraphs=True).scan_all_paragraphs is True)
check("모르는 값은 strict 로", config.RunOptions(mode="???").mode == "strict")
check("지시서 선택: strict / broad", analyze.system_for("strict") is analyze.SYSTEM_STRICT
      and analyze.system_for("broad") is analyze.SYSTEM_BROAD and analyze.SYSTEM_BROAD is analyze.SYSTEM)
check("strict 지시서는 broad 의 절반 이하 길이", len(analyze.SYSTEM_STRICT) * 2 <= len(analyze.SYSTEM),
      f"{len(analyze.SYSTEM_STRICT)} vs {len(analyze.SYSTEM)} 글자")
check("strict 지시서에 '확실한 것만' · 공백 없는 JSON 예시", "확실한 것만" in analyze.SYSTEM_STRICT
      and '{"findings":[{"i":3' in analyze.SYSTEM_STRICT)
check("_system_prompt(mode) 가 기준별 지시서로 시작",
      analyze._system_prompt("strict").startswith(analyze.SYSTEM_STRICT) and analyze._system_prompt("broad").startswith(analyze.SYSTEM))
check("선택지 둘 (확실한 것만 · 빠짐없이)", [c["id"] for c in config.MODE_CHOICES] == ["strict", "broad"])

# ── 2. 규칙 사전 ───────────────────────────────────────────────────
check("strict 규칙에 묵시형(IMP_*)이 없고 명시형·공개·사용자 규칙만",
      not any(r.startswith("IMP_") and r != "IMP_AUTO" for r in lexicon.STRICT_RULE_IDS)
      and {"NUM_RANGE", "RISK_PAPER", "CTRL_ALGO"} <= lexicon.STRICT_RULE_IDS
      and lexicon.is_strict_rule("USER_3") and not lexicon.is_strict_rule("IMP_OWN"))
hits = lexicon.scan("캐스케이드 구조와 적응형 임계 방식은 사내 독자 설계이며 현재 논문·특허 등 공개 실적 없음.")
paper = [h for h in hits if h.rid == "RISK_PAPER"]
check("'논문 … 공개 실적 없음' 은 뒤집힌 공개 신호", bool(paper) and all(lexicon.negated(
    "캐스케이드 구조와 적응형 임계 방식은 사내 독자 설계이며 현재 논문·특허 등 공개 실적 없음.", h.span) for h in paper))
hits2 = lexicon.scan("2026. 09 한국전지학회 추계학술대회 논문 발표 예정")
check("'논문 발표 예정' 은 그대로 공개 신호",
      any(h.rid == "RISK_PAPER" and not lexicon.negated("2026. 09 한국전지학회 추계학술대회 논문 발표 예정", h.span) for h in hits2))
check("분수 표기 1/33 도 수치 신호 (날짜 8/31 도 걸리지만 가중치는 낮음)",
      any(h.rid == "NUM_FRACTION" and h.matched == "1/33" and h.weight == 3 for h in lexicon.scan("2차 모델의 입력량을 1/33로 축소")))
check("strict_only 가 묵시형 신호를 걸러 낸다",
      [h.rid for h in lexicon.strict_only(lexicon.scan("자체 개발한 22~28° 각도 가변 마운트로 원가 40% 절감"))] == ["NUM_RANGE"])

# ── 3. 1차 스캔 ────────────────────────────────────────────────────
deck = extract.extract(str(SAMPLE))
h_s, t_s = analyze.prescreen(deck, config.RunOptions(mode="strict"))
h_b, t_b = analyze.prescreen(deck, config.RunOptions(mode="broad"))
check("strict 는 broad 보다 보내는 문단이 훨씬 적다", 0 < len(t_s) < len(t_b) / 2, f"{len(t_s)} vs {len(t_b)}")
check("strict 가 보내는 문단은 모두 STRICT_MIN_CHARS 이상이고 strict 규칙에 걸린 것",
      all(len(s.text) >= config.STRICT_MIN_CHARS and h_s.get(s.seg_id) for s in t_s))
check("strict 의 신호 목록에는 strict 규칙만", all(lexicon.is_strict_rule(h.rid) for hs in h_s.values() for h in hs))
check("strict + 전체 문단 검사 를 켜면 짧은 문단만 빼고 다 보낸다",
      len(analyze.prescreen(deck, config.RunOptions(mode="strict", scan_all_paragraphs=True))[1])
      == sum(1 for s in t_b if len(s.text) >= config.STRICT_MIN_CHARS))

# ── 4. 병합 ────────────────────────────────────────────────────────
segs = [Segment(1, 1, "body", "a", "판정 임계값을 직전 500프레임의 결함 밀도에 따라 자동 조정하는 적응형 임계 방식 적용"),
        Segment(2, 1, "body", "b", "캐스케이드 구조와 적응형 임계 방식은 사내 독자 설계이며 현재 논문·특허 등 공개 실적 없음."),
        Segment(3, 1, "body", "c", "38 ms → 11 ms"),
        Segment(4, 1, "body", "d", "2026. 09 한국전지학회 추계학술대회 논문 발표 예정 (초록 제출 마감 8/31)"),
        Segment(5, 1, "body", "e", "각도 가변 마운트를 자체 제작해 적용했고 결과가 좋았다."),
        Segment(6, 1, "body", "f", "검사 모듈 12조 설치 및 SOP 승인 (양산 개시)")]
mini = SimpleNamespace(segments=segs)
hits_all = {s.seg_id: lexicon.scan(s.text) for s in segs}
F = analyze.Finding
raw = [F(1, 1, "직전 500 프레임의 결함밀도에 따라 자동조정", "A", "제어·알고리즘", False, False, "x"),  # 공백·조사가 조금 다른 인용구
       F(3, 1, "38 ms → 11 ms", "A", "수치·범위한정", False, False, "x"),                          # 표 조각
       F(5, 1, "자체 제작", "B", "독자성주장", True, False, "x"),                                    # 수단 없는 B
       F(1, 1, "이 문단에 없는 인용구입니다", "A", "구성·구조", False, False, "x")]                 # 자리 못 찾음
res_s, marks_s = merge.resolve(mini, copy.deepcopy(raw), copy.deepcopy(hits_all), mode="strict")
res_b, marks_b = merge.resolve(mini, copy.deepcopy(raw), copy.deepcopy(hits_all), mode="broad")
keys_s = [(f.seg_id, f.grade, f.source, f.disclosure_risk) for f in res_s]
check("strict: 변형된 인용구도 자리를 찾아 A 로 남긴다",
      any(f.seg_id == 1 and f.span == (8, 34) for f in res_s), str([(f.seg_id, f.span) for f in res_s]))
check("strict: 자리를 못 찾은 후보는 버린다 (문단 전체 칠하기 없음)", all(f.span is not None for f in res_s))
check("strict: 수단 없는 B 와 표 조각은 버린다", not any(f.seg_id in (3, 5) for f in res_s), str(keys_s))
check("strict: '논문 발표 예정' 은 사전으로 살리고, '공개 실적 없음' 은 살리지 않는다",
      any(f.seg_id == 4 and f.source == "lexicon" and f.disclosure_risk for f in res_s)
      and not any(f.seg_id == 2 for f in res_s), str(keys_s))
check("strict: 출시·양산(가중치 6)은 사전만으로 살리지 않는다", not any(f.seg_id == 6 for f in res_s))
check("broad: 예전처럼 문단 전체 칠하기·점수 안전망·B 가 남는다",
      any(f.span is None for f in res_b) and any(f.seg_id == 2 and f.source == "lexicon" for f in res_b)
      and any(f.seg_id == 5 and f.grade == "B" for f in res_b) and any(f.seg_id == 6 for f in res_b),
      str([(f.seg_id, f.grade, f.source, f.span) for f in res_b]))
check("broad: 뒤집힌 공개 신호는 공개로 세지 않는다", not any(f.seg_id == 2 and f.disclosure_risk for f in res_b))
check("기본 인수(mode 없음)는 broad 와 같다",
      [(f.seg_id, f.grade, f.source) for f in merge.resolve(mini, copy.deepcopy(raw), copy.deepcopy(hits_all))[0]]
      == [(f.seg_id, f.grade, f.source) for f in res_b])

# ── 5. 묶음 호출 ───────────────────────────────────────────────────
by: dict[int, list] = {}
for s in t_b:
    by.setdefault(s.slide_no, []).append(s)
check("묶음 0 = 슬라이드마다 한 번", pipeline.batches_of(by, deck.slide_count, 0) == [[1], [2], [3], [4], [5]])
big = pipeline.batches_of(by, deck.slide_count, 10 ** 6)
check("예산이 크면 전부 한 묶음", big == [[1, 2, 3, 4, 5]], str(big))
mid = pipeline.batches_of(by, deck.slide_count, 2400)
sizes = [sum(len(s.text) + 16 for n in g for s in by[n]) for g in mid]
check("예산 2400: 묶음마다 예산 이하이고 순서대로 빠짐없이", all(x <= 2400 for x in sizes)
      and [n for g in mid for n in g] == [1, 2, 3, 4, 5] and 1 < len(mid) < 5, f"{mid} {sizes}")
check("문단 없는 장은 묶음에 넣지 않는다", pipeline.batches_of({2: by[2], 5: by[5]}, 6, 2400) == [[2, 5]])
prompt = analyze._build_user_prompt("예시", [2, 3], 5, by[2] + by[3], {})
check("여러 장 프롬프트: 머리글 '슬라이드 2~3' 과 문단 줄마다 '· 슬라이드 n' (따로 된 머리글은 없음 — 모델이 i 로 착각)",
      "슬라이드 2~3 / 전체 5" in prompt and "[본문 · 슬라이드 2]" in prompt and "[본문 · 슬라이드 3]" in prompt
      and "## 슬라이드" not in prompt)
check("한 장 프롬프트는 예전 그대로 (문단 줄에 슬라이드 표기 없음)",
      "슬라이드 2 / 전체 5" in analyze._build_user_prompt("예시", 2, 5, by[2], {})
      and "· 슬라이드" not in analyze._build_user_prompt("예시", 2, 5, by[2], {}))
check("묶음 호출 기본은 끔 (장마다 한 번) — 묶으면 후보를 놓친다 (2026-10-01 측정)", config.BATCH_CHARS == 0)
check("slide_label", (analyze.slide_label(3), analyze.slide_label([3, 4, 5]), analyze.slide_label([3, 5])) == ("3", "3~5", "3, 5"))

# ── 6. 파이프라인 (가짜 모델) ──────────────────────────────────────
stages_s: list[str] = []
res_s, st_s = pipeline.run(SAMPLE, TMP / "strict.pptx", config.RunOptions(mode="strict"),
                           progress=lambda p: stages_s.append(p.stage))
res_b, st_b = pipeline.run(SAMPLE, TMP / "broad.pptx", config.RunOptions(mode="broad"))
check("strict 는 broad 보다 후보가 적다", 0 < st_s["total"] < st_b["total"], f"{st_s['total']} vs {st_b['total']}")
check("strict 는 broad 보다 모델이 읽은 토큰이 적다 (지시서 절반 + 문단 일부)",
      st_s["prompt_tokens"] < st_b["prompt_tokens"], f"{st_s['prompt_tokens']} vs {st_b['prompt_tokens']}")
check("strict 결과에 문단 전체 칠하기가 없다", all(f.span is not None for f in res_s))
check("strict 결과의 모델 후보는 A 이거나 공개 관련", all(f.grade == "A" or f.disclosure_risk for f in res_s if f.source == "llm"),
      str([(f.grade, f.disclosure_risk, f.category) for f in res_s if f.source == "llm"]))
check("strict 결과에 STRICT_MIN_CHARS 미만 문단의 모델 후보가 없다",
      all(len(next(s.text for s in deck.segments if s.seg_id == f.seg_id)) >= config.STRICT_MIN_CHARS
          for f in res_s if f.source == "llm"))
check("진행 표시: 장마다 '슬라이드 n 분석 완료 · 시간 · 토큰'", any(s.startswith("슬라이드 3 분석 완료 ·") for s in stages_s), str(stages_s[:6]))
_saved = config.BATCH_CHARS
config.BATCH_CHARS = 2400
stages_m: list[str] = []
res_m, st_m = pipeline.run(SAMPLE, TMP / "strict_batched.pptx", config.RunOptions(mode="strict"),
                           progress=lambda p: stages_m.append(p.stage))
config.BATCH_CHARS = _saved
check("묶음 호출을 켜면(PM_BATCH_CHARS) 호출 수가 줄고 진행 표시에 '슬라이드 1~' 범위가 나온다",
      st_m["calls"] < st_s["calls"] and any("슬라이드 1~" in s for s in stages_m), f"{st_m['calls']} vs {st_s['calls']}")
check("집계에 판정 기준", st_s["mode"] == "strict" and st_s["mode_text"] == "확실한 것만" and st_b["mode_text"] == "빠짐없이")
log = (config.REVIEW_DIR / "run_log.tsv").read_text(encoding="utf-8-sig").splitlines()
check("실행 기록에 '판정기준' 칸과 값", "판정기준" in log[0].split("\t") and "확실한 것만" in log[1] and "빠짐없이" in log[2],
      log[0][:60])
check("결과 파일 둘 다 저장", (TMP / "strict.pptx").exists() and (TMP / "broad.pptx").exists())

# ── 7. 명령행 · 서버 · 화면 ──────────────────────────────────────
check("cli --mode 기본 strict, --mode broad 선택", cli._parse(["x.pptx"]).mode == "strict" and cli._parse(["x.pptx", "--mode", "broad"]).mode == "broad")
check("cli 전체 문단 검사: 기본 None(기준을 따름) / --scan-all / --no-scan-all",
      cli._parse(["x.pptx"]).scan_all is None and cli._parse(["x.pptx", "--scan-all"]).scan_all is True
      and cli._parse(["x.pptx", "--no-scan-all"]).scan_all is False)
from app import main as web  # noqa: E402

h = json.loads(web.health().body)
check("/api/health 에 판정 기준 기본값·선택지·기준별 전체 문단 기본값",
      h.get("mode_default") == "strict" and [c["id"] for c in h.get("mode_choices", [])] == ["strict", "broad"]
      and h.get("scan_all_default") == {"strict": False, "broad": True})
html = web.index().body.decode("utf-8")
check("화면에 판정 기준 선택 상자와 요청에 mode 전달", 'id="optMode"' in html and "chosenMode" in html and "fd.append(\"mode\"" in html)
o = web._options("qwen3:8b", False, None, False, mode="broad")
check("서버 옵션: mode 와 scan_all(None → 기준 기본값)", o.mode == "broad" and o.scan_all_paragraphs is True
      and web._options("qwen3:8b", False, None, False).mode == "strict"
      and web._options("qwen3:8b", False, None, False).scan_all_paragraphs is False)

# 별도 프로세스: mark.bat 이 부르는 방식 그대로, 기본(strict)과 --mode broad
env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PM_OLLAMA_HOST": MOCK_HOST,
       "PM_OUTPUT_DIR": str(TMP / "output"), "PM_REVIEW_DIR": str(TMP / "review_data")}
proc = subprocess.run([sys.executable, "-m", "app.cli", str(SAMPLE), "--out", str(TMP / "cli_strict")],
                      cwd=str(ROOT), env=env, capture_output=True, text=True, encoding="utf-8")
check("명령행 기본: '판정 기준 확실한 것만' 을 찍고 결과를 만든다",
      proc.returncode == 0 and "판정 기준 확실한 것만" in proc.stdout and (TMP / "cli_strict" / "회사보고자료_예시_특허마킹.pptx").exists(),
      (proc.stdout + proc.stderr).strip().splitlines()[-1] if (proc.stdout + proc.stderr).strip() else "")
proc = subprocess.run([sys.executable, "-m", "app.cli", str(SAMPLE), "--out", str(TMP / "cli_broad"), "--mode", "broad"],
                      cwd=str(ROOT), env=env, capture_output=True, text=True, encoding="utf-8")
check("명령행 --mode broad: '판정 기준 빠짐없이'", proc.returncode == 0 and "판정 기준 빠짐없이" in proc.stdout)

print("\n" + ("모두 통과" if not FAILS else f"실패 {FAILS}건") + f"  (임시 폴더: {TMP})")
sys.exit(1 if FAILS else 0)
