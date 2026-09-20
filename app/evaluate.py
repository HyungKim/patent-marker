"""
evaluate.py ─ 성능 측정 (최초 설정 vs 현재 설정, 같은 문제지로 채점)
=====================================================================

[이 파일이 하는 일]
  [검토 반영] 탭으로 쌓인 "사람이 확정한 정답" (review_data/dataset.jsonl) 을
  문제지 삼아, 도구의 판정 성능을 숫자로 잽니다. 회사 보고에 쓸 수 있도록
  "최초 설정" 과 "현재 설정" 을 **같은 문제로 동시에 채점** 해 비교합니다.

      문제지(확정 정답) ──▶ 최초 설정으로 채점 ──┐
                        └─▶ 현재 설정으로 채점 ──┴─▶ 비교 + 이력 누적
                                                    (review_data/eval_history.jsonl)

  - 최초 설정 : 기본 규칙 사전 + 기본 프롬프트 + 기본 모델(qwen3:8b).
                도구를 처음 설치했을 때의 상태. 매 회차 다시 채점합니다.
  - 현재 설정 : [사전에 추가] 표현 + 프롬프트 주입 예시 + 지금 쓰는 모델.
  - 문제지가 회차마다 자라기 때문에, 회차끼리 숫자를 직접 비교하기보다
    **같은 회차 안의 "최초 vs 현재" 차이(Δ)** 를 개선 폭으로 읽는 것이 정확합니다.

[측정 지표 — 쉬운 말로]
  - 재현율(recall)    : 잡아야 할 것 중 실제로 잡은 비율. 높을수록 "누락이 없다".
  - 정밀도(precision) : 잡은 것 중 진짜였던 비율.       높을수록 "헛짚지 않는다".
  - F1                : 위 둘의 조화 평균. 종합 점수 하나가 필요할 때 사용.
  - 등급 일치율       : 맞게 잡은 것 중 등급(A/B/C)까지 맞힌 비율.

[시간에 대하여]
  문단 묶음마다 모델을 두 번(최초·현재) 부르므로 문제지가 크면 몇 분 걸립니다.
  진행률은 웹 화면 [성능 기록] 탭에 표시됩니다.
"""
from __future__ import annotations

import datetime
import json
import threading
import time
from types import SimpleNamespace

from . import analyze, config, lexicon, merge, review
from .extract import Segment

# "최초 설정" 이 쓰는 모델. 도구의 출하 기본값과 같게 유지합니다.
# (이 모델이 PC 에 없으면 현재 모델로 대신 채점하고, 이력에 그 사실을 남깁니다)
BASELINE_MODEL = "qwen3:8b"

# 이력 항목에 남기는 인용구 예시의 최대 개수 (파일이 무한히 커지지 않도록)
KEEP_QUOTES = 6


def _history_path():
    return config.REVIEW_DIR / "eval_history.jsonl"


# ═════════════════════════════════════════════════════════════════
# 1. 문제지 만들기 ─ dataset.jsonl 의 확정 교정 기록 → 문단 단위 정답표
# ═════════════════════════════════════════════════════════════════
def build_eval_set() -> dict:
    """검토 반영 데이터에서 평가셋을 만든다.

    문단(text)마다:
      golds : 사람이 최종 확정한 정답 구간 (match·miss·regrade·gold 기록)
      negs  : 사람이 지운 오탐 구간       (fp 기록) — 개선 확인용 참고
    같은 구간이 서로 다른 회차에 다르게 기록됐으면 나중 기록을 따릅니다.
    """
    path = config.REVIEW_DIR / "dataset.jsonl"
    paras: dict[str, dict] = {}
    files: set[str] = set()
    if path.exists():
        for line in path.open(encoding="utf-8"):
            try:
                rec = json.loads(line)
            except Exception:
                continue
            text = (rec.get("text") or "").strip()
            quote = (rec.get("quote") or "").strip()
            typ = rec.get("type")
            if len(text) < 6 or not quote:
                continue
            span = merge._find_span(text, quote)
            if span is None:                      # 인용구를 원문에서 못 찾으면 제외
                continue
            p = paras.setdefault(text, {"golds": {}, "negs": {}})
            files.add(rec.get("file") or "")
            if typ in ("match", "miss", "regrade", "gold"):
                p["golds"][quote] = {"quote": quote, "span": span,
                                     "grade": rec.get("grade") or "B",
                                     "risk": bool(rec.get("risk"))}
                p["negs"].pop(quote, None)
            elif typ == "fp":
                p["negs"][quote] = {"quote": quote, "span": span}
                p["golds"].pop(quote, None)

    out = [{"text": t, "golds": list(p["golds"].values()),
            "negs": list(p["negs"].values())} for t, p in paras.items()]
    return {"paras": out,
            "golds": sum(len(p["golds"]) for p in out),
            "neg_only": sum(1 for p in out if not p["golds"]),
            "files": len({f for f in files if f})}


def eval_set_summary() -> dict:
    """화면 표시용 요약 (문단 원문은 보내지 않는다)."""
    es = build_eval_set()
    return {"paras": len(es["paras"]), "golds": es["golds"],
            "neg_only": es["neg_only"], "files": es["files"]}


# ═════════════════════════════════════════════════════════════════
# 2. 한 가지 설정으로 채점 ─ 실제 분석 경로(사전→모델→병합)를 그대로 재현
# ═════════════════════════════════════════════════════════════════
def _variant_hits(text: str, base_only: bool) -> list[lexicon.Hit]:
    """규칙 사전 스캔. 최초 설정은 [사전에 추가] 된 USER_ 규칙을 뺀다."""
    hits = lexicon.scan(text)
    if base_only:
        hits = [h for h in hits if not h.rid.startswith("USER_")]
    return hits


def _run_variant(paras: list[dict], model: str, system: str,
                 base_only: bool, progress=None) -> dict[str, list[dict]]:
    """문제지 전체를 한 설정으로 분석해 문단별 예측 목록을 돌려준다."""
    segs = [Segment(seg_id=i, slide_no=1, kind="body", addr=f"eval/{i}", text=p["text"])
            for i, p in enumerate(paras, 1)]
    hits_by_seg = {s.seg_id: _variant_hits(s.text, base_only) for s in segs}

    budget = max(config.NUM_CTX - len(system) - 1500, 1200)
    opts = config.RunOptions(model=model, think=False)
    findings: list[analyze.Finding] = []
    for chunk in analyze._batch(segs, budget):
        if STATE["cancel"].is_set():
            raise _Cancelled()
        hints = {s.seg_id: sorted({h.category for h in hits_by_seg.get(s.seg_id, [])})
                 for s in chunk}
        hints = {k: v for k, v in hints.items() if v}
        findings += analyze._analyze_batch("", 1, 1, chunk, hints, opts,
                                           system_override=system)
        if progress:
            progress()

    # 실제 파이프라인과 같은 병합·안전망·오탐 필터를 통과시킨다
    deck = SimpleNamespace(segments=segs)
    resolved, _marks = merge.resolve(deck, findings, hits_by_seg)
    preds: dict[str, list[dict]] = {}
    for f in resolved:
        text = segs[f.seg_id - 1].text
        preds.setdefault(text, []).append(
            {"span": tuple(f.span) if f.span else None, "grade": f.grade,
             "risk": bool(f.disclosure_risk), "quote": (f.quote or "")[:80]})
    return preds


# ═════════════════════════════════════════════════════════════════
# 3. 채점 ─ 예측과 정답을 맞대어 재현율·정밀도·F1·등급일치를 계산
# ═════════════════════════════════════════════════════════════════
def _ov(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return not (a[1] <= b[0] or a[0] >= b[1])


def _covers(pred: dict, gold: dict) -> bool:
    """예측이 정답 구간을 (일부라도) 덮는가. span=None 은 문단 전체 마킹."""
    return pred["span"] is None or _ov(pred["span"], tuple(gold["span"]))


def _overlap_len(pred: dict, gold: dict) -> int:
    if pred["span"] is None:
        return gold["span"][1] - gold["span"][0]
    a, b = pred["span"], gold["span"]
    return max(0, min(a[1], b[1]) - max(a[0], b[0]))


def _pct(a: int, b: int):
    return round(a / b * 100, 1) if b else None


def _score(paras: list[dict], preds_by_text: dict[str, list[dict]]) -> dict:
    tp = miss = grade_ok = pred_total = pred_correct = 0
    miss_items: list[dict] = []          # 못 잡은 정답
    fp_items: list[dict] = []            # 잘못 잡은 예측
    fp_spans: dict[str, list] = {}       # 문단별 오탐 구간 (설정 간 비교용)

    for p in paras:
        text = p["text"]
        preds = preds_by_text.get(text, [])
        pred_total += len(preds)

        for g in p["golds"]:
            hit = [pr for pr in preds if _covers(pr, g)]
            if hit:
                tp += 1
                best = max(hit, key=lambda pr: _overlap_len(pr, g))
                if best["grade"] == g["grade"] and best["risk"] == bool(g["risk"]):
                    grade_ok += 1
            else:
                miss += 1
                miss_items.append({"text": text, "quote": g["quote"],
                                   "grade": g["grade"]})

        spans = []
        for pr in preds:
            if any(_covers(pr, g) for g in p["golds"]):
                pred_correct += 1
            else:
                span = pr["span"] or (0, len(text))
                spans.append(span)
                fp_items.append({"text": text, "quote": pr["quote"], "span": span})
        fp_spans[text] = spans

    golds_total = tp + miss
    r = tp / golds_total if golds_total else None
    pr_ = pred_correct / pred_total if pred_total else None
    f1 = (round(2 * r * pr_ / (r + pr_) * 100, 1)
          if r is not None and pr_ is not None and (r + pr_) else None)
    return {"tp": tp, "miss": miss, "fp": pred_total - pred_correct,
            "preds": pred_total,
            "recall": _pct(tp, golds_total),
            "precision": _pct(pred_correct, pred_total),
            "f1": f1, "grade_acc": _pct(grade_ok, tp),
            "miss_items": miss_items, "fp_items": fp_items, "fp_spans": fp_spans}


def _fixes(base: dict, cur: dict) -> dict:
    """최초 설정과 현재 설정의 항목 단위 비교 — 무엇이 좋아지고 나빠졌나."""
    b_miss = {(m["text"], m["quote"]) for m in base["miss_items"]}
    c_miss = {(m["text"], m["quote"]) for m in cur["miss_items"]}

    def _unmatched(items, other_spans):
        """한쪽에만 있는 오탐 (같은 문단에서 구간이 겹치면 같은 오탐으로 본다)."""
        return [it for it in items
                if not any(_ov(tuple(it["span"]), tuple(s))
                           for s in other_spans.get(it["text"], []))]

    fp_fixed = _unmatched(base["fp_items"], cur["fp_spans"])
    new_fp = _unmatched(cur["fp_items"], base["fp_spans"])
    return {
        "miss_fixed": len(b_miss - c_miss),        # 최초가 놓친 것을 현재는 잡음
        "regressed": len(c_miss - b_miss),         # 최초는 잡았는데 현재가 놓침(퇴행)
        "fp_fixed": len(fp_fixed),                 # 최초의 오탐이 현재는 사라짐
        "new_fp": len(new_fp),                     # 현재 새로 생긴 오탐
        "miss_fixed_quotes": [q for _t, q in sorted(b_miss - c_miss)][:KEEP_QUOTES],
        "regressed_quotes": [q for _t, q in sorted(c_miss - b_miss)][:KEEP_QUOTES],
        "fp_fixed_quotes": [it["quote"] for it in fp_fixed][:KEEP_QUOTES],
        "still_miss_quotes": [m["quote"] for m in cur["miss_items"]][:KEEP_QUOTES],
    }


# ═════════════════════════════════════════════════════════════════
# 4. 이력 ─ 회차마다 설정 스냅샷과 "직전 대비 무엇이 바뀌었나" 를 함께 기록
# ═════════════════════════════════════════════════════════════════
def _config_snapshot() -> dict:
    rules: list[str] = []
    try:
        entries = json.loads((config.REVIEW_DIR / "extra_rules.json")
                             .read_text(encoding="utf-8"))
        rules = [e.get("keyword") or "" for e in entries if e.get("keyword")]
    except Exception:
        pass
    st = review.stats()
    # examples = 쌓인 풀 크기, injected = 실제로 프롬프트에 들어간 칸 수.
    # 프롬프트를 바꾸는 것은 injected 쪽이다. 풀은 계속 커지지만 판정에는 영향이 없으므로
    # 둘을 나눠 기록해야 "점수가 왜 올랐는지" 를 나중에 설명할 수 있다.
    return {"model": config.MODEL, "rules": rules,
            "examples": st["examples"], "injected": st["injected"],
            "dataset": st["dataset"]}


def _injected_change(prev_n, now_n) -> str | None:
    if prev_n is None:
        return f"기록 없음 → {now_n}"
    return f"{prev_n} → {now_n}" if prev_n != now_n else None


def _changes(prev: dict | None, snap: dict) -> dict | None:
    """직전 측정 이후 바뀐 것. 첫 측정이면 None."""
    if not prev:
        return None
    p = prev.get("config") or {}
    return {
        "rules_added": [r for r in snap["rules"] if r not in (p.get("rules") or [])],
        "rules_removed": [r for r in (p.get("rules") or []) if r not in snap["rules"]],
        "examples_delta": snap["examples"] - (p.get("examples") or 0),
        # 주입 칸 수가 달라졌다면 프롬프트 자체가 바뀐 것이라 회차 간 비교가 끊긴다.
        # 점수 변화의 원인이 '교정이 좋아져서' 가 아닐 수 있으므로 반드시 표시한다.
        # 이 항목이 없던 예전 회차(= 도구를 올린 직후 첫 회차)는 프롬프트가 가장 크게
        # 바뀐 회차인데도 '변경 없음' 으로 보이므로, 모른다는 사실을 그대로 적는다.
        "injected_change": _injected_change(p.get("injected"), snap["injected"]),
        "dataset_delta": snap["dataset"] - (p.get("dataset") or 0),
        "model_change": (f'{p.get("model")} → {snap["model"]}'
                         if p.get("model") and p.get("model") != snap["model"] else None),
    }


def history() -> list[dict]:
    p = _history_path()
    if not p.exists():
        return []
    out = []
    for line in p.open(encoding="utf-8"):
        try:
            out.append(json.loads(line))
        except Exception:
            continue
    return out


# ═════════════════════════════════════════════════════════════════
# 5. 실행 ─ 웹 화면의 [성능 측정 실행] 버튼과 연결
# ═════════════════════════════════════════════════════════════════
class _Cancelled(Exception):
    pass


# 진행 상황. main.py 의 /api/eval/status 가 이 값을 화면으로 보낸다.
STATE: dict = {"running": False, "stage": "", "done": 0, "total": 0,
               "error": "", "cancel": threading.Event()}
_LOCK = threading.Lock()


def status() -> dict:
    return {k: STATE[k] for k in ("running", "stage", "done", "total", "error")}


def cancel() -> None:
    STATE["cancel"].set()


def _total_batches(paras: list[dict], systems: list[str]) -> int:
    segs = [SimpleNamespace(seg_id=i, text=p["text"])
            for i, p in enumerate(paras, 1)]
    n = 0
    for sys_ in systems:
        budget = max(config.NUM_CTX - len(sys_) - 1500, 1200)
        n += len(analyze._batch(segs, budget))
    return n


def run_sync(note: str = "") -> dict:
    """측정 한 회차를 처음부터 끝까지 실행하고 이력 항목을 돌려준다.

    (웹 화면은 start() 로 별도 스레드에서 부르고, 테스트는 이 함수를 직접 부른다)
    """
    t0 = time.time()
    es = build_eval_set()
    if not es["paras"]:
        raise ValueError("평가셋이 비어 있습니다 — 먼저 [검토 반영] 탭에서 "
                         "검토완료본을 올려 [반영 저장] 을 해 주세요.")

    h = analyze.health()
    if not h.get("ok"):
        raise ValueError(f"Ollama 에 연결할 수 없습니다 ({config.OLLAMA_HOST}). "
                         "`ollama serve` 실행 여부를 확인하세요.")
    cur_model = config.MODEL
    if not analyze.model_available(cur_model, h.get("models", [])):
        raise ValueError(f"모델 {cur_model} 이 없습니다. `ollama pull {cur_model}` "
                         "로 먼저 내려받으세요.")
    base_model, base_fallback = BASELINE_MODEL, False
    if not analyze.model_available(base_model, h.get("models", [])):
        base_model, base_fallback = cur_model, True   # 8b 가 없으면 현재 모델로 대신

    base_system = analyze.SYSTEM                       # 예시 주입 없는 원래 지시서
    cur_system = analyze._system_prompt()              # 확정 사례가 붙은 지시서
    # 스냅샷은 '이번 회차가 실제로 쓴 설정' 이어야 한다. 채점이 끝난 뒤에 찍으면
    # 측정 도중에 [반영 저장] 이 들어온 경우 쓰지도 않은 설정이 기록된다.
    snap = _config_snapshot()
    paras = es["paras"]
    STATE.update(total=_total_batches(paras, [base_system, cur_system]), done=0)

    def tick():
        STATE["done"] += 1

    STATE["stage"] = f"최초 설정으로 채점 중 ({base_model})"
    base_preds = _run_variant(paras, base_model, base_system, base_only=True,
                              progress=tick)
    STATE["stage"] = f"현재 설정으로 채점 중 ({cur_model})"
    cur_preds = _run_variant(paras, cur_model, cur_system, base_only=False,
                             progress=tick)

    STATE["stage"] = "채점 집계 중"
    base_sc = _score(paras, base_preds)
    cur_sc = _score(paras, cur_preds)

    def _pub(sc: dict, model: str) -> dict:
        return {"model": model, "tp": sc["tp"], "miss": sc["miss"], "fp": sc["fp"],
                "preds": sc["preds"], "recall": sc["recall"],
                "precision": sc["precision"], "f1": sc["f1"],
                "grade_acc": sc["grade_acc"]}

    def _d(a, b):
        return round(a - b, 1) if a is not None and b is not None else None

    prev_entries = history()
    entry = {
        "run_no": len(prev_entries) + 1,
        "ts": datetime.datetime.now().isoformat(timespec="seconds"),
        "note": note,
        "eval_set": {"paras": len(paras), "golds": es["golds"],
                     "neg_only": es["neg_only"], "files": es["files"]},
        "baseline": {**_pub(base_sc, base_model), "fallback": base_fallback},
        "current": {**_pub(cur_sc, cur_model),
                    "user_rules": len(snap["rules"]), "examples": snap["examples"]},
        "delta": {"recall": _d(cur_sc["recall"], base_sc["recall"]),
                  "precision": _d(cur_sc["precision"], base_sc["precision"]),
                  "f1": _d(cur_sc["f1"], base_sc["f1"])},
        "fixes": _fixes(base_sc, cur_sc),
        "config": snap,
        "changes": _changes(prev_entries[-1] if prev_entries else None, snap),
        "duration_s": round(time.time() - t0, 1),
    }
    config.REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    with _history_path().open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return entry


def _worker(note: str) -> None:
    try:
        run_sync(note)
        STATE["stage"] = "완료"
    except _Cancelled:
        STATE["error"], STATE["stage"] = "사용자가 중단했습니다.", "중단됨"
    except Exception as e:  # noqa: BLE001  (어떤 오류든 화면에 보여 주기 위해)
        STATE["error"], STATE["stage"] = f"{type(e).__name__}: {e}", "오류"
    finally:
        STATE["running"] = False


def start(note: str = "") -> None:
    """측정을 별도 스레드에서 시작한다. 문제가 있으면 즉시 ValueError."""
    with _LOCK:
        if STATE["running"]:
            raise RuntimeError("이미 측정이 진행 중입니다.")
        if not build_eval_set()["paras"]:
            raise ValueError("평가셋이 비어 있습니다 — 먼저 [검토 반영] 탭에서 "
                             "검토완료본을 올려 [반영 저장] 을 해 주세요.")
        h = analyze.health()
        if not h.get("ok"):
            raise ValueError(f"Ollama 에 연결할 수 없습니다 ({config.OLLAMA_HOST}). "
                             "`ollama serve` 실행 여부를 확인하세요.")
        if not analyze.model_available(config.MODEL, h.get("models", [])):
            raise ValueError(f"모델 {config.MODEL} 이 없습니다. "
                             f"`ollama pull {config.MODEL}` 로 먼저 내려받으세요.")
        STATE.update(running=True, stage="준비 중", done=0, total=0, error="")
        STATE["cancel"].clear()
    threading.Thread(target=_worker, args=(note,), daemon=True).start()
