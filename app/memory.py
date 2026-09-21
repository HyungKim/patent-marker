"""
memory.py ─ 검토 데이터가 쌓일수록 계속 나아지게 하는 세 장치 ("검토 기억")
=====================================================================

[이 파일이 하는 일]
  [검토 반영] 탭이 쌓아 둔 review_data/dataset.jsonl (사람이 확정한 교정 내역) 을 읽어
  다음 분석에 세 가지 방식으로 되살립니다. 모델(qwen3)을 학습시키는 것이 아니라,
  모델 **주변에 기억을 쌓는** 방식이라 GPU 없이 CPU 에서 돌고, 데이터 양에 비례해 좋아집니다.

    ① 문단 기억   같은 문단(90% 이상 같은 문단 포함)이 다시 나오면 모델에게 묻지 않고
                   사람의 최종 판정(형광펜 위치·등급)을 그대로 적용한다.  → apply_memory()
    ② 제외 사전   검토에서 지운 표현을 모델이 또 잡으면 결과에서 뺀다.       → apply_excludes()
       자동 규칙   검토에서 추가한 표현은 규칙 사전에 자동 등록된다
                   (숫자·공백은 달라도, 핵심어 어순이 바뀌어도 잡힘).            → auto_hits()
    ③ 유사 사례   지금 분석하는 슬라이드와 비슷한 확정 사례를 골라 모델 지시서에 붙인다.
                   비슷함은 글자 겹침(기본) 또는 bge-m3 임베딩(설치돼 있으면 '뜻' 기준).  → examples_for()

[켜고 끄기]
  실행 옵션(config.RunOptions.learn / learn_embed) 과 환경 변수(PM_LEARN, PM_LEARN_EMBED)로 켜고 끕니다.
  화면의 "검토 학습 적용" · "뜻 기준 검색(bge-m3)" 체크박스, 명령행 --no-learn / --no-embed 가 이 값을 정합니다.
  분석 스레드마다 set_context() 로 현재 옵션을 걸어 두면, 규칙 사전(lexicon.scan)이 자동 규칙을 붙일지 결정합니다.

[정직한 채점을 위해]
  성능 측정(evaluate.py)은 채점하는 바로 그 문단에서 나온 기억·규칙·사례를 빼고(exclude_texts) 돌립니다.
  그래야 "외워서 맞힌 것" 이 아니라 "비슷한 문단으로 옮겨 가는 힘" 이 점수에 잡힙니다.

[임베딩(bge-m3)]
  Ollama 에 bge-m3 가 있으면 /api/embed 로 문장을 1,024개 숫자(좌표)로 바꿔 뜻이 가까운 사례를 찾습니다.
  없거나 실패하면 조용히 글자 겹침으로 내려갑니다. 사례의 좌표는 review_data/embed_cache.json 에 저장해 재사용합니다.
"""
from __future__ import annotations

import difflib
import hashlib
import json
import math
import re
import threading
import time
import urllib.request
from dataclasses import dataclass, field

from . import config

# ═════════════════════════════════════════════════════════════════
# 실행 옵션(스레드별) — 규칙 사전이 자동 규칙을 붙일지, 채점 때 무엇을 뺄지
# ═════════════════════════════════════════════════════════════════
_ctx = threading.local()


def set_context(learn: bool, exclude_texts: set[str] | None = None) -> None:
    """분석 스레드가 시작할 때 부른다. learn=False 면 이 스레드에서는 세 장치가 모두 꺼진다."""
    _ctx.learn = learn
    _ctx.exclude = {_norm(t) for t in (exclude_texts or set())}


def clear_context() -> None:
    _ctx.learn = None
    _ctx.exclude = set()


def _learn_on() -> bool:
    v = getattr(_ctx, "learn", None)
    return config.LEARN if v is None else bool(v)


def _excluded() -> set[str]:
    return getattr(_ctx, "exclude", set()) or set()


# ═════════════════════════════════════════════════════════════════
# 글자 다루기
# ═════════════════════════════════════════════════════════════════
_WS = re.compile(r"\s+")
_NUM = re.compile(r"\d+(?:[.,]\d+)?")
_TOKEN = re.compile(r"[가-힣A-Za-z0-9㎛㎜%±°~]+")
# 한국어 조사·어미 — 핵심어를 뽑을 때 꼬리에서 떼어 낸다 (형태소 분석 없이 쓰는 간단한 목록)
_PARTICLES = ("으로써", "으로서", "으로는", "에서는", "에서", "에게", "까지", "부터", "이며", "하며", "하고",
              "되어", "된다", "한다", "하는", "되는", "으로", "로써", "로서", "이다", "였다", "된", "한", "들",
              "의", "을", "를", "이", "가", "은", "는", "도", "과", "와", "에", "로", "만")


def _norm(text: str) -> str:
    return _WS.sub(" ", (text or "")).strip().lower()


def _bigrams(text: str) -> set[str]:
    """두 글자 조각 집합. 공백을 빼고, 숫자는 전부 '0' 으로 보아 값만 바뀐 문장을 같게 본다."""
    s = re.sub(r"\d", "0", _WS.sub("", (text or "")).lower())
    return {s[i:i + 2] for i in range(len(s) - 1)} if len(s) > 1 else ({s} if s else set())


def similarity(a: str, b: str) -> float:
    """두 글자열의 '글자 겹침' 유사도 (0~1). 두 글자 조각 집합의 Dice 계수."""
    ga, gb = _bigrams(a), _bigrams(b)
    if not ga or not gb:
        return 0.0
    return 2 * len(ga & gb) / (len(ga) + len(gb))


def keywords(quote: str, limit: int = 4) -> list[str]:
    """인용구에서 핵심어를 뽑는다. 조사·어미를 떼고 두 글자 이상, 숫자만인 것은 뺀다."""
    out: list[str] = []
    for tok in _TOKEN.findall(quote or ""):
        t = tok
        for p in _PARTICLES:
            if len(t) - len(p) >= 2 and t.endswith(p) and re.search(r"[가-힣]", t):
                t = t[:-len(p)]
                break
        if len(t) < 2 or _NUM.fullmatch(t) or t in out:
            continue
        out.append(t)
    out.sort(key=len, reverse=True)
    return out[:limit]


def phrase_pattern(quote: str) -> re.Pattern | None:
    """인용구를 '숫자·공백이 달라도 맞는' 정규식으로. 예: 원가 40% 절감 → 원가\\s*\\d+…%\\s*절감"""
    q = _WS.sub(" ", (quote or "")).strip()
    if len(q) < 3:
        return None
    parts = []
    for piece in re.split(r"(\d+(?:[.,]\d+)?|\s+)", q):
        if not piece:
            continue
        if _NUM.fullmatch(piece):
            parts.append(r"\s*\d+(?:[.,]\d+)?\s*")     # 숫자는 값이 달라도, 앞뒤 공백이 있어도
        elif piece.isspace():
            parts.append(r"\s*")
        else:
            parts.append(re.escape(piece))
    try:
        return re.compile("".join(parts))
    except re.error:
        return None


def _trim(text: str, span: tuple[int, int]) -> tuple[int, int]:
    """구간 양끝의 공백을 뗀다 (패턴의 \\s* 가 앞뒤 공백까지 물고 오는 것을 정리)."""
    a, b = span
    while a < b and text[a].isspace():
        a += 1
    while b > a and text[b - 1].isspace():
        b -= 1
    return (a, b)


def _sentences(text: str) -> list[tuple[int, int]]:
    """문단을 문장 단위 (시작, 끝) 로 자른다."""
    spans, start = [], 0
    for m in re.finditer(r"[.!?。]\s+|\n", text):
        spans.append((start, m.end()))
        start = m.end()
    if start < len(text):
        spans.append((start, len(text)))
    return spans or [(0, len(text))]


def locate(text: str, quote: str) -> tuple[int, int] | None:
    """검토에서 확정한 인용구를 (조금 바뀐) 새 문단에서 찾는다.
    1) 그대로  2) 숫자·공백 무시  3) 가장 길게 겹치는 조각이 인용구의 60% 이상이면 그 조각"""
    if not quote:
        return None
    i = text.find(quote)
    if i >= 0:
        return (i, i + len(quote))
    pat = phrase_pattern(quote)
    if pat:
        m = pat.search(text)
        if m:
            return _trim(text, m.span())
    sm = difflib.SequenceMatcher(None, text, quote, autojunk=False)
    blk = sm.find_longest_match(0, len(text), 0, len(quote))
    if blk.size >= max(4, int(len(quote) * 0.6)):
        return (blk.a, blk.a + blk.size)
    # 긴 인용구(문장 전체 등)는 단어 몇 개가 바뀌면 한 덩어리로는 안 맞는다 → 맞는 조각들의 합이
    # 인용구의 60% 이상이면 첫 조각~마지막 조각을 자리로 본다 (사람이 문단 전체를 칠한 경우가 여기에 해당)
    blocks = [b for b in sm.get_matching_blocks() if b.size >= 3]
    if blocks and sum(b.size for b in blocks) >= max(6, int(len(quote) * 0.6)):
        return _trim(text, (blocks[0].a, blocks[-1].a + blocks[-1].size))
    return None


# ═════════════════════════════════════════════════════════════════
# 검토 데이터 → 기억 색인
# ═════════════════════════════════════════════════════════════════
@dataclass
class Para:
    """검토를 거친 문단 하나의 최종 상태."""

    text: str
    golds: list[dict]                    # {quote, grade, risk, category}
    negs: list[dict]                     # {quote, category}
    date: str = ""
    grams: set = field(default_factory=set)   # 두 글자 조각 (유사도 계산용, 한 번만 만든다)


@dataclass
class Rule:
    """검토에서 추가된 표현으로 만든 자동 규칙."""

    rid: str
    quote: str
    category: str
    risk: bool
    source: str                          # 어느 문단에서 나왔나 (_norm)
    pattern: re.Pattern | None
    tokens: list[str]
    grade: str = "B"


@dataclass
class Index:
    paras: dict[str, Para] = field(default_factory=dict)      # _norm(text) → Para
    rules: list[Rule] = field(default_factory=list)
    excludes: list[dict] = field(default_factory=list)         # {quote, pattern, norm, source}
    examples: list[dict] = field(default_factory=list)         # {kind, quote, text, grade, risk, category, source, grams}
    built_from: tuple = ()


_INDEX: Index | None = None
_LOCK = threading.Lock()


def _dataset_path():
    return config.REVIEW_DIR / "dataset.jsonl"


def _records() -> list[dict]:
    p = _dataset_path()
    if not p.exists():
        return []
    out = []
    with p.open(encoding="utf-8") as fh:
        for line in fh:
            try:
                out.append(json.loads(line))
            except Exception:  # noqa: BLE001  (깨진 줄은 건너뜀)
                continue
    return out


def _cat(c) -> str:
    return c if c in config.CATEGORIES else "블랙박스용어"


def _build() -> Index:
    """dataset.jsonl 전체를 읽어 색인을 만든다. 같은 구간이 여러 번 기록됐으면 나중 것을 따른다."""
    idx = Index()
    paras: dict[str, Para] = {}
    for rec in _records():
        text = (rec.get("text") or "").strip()
        quote = (rec.get("quote") or "").strip()
        typ = rec.get("type")
        if len(text) < 6 or not quote:
            continue
        key = _norm(text)
        p = paras.get(key)
        if p is None:
            p = paras[key] = Para(text=text, golds=[], negs=[], date=rec.get("date") or "")
        elif len(text) > len(p.text):
            p.text = text                     # 예전엔 400자로 잘라 저장했으므로 긴 쪽을 남긴다
        p.date = rec.get("date") or p.date
        p.golds = [g for g in p.golds if g["quote"] != quote]
        p.negs = [n for n in p.negs if n["quote"] != quote]
        if typ in ("match", "miss", "regrade", "gold"):
            p.golds.append({"quote": quote, "grade": (rec.get("grade") or "B").upper(),
                            "risk": bool(rec.get("risk")), "category": _cat(rec.get("category"))})
        elif typ == "fp":
            p.negs.append({"quote": quote, "category": _cat(rec.get("category"))})
    for p in paras.values():
        p.grams = _bigrams(p.text)
    idx.paras = paras

    gold_quotes = {_norm(g["quote"]) for p in paras.values() for g in p.golds}
    neg_quotes = {_norm(n["quote"]) for p in paras.values() for n in p.negs}
    seen_rules: set[str] = set()
    seen_ex: set[str] = set()
    n = 0
    for key, p in paras.items():
        for g in p.golds:
            q = _norm(g["quote"])
            if q in neg_quotes:               # 어디선가는 지워진 표현 — 규칙으로 굳히지 않는다
                continue
            if q not in seen_rules and len(q) >= 3:
                seen_rules.add(q)
                n += 1
                idx.rules.append(Rule(rid=f"AUTO_{n}", quote=g["quote"], category=g["category"],
                                      risk=g["risk"], source=key, pattern=phrase_pattern(g["quote"]),
                                      tokens=keywords(g["quote"]), grade=g["grade"]))
            if q not in seen_ex:
                seen_ex.add(q)
                idx.examples.append({"kind": "include", "quote": g["quote"], "text": p.text,
                                     "grade": g["grade"], "risk": g["risk"], "category": g["category"],
                                     "source": key, "grams": _bigrams(p.text)})
        for ng in p.negs:
            q = _norm(ng["quote"])
            if q in gold_quotes:
                continue
            if len(q) >= 6 and q not in seen_ex:
                idx.excludes.append({"quote": ng["quote"], "norm": q, "source": key,
                                     "pattern": phrase_pattern(ng["quote"])})
            if q not in seen_ex:
                seen_ex.add(q)
                idx.examples.append({"kind": "exclude", "quote": ng["quote"], "text": p.text,
                                     "grade": "", "risk": False, "category": ng["category"],
                                     "source": key, "grams": _bigrams(p.text)})
    return idx


def index() -> Index:
    """색인을 돌려준다. dataset.jsonl 이 바뀌었으면 다시 만든다 (몇 백 건이면 수십 ms)."""
    global _INDEX
    p = _dataset_path()
    stamp = (str(p), p.stat().st_mtime if p.exists() else 0, p.stat().st_size if p.exists() else 0)
    with _LOCK:
        if _INDEX is None or _INDEX.built_from != stamp:
            _INDEX = _build()
            _INDEX.built_from = stamp
        return _INDEX


def stats() -> dict:
    """화면 안내용 개수."""
    ix = index()
    return {"paras": len(ix.paras), "rules": len(ix.rules), "excludes": len(ix.excludes),
            "examples": len(ix.examples), "embed_model": config.EMBED_MODEL}


# ═════════════════════════════════════════════════════════════════
# ② 자동 규칙 — 규칙 사전(lexicon.scan)이 부른다
# ═════════════════════════════════════════════════════════════════
def auto_hits(text: str) -> list:
    """검토에서 추가된 표현이 이 문단에 (변형된 꼴로라도) 있으면 Hit 목록으로 돌려준다."""
    if not _learn_on():
        return []
    from .lexicon import Hit                      # 순환 import 방지 — 호출 시점에 가져온다

    ix = index()
    if not ix.rules:
        return []
    excl = _excluded()
    key = _norm(text)
    hits = []
    sents = None
    for r in ix.rules:
        if excl and r.source in excl:
            continue
        if r.source == key:                       # 자기 자신에서 나온 규칙은 채점 때 의미가 없다
            pass
        m = r.pattern.search(text) if r.pattern else None
        span = m.span() if m else None
        if span is None and len(r.tokens) >= 2:
            # 핵심어가 같은 문장 안에 모두 있으면 (어순 무관) 첫 핵심어~마지막 핵심어 구간.
            # 확정 표현이 문장 여러 개에 걸친 긴 것이면 문단 전체에서도 찾는다.
            if sents is None:
                sents = _sentences(text)
            windows = list(sents) + ([(0, len(text))] if len(r.quote) > 40 else [])
            for s, e in windows:
                sent = text[s:e]
                pos = [sent.find(t) for t in r.tokens]
                if all(pp >= 0 for pp in pos):
                    lo = min(pos)
                    hi = max(pp + len(t) for pp, t in zip(pos, r.tokens))
                    span = (s + lo, s + hi)
                    break
        if span is None:
            continue
        span = _trim(text, span)
        if span[0] >= span[1]:
            continue
        disclosure = r.risk and r.category == "공개이력"
        hits.append(Hit(r.rid, r.category, 8, span, text[span[0]:span[1]],
                        f"검토에서 확정된 표현과 유사: '{r.quote}'", disclosure))
    return hits


# ═════════════════════════════════════════════════════════════════
# ① 문단 기억 · ② 제외 사전 — 병합 결과에 적용
# ═════════════════════════════════════════════════════════════════
def match_para(text: str, exclude: set[str] | None = None) -> Para | None:
    """이 문단과 같은(또는 90% 이상 같은) 검토 문단을 찾는다."""
    ix = index()
    if not ix.paras:
        return None
    key = _norm(text)
    excl = exclude if exclude is not None else _excluded()
    if key in ix.paras and key not in excl:
        return ix.paras[key]
    best, best_s = None, 0.0
    n = len(key)
    g = _bigrams(key)
    if not g:
        return None
    for k, p in ix.paras.items():
        if k in excl or abs(len(k) - n) > max(20, int(n * 0.3)) or not p.grams:
            continue
        s = 2 * len(g & p.grams) / (len(g) + len(p.grams))
        if s > best_s:
            best, best_s = p, s
    return best if best_s >= config.MEMORY_SIM else None


def apply_memory(deck, resolved: list, make_finding) -> tuple[list, dict]:
    """①: 기억된 문단은 사람 판정으로 통째로 바꾼다.

    make_finding(seg, quote, span, grade, risk, category, reason) → Finding  (analyze.Finding 생성기)
    돌려주는 값: (새 후보 목록, {"paras": 적용 문단 수, "added": 추가 건, "dropped": 뺀 건})
    """
    out, st = [], {"paras": 0, "added": 0, "dropped": 0}
    if not _learn_on() or not index().paras:
        return resolved, st
    remembered: dict[int, Para] = {}
    for seg in deck.segments:
        if len(seg.text.strip()) < 6:
            continue
        p = match_para(seg.text)
        if p is not None:
            remembered[seg.seg_id] = p
    if not remembered:
        return resolved, st
    seg_map = {s.seg_id: s for s in deck.segments}
    replacement: dict[int, list] = {}
    for seg_id, p in remembered.items():
        seg = seg_map[seg_id]
        new_items = []
        for g in p.golds:
            span = locate(seg.text, g["quote"])
            if span is None:
                continue
            new_items.append(make_finding(seg, seg.text[span[0]:span[1]], span, g["grade"], g["risk"],
                                          g["category"], f"검토 기억 ({p.date} 검토에서 확정)"))
        if p.golds and not new_items:
            continue          # 정답 자리를 하나도 못 찾았다 — 기억을 적용하지 않고 모델 답을 그대로 둔다 (안전장치)
        replacement[seg_id] = new_items
    for f in resolved:
        if f.seg_id in replacement:
            st["dropped"] += 1
            continue
        out.append(f)
    for seg_id, items in replacement.items():
        st["paras"] += 1
        st["added"] += len(items)
        out.extend(items)
    return out, st


def apply_excludes(resolved: list) -> tuple[list, int]:
    """②: 검토에서 지운 표현을 (검토 기억이 아닌) 후보가 또 잡았으면 뺀다."""
    if not _learn_on():
        return resolved, 0
    ix = index()
    if not ix.excludes:
        return resolved, 0
    excl = _excluded()
    out, dropped = [], 0
    for f in resolved:
        if getattr(f, "source", "") == "memory" or f.disclosure_risk:
            out.append(f)
            continue
        q = _norm(f.quote)
        hit = False
        for ex in ix.excludes:
            if excl and ex["source"] in excl:
                continue
            if ex["norm"] in q or (len(q) >= 6 and q in ex["norm"]):
                hit = True
            elif ex["pattern"] is not None and ex["pattern"].search(f.quote or ""):
                hit = True
            if hit:
                break
        if hit:
            dropped += 1
        else:
            out.append(f)
    return out, dropped


def rescue_confirmed(deck, resolved: list, hits_by_seg: dict, make_finding) -> tuple[list, int]:
    """②-보강: 검토에서 확정된 표현(자동 규칙)이 걸린 구간이 아직 아무 후보로도 덮여 있지 않으면 후보로 넣는다.

    merge 의 안전망은 '문단 단위' 라 모델이 같은 문단에서 다른 구간만 잡으면 확정 표현이 빠진다.
    사람이 확정한 표현은 구간 단위로 반드시 남긴다. 돌려주는 값: (후보 목록, 추가 건수)
    """
    if not _learn_on():
        return resolved, 0
    ix = index()
    if not ix.rules:
        return resolved, 0
    rules = {r.rid: r for r in ix.rules}
    seg_map = {s.seg_id: s for s in deck.segments}
    covered: dict[int, list] = {}
    for f in resolved:
        covered.setdefault(f.seg_id, []).append(f.span)
    added = 0
    for seg_id, hits in hits_by_seg.items():
        seg = seg_map.get(seg_id)
        if seg is None:
            continue
        for h in hits:
            r = rules.get(h.rid)
            if r is None:
                continue
            spans = covered.setdefault(seg_id, [])
            if any(sp is None or not (sp[1] <= h.span[0] or sp[0] >= h.span[1]) for sp in spans):
                continue                                  # 이미 (일부라도) 덮여 있음
            resolved.append(make_finding(seg, seg.text[h.span[0]:h.span[1]], h.span, r.grade, r.risk,
                                         r.category, f"검토에서 확정된 표현과 유사: '{r.quote}'"))
            spans.append(h.span)
            added += 1
    return resolved, added


# ═════════════════════════════════════════════════════════════════
# ③ 유사 사례 검색 (글자 겹침 / bge-m3 임베딩)
# ═════════════════════════════════════════════════════════════════
_DIRECT = urllib.request.build_opener(urllib.request.ProxyHandler({}))
_EMBED_CACHE: dict[str, list[float]] | None = None


def _cache_path():
    return config.REVIEW_DIR / "embed_cache.json"


def _load_cache() -> dict[str, list[float]]:
    global _EMBED_CACHE
    if _EMBED_CACHE is None:
        try:
            _EMBED_CACHE = json.loads(_cache_path().read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            _EMBED_CACHE = {}
    return _EMBED_CACHE


def _save_cache() -> None:
    try:
        _cache_path().parent.mkdir(parents=True, exist_ok=True)
        _cache_path().write_text(json.dumps(_EMBED_CACHE or {}), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass


def _sha(text: str) -> str:
    return hashlib.sha1(_norm(text).encode("utf-8")).hexdigest()[:20]


def embed_available(models: list[str] | None = None) -> bool:
    """bge-m3 가 Ollama 에 내려받아져 있는가."""
    if models is None:
        try:
            from . import analyze
            models = analyze.health().get("models", [])
        except Exception:  # noqa: BLE001
            return False
    from .analyze import model_available
    return model_available(config.EMBED_MODEL, models)


def embed(texts: list[str], timeout: int = 300, store: bool = True) -> list[list[float]]:
    """Ollama /api/embed 로 좌표를 구한다. 이미 구한 것은 캐시에서. 실패하면 예외.

    store=True 는 검토 사례처럼 다시 쓸 문장용(파일에 저장). 분석 중인 문서의 문단은 store=False 로
    한 번만 쓰고 버린다 — 그래야 캐시 파일이 문서를 돌릴 때마다 커지지 않는다.
    """
    cache = _load_cache()
    todo = list(dict.fromkeys(t for t in texts if _sha(t) not in cache))
    fresh: dict[str, list[float]] = {}
    if todo:
        body = json.dumps({"model": config.EMBED_MODEL, "input": todo}).encode()
        req = urllib.request.Request(f"{config.OLLAMA_HOST}/api/embed", data=body,
                                     headers={"Content-Type": "application/json"}, method="POST")
        with _DIRECT.open(req, timeout=timeout) as r:
            data = json.loads(r.read().decode())
        vecs = data.get("embeddings") or []
        if len(vecs) != len(todo):
            raise RuntimeError("임베딩 개수가 맞지 않음")
        for t, v in zip(todo, vecs):
            fresh[_sha(t)] = [round(x, 5) for x in v]
        if store:
            cache.update(fresh)
            _save_cache()
    return [cache.get(_sha(t)) or fresh[_sha(t)] for t in texts]


def _cos(a: list[float], b: list[float]) -> float:
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x * x for x in a)) or 1.0
    nb = math.sqrt(sum(x * x for x in b)) or 1.0
    return dot / (na * nb)


def examples_for(texts: list[str], use_embed: bool = False, k_pos: int = 8, k_neg: int = 4,
                 exclude: set[str] | None = None) -> tuple[str, dict]:
    """③: 이 슬라이드(문단 목록)와 가장 비슷한 확정 사례를 골라 지시서에 붙일 블록을 만든다.

    돌려주는 값: (블록 문자열 — 없으면 "", {"pos": n, "neg": n, "mode": "embed"|"ngram"|"off"})
    """
    info = {"pos": 0, "neg": 0, "mode": "off"}
    if not _learn_on():
        return "", info
    ix = index()
    excl = exclude if exclude is not None else _excluded()
    pool = [e for e in ix.examples if not (excl and e["source"] in excl)]
    texts = [t for t in texts if len(t.strip()) >= 6]
    if not pool or not texts:
        return "", info

    scores: list[float] = [0.0] * len(pool)
    mode = "ngram"
    if use_embed:
        try:
            qv = embed(texts, store=False)               # 분석 문서의 문단 — 저장하지 않음
            ev = embed([e["text"] for e in pool])        # 검토 사례 — 저장해 두고 재사용
            for j, e in enumerate(ev):
                scores[j] = max(_cos(q, e) for q in qv)
            mode = "embed"
        except Exception:  # noqa: BLE001  (미설치·실패 → 글자 겹침으로)
            mode = "ngram"
    if mode == "ngram":
        tg = [_bigrams(t) for t in texts]
        for j, e in enumerate(pool):
            g = e["grams"]
            best = 0.0
            for q in tg:
                if q and g:
                    s = 2 * len(q & g) / (len(q) + len(g))
                    if s > best:
                        best = s
            scores[j] = best
    order = sorted(range(len(pool)), key=lambda j: -scores[j])
    pos = [pool[j] for j in order if pool[j]["kind"] == "include"][:k_pos]
    neg = [pool[j] for j in order if pool[j]["kind"] == "exclude"][:k_neg]
    info.update(pos=len(pos), neg=len(neg), mode=mode)
    if not pos and not neg:
        return "", info
    lines = ["## 사내 검토로 확정된 사례 (이 슬라이드와 비슷한 것 — 반드시 따를 것)"]
    if neg:
        lines.append("다음 구간은 검토 결과 특허 후보가 아니었다. 유사한 문장을 반환하지 마라.")
        lines += [f'- "{e["quote"]}"' for e in neg]
    if pos:
        lines.append("다음 구간은 검토 결과 후보가 맞았다. 유사한 표현·같은 유형의 문장을 놓치지 마라.")
        lines += [f'- [{e["grade"]}{"·C공개" if e["risk"] else ""}·{e["category"]}] "{e["quote"]}"' for e in pos]
    return "\n".join(lines), info
