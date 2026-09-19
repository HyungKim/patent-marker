"""
review.py ─ 검토완료본을 읽어 "사람의 교정" 을 학습 재료로 바꾸는 단계
=====================================================================

[이 파일이 하는 일]
  사람이 PowerPoint 에서 형광펜을 고쳐 놓은 '검토완료본' 을 읽어,
  도구가 원래 내렸던 판정(분석 때 자동 저장해 둔 기록)과 비교합니다.

      검토완료 PPTX ──▶ read_reviewed() ──▶ diff() ──▶ 교정 내역
                                                        ├─ 누락   : 사람이 새로 칠함 (모델이 놓침)
                                                        ├─ 오탐   : 사람이 지움     (모델이 잘못 잡음)
                                                        ├─ 등급변경: 색을 바꿈
                                                        └─ 일치   : 그대로 둠

  [반영 저장] 을 누르면 세 곳에 자동으로 쌓입니다.
    1. dataset.jsonl   — 교정 기록 원본 (나중에 미세조정 재료)
    2. examples.json   — 판정 프롬프트에 주입되는 모범답안·반례 (analyze.py 가 읽음)
    3. extra_rules.json— [사전에 추가] 로 등록한 표현 (lexicon.py 가 읽음)

[검토자의 형광펜 색 규약]
  노랑 계열 = A (구체 수단 드러남) · 하늘/파랑 계열 = B (묵시) · 살구/빨강 계열 = ⚠ 공개 리스크
  PowerPoint 기본 형광펜 색을 써도 됩니다 — 정확한 색이 아니라 색 계열(색상환 거리)로 판정합니다.
  형광펜을 지우면 "후보 아님", 색을 바꾸면 "등급 정정" 입니다.

[짝 맞추기의 원리]
  검토본은 마킹본을 고친 것이라 글자 내용은 원본과 같습니다. 그래서 문단 글자를
  이어붙여 만든 지문(fingerprint)으로 "어느 분석 기록과 비교할지" 를 찾습니다.
  【특허검토필요】 문구와 배지·요약 슬라이드는 도구가 넣은 것이므로 빼고 비교합니다.
"""
from __future__ import annotations

import colorsys
import datetime
import hashlib
import json
import re
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

from . import config, lexicon
from .extract import _walk_shapes

# 요약 슬라이드를 알아보는 표식 (mark.append_summary 가 붙이는 제목)
SUMMARY_TITLE = "특허 검토 마킹 요약"
# 배지 도형 이름 (mark.add_slide_badges 가 붙임)
BADGE_NAME = "특허검토필요 배지"
BADGE_RE = re.compile(rf"^{config.BADGE_TEXT}\s*\d+\s*건$")

# 색 계열 → 등급 판정용 기준 색상각(hue). GRADE_COLOR 에서 유도한 값.
#   A=FFD54F(노랑, 46°) · B=9FD8F5(하늘, 203°) · R=FF9E80(살구, 14°)
_HUE_ANCHOR = {"A": 46.0, "B": 203.0, "R": 14.0}


# ═════════════════════════════════════════════════════════════════
# 저장 위치
# ═════════════════════════════════════════════════════════════════
def _dir() -> Path:
    d = config.REVIEW_DIR
    (d / "analyses").mkdir(parents=True, exist_ok=True)
    return d


def _dataset_path() -> Path:
    return _dir() / "dataset.jsonl"


def _examples_path() -> Path:
    return _dir() / "examples.json"


def _rules_path() -> Path:
    return _dir() / "extra_rules.json"


# ═════════════════════════════════════════════════════════════════
# 1. 지문(fingerprint) ─ "같은 문서인가" 를 글자 내용으로 판별
# ═════════════════════════════════════════════════════════════════
def fingerprint_texts(texts: list[str]) -> str:
    joined = "\x1f".join(t for t in texts if t and t.strip())
    return hashlib.sha256(joined.encode("utf-8")).hexdigest()[:24]


def deck_fingerprint(deck) -> str:
    """extract.Deck 에서 지문을 만든다. 제목·본문·표만 쓴다 (노트·차트 제외)."""
    return fingerprint_texts(
        [s.text for s in deck.segments if s.kind in ("title", "body", "table")]
    )


# ═════════════════════════════════════════════════════════════════
# 2. 분석 기록 자동 저장 (main.py 가 분석 직후 호출)
# ═════════════════════════════════════════════════════════════════
def save_archive(deck, findings) -> str:
    """판정 결과를 지문 이름의 JSON 으로 남긴다. 나중에 검토본과 비교할 기준."""
    fp = deck_fingerprint(deck)
    data = {
        "fingerprint": fp,
        "date": datetime.date.today().isoformat(),
        "slide_count": deck.slide_count,
        "findings": [
            {
                "seg_id": f.seg_id, "slide_no": f.slide_no, "quote": f.quote,
                "grade": f.grade, "category": f.category, "implicit": f.implicit,
                "disclosure_risk": f.disclosure_risk, "reason": f.reason,
                "source": f.source, "span": list(f.span) if f.span else None,
            }
            for f in findings
        ],
        "segments": [
            {"seg_id": s.seg_id, "slide_no": s.slide_no, "kind": s.kind, "text": s.text}
            for s in deck.segments if s.kind in ("title", "body", "table")
        ],
    }
    (_dir() / "analyses" / f"{fp}.json").write_text(
        json.dumps(data, ensure_ascii=False), encoding="utf-8"
    )
    return fp


def load_archive(fp: str) -> dict | None:
    p = _dir() / "analyses" / f"{fp}.json"
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


# ═════════════════════════════════════════════════════════════════
# 3. 검토완료본 읽기 ─ 형광펜 색과 위치를 꺼낸다
# ═════════════════════════════════════════════════════════════════
def color_to_grade(hexv: str) -> tuple[str, bool]:
    """형광펜 색 → (등급, 공개리스크). 색 계열(hue)이 가장 가까운 기준을 따른다."""
    try:
        r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except (ValueError, TypeError):
        return "B", False
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if s < 0.15:                       # 회색·검정 등 무채색 → 기본 B(발굴 후보)
        return "B", False
    hue = h * 360
    best = min(_HUE_ANCHOR, key=lambda k: min(abs(hue - _HUE_ANCHOR[k]),
                                              360 - abs(hue - _HUE_ANCHOR[k])))
    if best == "R":
        return "B", True               # 리스크 색 = 등급 B + 공개 리스크 표시
    return best, False


def _para_review(p_el) -> tuple[str, list[dict], list[int]]:
    """문단 하나를 읽어 (도구 문구를 뺀 원문, 형광펜 구간 목록, 고아 태그 위치) 를 돌려준다.

    - 【특허검토필요】 런은 도구가 넣은 것이므로 글자에서 제외한다.
      (형광펜을 지웠는데 문구만 남아 있으면 그 자리가 '사람이 지운 흔적' 이 된다)
    - 같은 색이 이어지는 런들은 구간 하나로 합친다.
    """
    clean: list[str] = []
    spans: list[dict] = []
    tags: list[int] = []
    pos = 0
    cur: tuple[int, str] | None = None      # (시작 위치, 색)

    def _close():
        nonlocal cur
        if cur is not None and pos > cur[0]:
            spans.append({"start": cur[0], "end": pos, "color": cur[1]})
        cur = None

    for child in p_el:
        if etree.QName(child).localname not in ("r", "fld"):
            continue
        t = child.find(qn("a:t"))
        if t is None:
            continue
        text = t.text or ""
        if text.strip() == config.TAG_TEXT:      # 도구가 붙인 문구는 건너뜀
            _close()
            tags.append(pos)
            continue
        rPr = child.find(qn("a:rPr"))
        hl = None
        if rPr is not None:
            c = rPr.find(f"{qn('a:highlight')}/{qn('a:srgbClr')}")
            if c is not None:
                hl = (c.get("val") or "").upper()
        if hl:
            if cur is None:
                cur = (pos, hl)
            elif cur[1] != hl:
                _close()
                cur = (pos, hl)
        else:
            _close()
        clean.append(text)
        pos += len(text)
    _close()
    return "".join(clean), spans, tags


def read_reviewed(path: str) -> dict:
    """검토완료 PPTX 를 읽어 문단별 형광펜 정보와 지문을 돌려준다."""
    prs = Presentation(path)
    paras: list[dict] = []
    texts_for_fp: list[str] = []

    for s_idx, slide in enumerate(prs.slides, 1):
        # 요약 슬라이드(도구가 붙인 것)는 통째로 건너뛴다
        slide_paras: list[tuple] = []
        is_summary = False
        for shp, _pid in _walk_shapes(slide.shapes):
            if getattr(shp, "name", "") == BADGE_NAME:
                continue
            if getattr(shp, "has_chart", False) and shp.has_chart:
                continue
            frames = []
            if getattr(shp, "has_table", False) and shp.has_table:
                for row in shp.table.rows:
                    for cell in row.cells:
                        frames.append(cell.text_frame)
            elif getattr(shp, "has_text_frame", False) and shp.has_text_frame:
                frames.append(shp.text_frame)
            for tf in frames:
                for para in tf.paragraphs:
                    text, spans, tags = _para_review(para._p)
                    if not text.strip():
                        continue
                    if text.strip().startswith(SUMMARY_TITLE):
                        is_summary = True
                    if BADGE_RE.match(text.strip()):
                        continue
                    slide_paras.append((text, spans, tags))
        if is_summary:
            continue
        for text, spans, tags in slide_paras:
            texts_for_fp.append(text)
            paras.append({"slide_no": s_idx, "text": text,
                          "spans": spans, "tags": tags})

    return {"fingerprint": fingerprint_texts(texts_for_fp),
            "slide_count": len(prs.slides), "paras": paras}


# ═════════════════════════════════════════════════════════════════
# 4. 비교(diff) ─ 사람의 최종본과 도구의 판정을 맞대어 본다
# ═════════════════════════════════════════════════════════════════
def _snippet(text: str, start: int, end: int, pad: int = 55) -> str:
    lo, hi = max(0, start - pad), min(len(text), end + pad)
    out = text[lo:hi].replace("\n", " ")
    return ("…" if lo > 0 else "") + out + ("…" if hi < len(text) else "")


def _overlap(a: tuple[int, int], b: tuple[int, int]) -> bool:
    return not (a[1] <= b[0] or a[0] >= b[1])


def _lexicon_covers(text: str, span: tuple[int, int]) -> bool:
    """규칙 사전이 이 구간을 (일부라도) 이미 잡는가."""
    return any(_overlap(h.span, span) for h in lexicon.scan(text))


def diff(reviewed: dict) -> dict:
    """검토본과 분석 기록을 비교해 교정 내역을 만든다.

    분석 기록이 없으면(다른 PC 에서 분석했거나 지웠으면) '축소 모드' 로 동작:
      - 형광펜 구간 전부를 '확정 라벨' 로 수집하고
      - 고아 【특허검토필요】 문구(형광펜만 지워진 자리)로 오탐을 추정한다.
    """
    arch = load_archive(reviewed["fingerprint"])
    mode = "archive" if arch else "fallback"

    # 분석 기록의 문단별 판정 목록:  (slide, text) → [finding, ...]
    by_para: dict[tuple[int, str], list[dict]] = {}
    if arch:
        seg_text = {s["seg_id"]: (s["slide_no"], s["text"]) for s in arch["segments"]}
        for f in arch["findings"]:
            key = seg_text.get(f["seg_id"])
            if key and f.get("span"):
                by_para.setdefault(key, []).append(f)

    items: list[dict] = []
    n = {"match": 0, "miss": 0, "fp": 0, "regrade": 0, "gold": 0}

    for para in reviewed["paras"]:
        text = para["text"]
        human = []
        for sp in para["spans"]:
            grade, risk = color_to_grade(sp["color"])
            human.append({"span": (sp["start"], sp["end"]), "grade": grade,
                          "risk": risk, "quote": text[sp["start"]:sp["end"]].strip()})

        model = by_para.get((para["slide_no"], text), []) if arch else None

        if model is not None:
            used = set()
            for h in human:
                pair_i = next((i for i, f in enumerate(model)
                               if i not in used and _overlap(tuple(f["span"]), h["span"])), None)
                pair = model[pair_i] if pair_i is not None else None
                base = {"slide": para["slide_no"], "quote": h["quote"],
                        "grade": h["grade"], "risk": h["risk"],
                        "snippet": _snippet(text, *h["span"]), "text": text[:400]}
                if pair is None:
                    n["miss"] += 1
                    items.append({**base, "type": "miss",
                                  "lexicon_gap": not _lexicon_covers(text, h["span"])})
                else:
                    used.add(pair_i)
                    same = (pair["grade"] == h["grade"]
                            and bool(pair["disclosure_risk"]) == h["risk"])
                    if same:
                        n["match"] += 1
                        items.append({**base, "type": "match"})
                    else:
                        n["regrade"] += 1
                        items.append({**base, "type": "regrade",
                                      "from_grade": pair["grade"],
                                      "from_risk": bool(pair["disclosure_risk"])})
            for i, f in enumerate(model):
                if i in used:
                    continue
                n["fp"] += 1
                items.append({"type": "fp", "slide": para["slide_no"],
                              "quote": f["quote"], "grade": f["grade"],
                              "risk": bool(f["disclosure_risk"]),
                              "snippet": _snippet(text, *f["span"]), "text": text[:400]})
        else:
            # 축소 모드: 확정 라벨 + 고아 태그로 오탐 추정
            for h in human:
                n["gold"] += 1
                items.append({"type": "gold", "slide": para["slide_no"],
                              "quote": h["quote"], "grade": h["grade"], "risk": h["risk"],
                              "snippet": _snippet(text, *h["span"]), "text": text[:400],
                              "lexicon_gap": not _lexicon_covers(text, h["span"])})
            for tpos in para["tags"]:
                if any(abs(h["span"][1] - tpos) <= 2 for h in human):
                    continue                      # 형광펜이 그대로면 고아가 아님
                n["fp"] += 1
                items.append({"type": "fp", "slide": para["slide_no"],
                              "quote": text[max(0, tpos - 30):tpos].strip(),
                              "grade": "B", "risk": False,
                              "snippet": _snippet(text, max(0, tpos - 30), tpos),
                              "text": text[:400]})

    total = sum(n.values())
    agree = round(n["match"] / total * 100) if (mode == "archive" and total) else None
    return {"mode": mode, "fingerprint": reviewed["fingerprint"],
            "counts": n, "agree": agree, "items": items}


# ═════════════════════════════════════════════════════════════════
# 5. 저장 ─ 데이터셋에 쌓고, 프롬프트 예시를 갱신
# ═════════════════════════════════════════════════════════════════
def save_dataset(filename: str, mode: str, items: list[dict]) -> dict:
    today = datetime.date.today().isoformat()
    with _dataset_path().open("a", encoding="utf-8") as fh:
        for it in items:
            fh.write(json.dumps({"date": today, "file": filename, "mode": mode, **it},
                                ensure_ascii=False) + "\n")
    _curate_examples(items)
    return stats()


def _curate_examples(items: list[dict], keep: int = 12) -> None:
    """교정 사례를 프롬프트 주입용 예시 풀에 넣는다. 오탐·누락을 우선한다."""
    try:
        pool = json.loads(_examples_path().read_text(encoding="utf-8"))
    except Exception:
        pool = []
    for it in items:
        if it.get("type") not in ("fp", "miss", "gold", "regrade"):
            continue
        quote = (it.get("quote") or "").strip()
        if len(quote) < 2:
            continue
        ex = {"kind": "exclude" if it["type"] == "fp" else "include",
              "quote": quote[:80],
              "grade": it.get("grade", "B"), "risk": bool(it.get("risk")),
              "snippet": (it.get("snippet") or "")[:160]}
        pool = [p for p in pool if not (p["kind"] == ex["kind"] and p["quote"] == ex["quote"])]
        pool.insert(0, ex)
    # 반례(exclude)와 정답(include)을 절반씩 남긴다
    excl = [p for p in pool if p["kind"] == "exclude"][: keep // 2]
    incl = [p for p in pool if p["kind"] == "include"][: keep - len(excl)]
    _examples_path().write_text(json.dumps(excl + incl, ensure_ascii=False, indent=1),
                                encoding="utf-8")


_EX_CACHE: tuple[float, str] = (-1.0, "")


def examples_block() -> str:
    """analyze.py 가 시스템 프롬프트 뒤에 붙이는 '사내 검토 확정 사례' 블록."""
    global _EX_CACHE
    p = _examples_path()
    if not p.exists():
        return ""
    mtime = p.stat().st_mtime
    if mtime == _EX_CACHE[0]:
        return _EX_CACHE[1]
    try:
        pool = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return ""
    excl = [e for e in pool if e["kind"] == "exclude"][:4]
    incl = [e for e in pool if e["kind"] == "include"][:4]
    lines: list[str] = []
    if excl or incl:
        lines.append("## 사내 검토로 확정된 사례 (반드시 따를 것)")
    if excl:
        lines.append("다음과 같은 구간은 검토 결과 특허 후보가 아니었다. 유사한 문장을 반환하지 마라.")
        lines += [f'- "{e["quote"]}"' for e in excl]
    if incl:
        lines.append("다음과 같은 구간은 검토 결과 후보가 맞았다. 유사한 표현을 놓치지 마라.")
        lines += [f'- [{e["grade"]}{"·⚠공개" if e.get("risk") else ""}] "{e["quote"]}"'
                  for e in incl]
    block = "\n".join(lines)
    _EX_CACHE = (mtime, block)
    return block


# ═════════════════════════════════════════════════════════════════
# 6. [사전에 추가] ─ 놓친 표현을 규칙 사전에 등록
# ═════════════════════════════════════════════════════════════════
def add_rule(keyword: str) -> dict:
    keyword = (keyword or "").strip()
    if len(keyword) < 2:
        raise ValueError("두 글자 이상의 표현을 입력하세요.")
    try:
        rules = json.loads(_rules_path().read_text(encoding="utf-8"))
    except Exception:
        rules = []
    if any(r.get("keyword") == keyword for r in rules):
        return stats()
    rules.append({"keyword": keyword, "date": datetime.date.today().isoformat()})
    _rules_path().write_text(json.dumps(rules, ensure_ascii=False, indent=1),
                             encoding="utf-8")
    lexicon.load_user_rules()          # 다음 분석부터 바로 적용
    return stats()


def stats() -> dict:
    d = _dir()
    n_data = 0
    if _dataset_path().exists():
        n_data = sum(1 for _ in _dataset_path().open(encoding="utf-8"))
    n_ex = 0
    if _examples_path().exists():
        try:
            n_ex = len(json.loads(_examples_path().read_text(encoding="utf-8")))
        except Exception:
            pass
    n_rules = 0
    if _rules_path().exists():
        try:
            n_rules = len(json.loads(_rules_path().read_text(encoding="utf-8")))
        except Exception:
            pass
    n_arch = len(list((d / "analyses").glob("*.json")))
    return {"dataset": n_data, "examples": n_ex, "rules": n_rules, "archives": n_arch}
