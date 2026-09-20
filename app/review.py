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

  [반영 저장] 을 누르면 네 곳에 자동으로 쌓입니다.
    1. dataset.jsonl    — 교정 기록 원본 (나중에 미세조정 재료)
    2. examples.json    — 모범답안·반례가 계속 쌓이는 풀. 이 중 잘 고른 8건만
                          판정 프롬프트에 주입됩니다 (analyze.py 가 읽음).
    3. extra_rules.json — [사전에 추가] 로 등록한 표현 (lexicon.py 가 읽음)
    4. change_log.jsonl — 이번 반영으로 "무엇이 어디서 바뀌었는지" 변경 일지.
                          저장 직후 화면에 리포트로 뜨고, [성능 기록] 탭의
                          개선 타임라인에서 측정 결과와 함께 계속 볼 수 있습니다.

[검토자의 형광펜 색 규약]
  노랑 계열 = A (구체 수단 드러남) · 하늘/파랑 계열 = B (묵시) · 살구/빨강 계열 = C (공개 관련정보)
  PowerPoint 기본 형광펜 색을 써도 됩니다 — 정확한 색이 아니라 색 계열(색상환 거리)로 판정합니다.
  형광펜을 지우면 "후보 아님", 색을 바꾸면 "등급 정정" 입니다.

[짝 맞추기의 원리]
  검토본은 마킹본을 고친 것이라 글자 내용은 원본과 같습니다. 그래서 문단 글자를
  이어붙여 만든 지문(fingerprint)으로 "어느 분석 기록과 비교할지" 를 찾습니다.
  범례 상자와 (옵션으로 붙인) 【출원검토필요】 문구는 도구가 넣은 것이므로 빼고 비교합니다.
  예전 버전이 붙이던 배지·요약 슬라이드·【특허검토필요】 문구도 같은 방식으로 뺍니다
  (지금은 붙이지 않지만, 그때 만든 마킹본을 올려도 읽을 수 있어야 하므로 남겨 둡니다).
"""
from __future__ import annotations

import colorsys
import datetime
import hashlib
import json
import os
import re
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

from . import config, lexicon
from .extract import _walk_shapes

# 도구가 넣은 것들을 알아보는 표식.
# 지금 붙이는 것은 범례 상자뿐이지만, 예전 버전이 만든 마킹본(배지·요약 슬라이드가 있는
# 파일)을 올려도 그대로 읽히도록 옛 표식을 함께 둔다.
SUMMARY_TITLES = ("출원 검토 마킹 요약", "특허 검토 마킹 요약")     # 예전 요약 슬라이드 제목
TOOL_SHAPE_NAMES = frozenset(
    (config.LEGEND_NAME, *config.LEGACY_BADGE_NAMES)
)                                                                     # 범례·(예전)배지 도형 이름
TAG_TEXTS = frozenset((config.TAG_TEXT, *config.LEGACY_TAG_TEXTS))    # 구간 뒤 문구
BADGE_RE = re.compile(                                                # 예전 배지 글귀 "○○ N건"
    "^(?:" + "|".join(re.escape(t) for t in config.LEGACY_BADGE_TEXTS) + r")\s*\d+\s*건$"
)

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


def _changelog_path() -> Path:
    return _dir() / "change_log.jsonl"


def _write_json(path: Path, data) -> None:
    """JSON 을 '통째로 바꿔치기' 하는 방식으로 저장한다.

    그냥 write_text 로 덮어쓰면, 쓰는 도중에 전원이 꺼지거나 프로그램이 죽었을 때
    파일이 반쯤 쓰인 상태로 남아 다음 실행에서 읽지 못하게 됩니다.
    임시 파일에 다 쓴 뒤 한 번에 이름을 바꾸면 그런 중간 상태가 생기지 않습니다.
    (os.replace 는 Windows 에서도 같은 폴더 안이면 원자적으로 동작합니다)
    """
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, path)


def _quarantine(path: Path) -> Path | None:
    """읽지 못하는 파일을 지우지 말고 옆으로 치워 둔다. 옮긴 경로를 돌려준다."""
    if not path.exists():
        return None
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    broken = path.with_name(f"{path.stem}.broken-{stamp}{path.suffix}")
    try:
        os.replace(path, broken)
        return broken
    except Exception:      # noqa: BLE001  (치우기에 실패해도 저장은 계속한다)
        return None


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
    """형광펜 색 → (등급, 공개 관련정보인가). 색 계열(hue)이 가장 가까운 기준을 따른다."""
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
        return "B", True               # 살구/빨강 = C(공개 관련정보) 색
    return best, False


def _para_review(p_el) -> tuple[str, list[dict], list[int]]:
    """문단 하나를 읽어 (도구 문구를 뺀 원문, 형광펜 구간 목록, 고아 태그 위치) 를 돌려준다.

    - 【출원검토필요】(예전 버전은 【특허검토필요】) 런은 도구가 넣은 것이므로 글자에서 제외한다.
      (형광펜을 지웠는데 문구만 남아 있으면 그 자리가 '사람이 지운 흔적' 이 된다.
       문구 옵션을 끄고 만든 마킹본에는 이 런이 아예 없다 — 그래도 동작에는 지장이 없다)
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
        if text.strip() in TAG_TEXTS:            # 도구가 붙인 문구는 건너뜀
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
        # 예전 버전이 붙이던 요약 슬라이드는 통째로 건너뛴다 (지금은 붙이지 않음)
        slide_paras: list[tuple] = []
        is_summary = False
        for shp, _pid in _walk_shapes(slide.shapes):
            if getattr(shp, "name", "") in TOOL_SHAPE_NAMES:    # 범례 상자·(예전)배지
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
                    if text.strip().startswith(SUMMARY_TITLES):
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


def _lexicon_probe(text: str, span: tuple[int, int]) -> tuple[bool, str | None]:
    """구간과 겹치는 규칙 사전 신호를 한 번만 훑어 두 가지를 돌려준다.

      (사전이 이 구간을 이미 잡는가, 대표 카테고리)

    카테고리는 예시를 고를 때 '종류가 겹치지 않게' 뽑는 기준이 된다.
    모델이 놓친 구간(miss)에는 판정 기록이 없으므로 여기서 유추한다.
    겹치는 신호가 없으면 (False, None) ─ 사전에도 없는 새 표현이라는 뜻이다.
    """
    hits = [h for h in lexicon.scan(text) if _overlap(h.span, span)]
    if not hits:
        return False, None
    return True, max(hits, key=lambda h: h.weight).category


def diff(reviewed: dict) -> dict:
    """검토본과 분석 기록을 비교해 교정 내역을 만든다.

    분석 기록이 없으면(다른 PC 에서 분석했거나 지웠으면) '축소 모드' 로 동작:
      - 형광펜 구간 전부를 '확정 라벨' 로 수집하고
      - 고아 【출원검토필요】 문구(형광펜만 지워진 자리)로 오탐을 추정한다.
        (문구 옵션을 끄고 만든 마킹본이면 이 단서가 없으므로 확정 라벨만 모인다)
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
                    covered, cat = _lexicon_probe(text, h["span"])
                    n["miss"] += 1
                    items.append({**base, "type": "miss", "category": cat,
                                  "lexicon_gap": not covered})
                else:
                    used.add(pair_i)
                    same = (pair["grade"] == h["grade"]
                            and bool(pair["disclosure_risk"]) == h["risk"])
                    if same:
                        n["match"] += 1
                        items.append({**base, "type": "match",
                                      "category": pair.get("category")})
                    else:
                        n["regrade"] += 1
                        items.append({**base, "type": "regrade",
                                      "category": pair.get("category"),
                                      "from_grade": pair["grade"],
                                      "from_risk": bool(pair["disclosure_risk"])})
            for i, f in enumerate(model):
                if i in used:
                    continue
                n["fp"] += 1
                items.append({"type": "fp", "slide": para["slide_no"],
                              "quote": f["quote"], "grade": f["grade"],
                              "risk": bool(f["disclosure_risk"]),
                              "category": f.get("category"),
                              "snippet": _snippet(text, *f["span"]), "text": text[:400]})
        else:
            # 축소 모드: 확정 라벨 + 고아 태그로 오탐 추정
            for h in human:
                covered, cat = _lexicon_probe(text, h["span"])
                n["gold"] += 1
                items.append({"type": "gold", "slide": para["slide_no"],
                              "quote": h["quote"], "grade": h["grade"], "risk": h["risk"],
                              "category": cat,
                              "snippet": _snippet(text, *h["span"]), "text": text[:400],
                              "lexicon_gap": not covered})
            for tpos in para["tags"]:
                if any(abs(h["span"][1] - tpos) <= 2 for h in human):
                    continue                      # 형광펜이 그대로면 고아가 아님
                span = (max(0, tpos - 30), tpos)
                n["fp"] += 1
                items.append({"type": "fp", "slide": para["slide_no"],
                              "quote": text[span[0]:span[1]].strip(),
                              "grade": "B", "risk": False,
                              "category": _lexicon_probe(text, span)[1],
                              "snippet": _snippet(text, *span),
                              "text": text[:400]})

    total = sum(n.values())
    agree = round(n["match"] / total * 100) if (mode == "archive" and total) else None
    return {"mode": mode, "fingerprint": reviewed["fingerprint"],
            "counts": n, "agree": agree, "items": items}


def preview_many(named_paths: list[tuple[str, str]]) -> list[dict]:
    """검토완료본 여러 개를 한 번에 읽어 파일별 교정 내역을 만든다.

    - 파일 하나가 깨져 있어도 전체를 실패시키지 않는다: 그 파일에만 error 를 채운다.
    - 같은 문서(지문 동일)의 검토본이 두 개 들어오면 첫 번째만 처리하고
      두 번째는 duplicate_of 로 표시한다. 두 벌의 교정이 서로 다를 수 있는데
      어느 쪽이 옳은지 기계가 고를 수 없으므로, 한 번에 한 벌만 반영한다.
      (두 번째 벌을 반영하려면 따로 다시 올리면 된다 — 나중 기록이 우선된다)
    """
    out: list[dict] = []
    seen: dict[str, str] = {}          # 지문 → 먼저 온 파일 이름
    for name, path in named_paths:
        if not name.lower().endswith(".pptx"):
            out.append({"filename": name, "error": "PPTX 파일만 지원합니다."})
            continue
        try:
            result = diff(read_reviewed(path))
        except Exception as e:  # noqa: BLE001  (한 파일의 오류가 배치 전체를 막지 않도록)
            out.append({"filename": name, "error": f"{type(e).__name__}: {e}"})
            continue
        first = seen.get(result["fingerprint"])
        if first:
            out.append({"filename": name, "duplicate_of": first})
            continue
        seen[result["fingerprint"]] = name
        result["filename"] = name
        out.append(result)
    return out


# ═════════════════════════════════════════════════════════════════
# 5. 저장 ─ 데이터셋에 쌓고, 프롬프트 예시를 갱신
# ═════════════════════════════════════════════════════════════════
# 프롬프트에 넣을 예시를 고르는 기준. 숫자를 키우면 판정 근거가 늘지만
# 그만큼 본문에 쓸 수 있는 분량이 줄어듭니다 (docs/01 참고).
#
# 반례와 정답에 칸을 다르게 주는 이유 — 쌓이는 카테고리 폭이 서로 다릅니다.
#   · 오탐(반례)은 경영지표 계열에 몰립니다(효과만기재·비교우위·수치). 4칸이면 대개 포화이고,
#     더 주면 같은 카테고리가 반복돼 칸만 낭비됩니다.
#   · 누락(정답)은 구성·제어·공정·독자성 등 넓게 퍼집니다. 8칸까지 새 카테고리가 붙습니다.
# 실제 데이터가 이와 다르면 아래 숫자를 조정하세요 (회차 비교가 끊기므로 자주 바꾸지는 말 것).
EXAMPLE_INJECT_EXCLUDE = 4  # 반례(오탐) 칸 수
EXAMPLE_INJECT_INCLUDE = 8  # 정답(누락) 칸 수
EXAMPLE_WINDOW_DAYS = 180   # 이 기간 안의 교정을 먼저 고른다 (모자라면 그 이전 것도)
EXAMPLE_PER_BATCH = 2       # 한 번의 [반영 저장] 이 차지할 수 있는 최대 칸 수


def _inject_n(kind: str) -> int:
    """종류별 주입 칸 수. 상수를 호출 시점에 읽으므로 값을 바꾸면 바로 반영됩니다."""
    return EXAMPLE_INJECT_EXCLUDE if kind == "exclude" else EXAMPLE_INJECT_INCLUDE


def _next_batch() -> int:
    """다음 [반영 저장] 에 찍을 번호. 파일 여러 개를 한 번에 올려도 하나로 묶으려고,
    commit_with_log 가 저장 전에 한 번만 계산해 모든 파일에 같은 값을 넘깁니다."""
    pool = _examples_pool()
    return max((p.get("batch") or 0) for p in pool) + 1 if pool else 1


def save_dataset(filename: str, mode: str, items: list[dict],
                 batch: int | None = None) -> dict:
    today = datetime.date.today().isoformat()
    with _dataset_path().open("a", encoding="utf-8") as fh:
        for it in items:
            fh.write(json.dumps({"date": today, "file": filename, "mode": mode, **it},
                                ensure_ascii=False) + "\n")
    _curate_examples(items, batch=batch)
    return stats()


def _curate_examples(items: list[dict], keep: int | None = None,
                     batch: int | None = None) -> None:
    """교정 사례를 프롬프트 주입용 예시 풀에 쌓는다.

    풀은 **버리지 않고 계속 쌓입니다** (keep 을 주면 그만큼만 남깁니다).
    실제로 프롬프트에 들어가는 것은 _pick() 이 고른 몇 건뿐이므로
    (기본: 반례 4 + 정답 8), 풀이 커져도 분석 비용은 늘지 않습니다.
    대신 나중에 '좋은 것' 을 고를 재료가 남고, 회사 PC 에서 교정 이력을
    통째로 가지고 나올 수 있습니다.

    한 번의 [반영 저장] 으로 들어온 항목에는 같은 batch 번호를 찍습니다
    (파일을 여러 개 올려도 하나로 셉니다 — commit_with_log 가 번호를 넘겨 줍니다).
    한 번의 반영이 주입 칸을 독점하지 못하게 막는 데 쓰입니다.

    같은 표현이 이미 풀에 있으면 **반례·정답을 가리지 않고** 지우고 새로 넣습니다.
    한 표현을 "빼라" 와 "잡아라" 로 동시에 가르치면 프롬프트가 서로 모순되므로,
    검토에서 늘 그렇듯 나중 기록이 이깁니다.
    """
    pool, ok = _read_examples()
    if not ok:                                   # 파일이 깨져 못 읽은 경우
        moved = _quarantine(_examples_path())    # 지우지 않고 옆으로 치운다
        _append_log({"ts": datetime.datetime.now().isoformat(timespec="seconds"),
                     "kind": "examples_recovered",
                     "moved_to": moved.name if moved else None})
    today = datetime.date.today().isoformat()
    if batch is None:
        batch = max((p.get("batch") or 0) for p in pool) + 1 if pool else 1
    for it in items:
        itype = it.get("type")
        if itype not in ("fp", "miss", "gold", "regrade"):
            continue
        quote = (it.get("quote") or "").strip()
        if len(quote) < 2:
            continue
        ex = {"kind": "exclude" if itype == "fp" else "include",
              "type": itype,
              "quote": quote[:80],
              "grade": it.get("grade", "B"), "risk": bool(it.get("risk")),
              "category": it.get("category"),
              "date": today, "batch": batch,
              "snippet": (it.get("snippet") or "")[:160]}
        pool = [p for p in pool if p.get("quote") != ex["quote"]]
        pool.insert(0, ex)                       # 최신이 앞
    if keep is not None:
        excl = [p for p in pool if p.get("kind") == "exclude"][: keep // 2]
        incl = [p for p in pool if p.get("kind") == "include"][: keep - len(excl)]
        pool = excl + incl
    _write_json(_examples_path(), pool)


def _pick(pool: list[dict], kind: str, n: int | None = None) -> list[dict]:
    """쌓인 풀에서 프롬프트에 넣을 예시를 고른다 (n 을 안 주면 종류별 기본 칸 수).

    그냥 최신순으로 자르면 문제가 셋 있었습니다.
      · 오탐이 몰린 보고서 하나가 칸을 전부 차지함
      · 카테고리가 11종인데 같은 종류만 뽑힘
      · 축소 모드의 '확정 라벨'(gold, 모델이 이미 잘 잡던 것)이 정답 칸을 낭비함

    먼저 **확증된 교정**(miss·regrade·fp)으로 칸을 채우고, 모자랄 때만
    축소 모드의 확정 라벨(gold)로 채웁니다. gold 는 모델이 실제로 놓쳤다는
    확증이 없어서 — 이미 잘 잡던 것일 수도 있어서 — 정보량이 적습니다.

    각 단계 안에서는 조건을 걸고 고르되, 재료가 모자라면 하나씩 풀어 갑니다.
    (조건을 다 풀면 예전과 같은 '최신순'이 되므로 칸이 비는 일은 없습니다.)

        0단계  최근 것 · 카테고리 안 겹침 · 한 반영당 2칸까지
        1단계  └ 기간 제한을 푼다 (오래된 것도 후보로)
        2단계  └ 반영당 칸 수 제한을 푼다
        3단계  └ 카테고리 제한을 푼다  = 최신순
    """
    if n is None:
        n = _inject_n(kind)
    cand = [e for e in pool if e.get("kind") == kind]
    if not cand:
        return []

    cutoff = (datetime.date.today()
              - datetime.timedelta(days=EXAMPLE_WINDOW_DAYS)).isoformat()
    picked: list[dict] = []
    seen_cat: set[str] = set()
    per_batch: dict[object, int] = {}

    def take(src: list[dict], relax: int) -> None:
        for e in src:                             # src 는 최신이 앞
            if len(picked) >= n:
                return
            if any(e is p for p in picked):
                continue
            if relax < 1 and (e.get("date") or "") < cutoff:
                continue
            if relax < 2 and per_batch.get(e.get("batch"), 0) >= EXAMPLE_PER_BATCH:
                continue
            # 카테고리가 없는 것(= 사전에 아예 없던 새 표현, 가장 값진 교정)을
            # 하나의 '미분류' 로 묶으면 그런 예시가 몇 건 쌓이든 한 칸밖에 못 들어간다.
            # 그래서 표현별로 다른 칸으로 세되, 서로 겹치지는 않게 표현을 키로 쓴다.
            cat = e.get("category") or f"미분류:{e.get('quote')}"
            if relax < 3 and cat in seen_cat:
                continue
            picked.append(e)
            seen_cat.add(cat)
            per_batch[e.get("batch")] = per_batch.get(e.get("batch"), 0) + 1

    confirmed = [e for e in cand if e.get("type") != "gold"]
    unconfirmed = [e for e in cand if e.get("type") == "gold"]
    for src in (confirmed, unconfirmed):
        for relax in range(4):
            take(src, relax)
            if len(picked) >= n:
                break
        if len(picked) >= n:
            break
    return picked


# ═════════════════════════════════════════════════════════════════
# 5.5 변경 일지 ─ 반영 전후로 무엇이 어디서 바뀌었는지 기록
# ═════════════════════════════════════════════════════════════════
def _read_examples() -> tuple[list[dict], bool]:
    """(풀, 제대로 읽었는가). 파일이 아직 없으면 ([], True) — 처음 쓰는 상태입니다.

    '파일 없음' 과 '파일이 깨져서 못 읽음' 을 구분하는 게 요점입니다. 둘을 똑같이
    빈 풀로 취급하면, 한 번 깨진 examples.json 위에 다음 [반영 저장] 이 빈 풀을
    덮어써서 그동안 쌓은 교정 이력이 통째로 사라집니다.
    """
    p = _examples_path()
    if not p.exists():
        return [], True
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:      # noqa: BLE001  (깨진 JSON·인코딩 오류 모두)
        return [], False
    return (data, True) if isinstance(data, list) else ([], False)


def _examples_pool() -> list[dict]:
    return _read_examples()[0]


def _injected(pool: list[dict]) -> list[dict]:
    """프롬프트에 실제로 주입되는 부분 (examples_block 과 같은 선별)."""
    return [{"kind": e["kind"], "quote": e["quote"],
             "grade": e.get("grade"), "risk": bool(e.get("risk")),
             "category": e.get("category")}
            for e in _pick(pool, "exclude") + _pick(pool, "include")]


def _append_log(entry: dict) -> None:
    _dir()
    with _changelog_path().open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(entry, ensure_ascii=False) + "\n")


def change_log() -> list[dict]:
    """쌓인 변경 일지 전체. [성능 기록] 탭의 개선 타임라인이 읽는다."""
    p = _changelog_path()
    if not p.exists():
        return []
    out: list[dict] = []
    for line in p.open(encoding="utf-8"):
        try:
            out.append(json.loads(line))
        except Exception:
            continue
    return out


def commit_with_log(batches: list[dict]) -> dict:
    """[반영 저장] 의 진입점 — 저장하고, 반영 전후의 변화를 일지에 남긴다.

    batches: [{filename, mode, items}, ...]  (파일 여러 개를 한 번에)
    돌려주는 값의 change 가 "이번 반영으로 바뀐 것" 리포트가 된다:
      · dataset 몇 건 → 몇 건 (유형별 추가 내역)
      · 프롬프트 주입 예시에 새로 들어온 것 / 밀려난 것
      · 사전 표현 개수 전후
    """
    before = stats()
    inj_before = _injected(_examples_pool())

    per_file: list[dict] = []
    added = {"miss": 0, "fp": 0, "regrade": 0, "match": 0, "gold": 0}
    # 이번 반영 전체가 한 batch. 파일마다 번호를 새로 따면 파일을 여러 개 올린
    # 반영 한 번이 EXAMPLE_PER_BATCH 제한을 우회해 주입 칸을 모두 차지한다.
    commit_batch = _next_batch()
    for b in batches:
        items = b.get("items") or []
        if not items:
            continue
        save_dataset(b.get("filename") or "unknown.pptx",
                     b.get("mode") or "archive", items, batch=commit_batch)
        c = {k: sum(1 for it in items if it.get("type") == k) for k in added}
        for k in added:
            added[k] += c[k]
        per_file.append({"filename": b.get("filename") or "unknown.pptx",
                         "mode": b.get("mode") or "archive",
                         "items": len(items), **c})
    if not per_file:
        raise ValueError("저장할 항목이 없습니다.")

    after = stats()
    inj_after = _injected(_examples_pool())

    def _k(e: dict) -> tuple:
        return (e["kind"], e["quote"])

    bkeys = {_k(e) for e in inj_before}
    akeys = {_k(e) for e in inj_after}
    entry = {
        "ts": datetime.datetime.now().isoformat(timespec="seconds"),
        "kind": "review_commit",
        "files": per_file,
        "added": {**added, "total": sum(added.values())},
        "dataset": {"before": before["dataset"], "after": after["dataset"]},
        "examples": {
            "before": before["examples"], "after": after["examples"],
            "entered": [e for e in inj_after if _k(e) not in bkeys],
            "left": [e for e in inj_before if _k(e) not in akeys],
        },
        "rules": {"before": before["rules"], "after": after["rules"]},
    }
    _append_log(entry)
    return {"change": entry, "stats": after,
            "saved_files": len(per_file),
            "saved_items": sum(f["items"] for f in per_file)}


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
    excl = _pick(pool, "exclude")
    incl = _pick(pool, "include")
    lines: list[str] = []
    if excl or incl:
        lines.append("## 사내 검토로 확정된 사례 (반드시 따를 것)")
    if excl:
        lines.append("다음과 같은 구간은 검토 결과 특허 후보가 아니었다. 유사한 문장을 반환하지 마라.")
        lines += [f'- "{e["quote"]}"' for e in excl]
    if incl:
        lines.append("다음과 같은 구간은 검토 결과 후보가 맞았다. 유사한 표현을 놓치지 마라.")
        lines += [f'- [{e["grade"]}{"·C공개" if e.get("risk") else ""}] "{e["quote"]}"'
                  for e in incl]
    block = "\n".join(lines)
    _EX_CACHE = (mtime, block)
    return block


# ═════════════════════════════════════════════════════════════════
# 6. [사전에 추가] ─ 놓친 표현을 규칙 사전에 등록
# ═════════════════════════════════════════════════════════════════
def _read_rules() -> tuple[list[dict], bool]:
    """(사전, 제대로 읽었는가). examples 와 같은 이유로 '없음' 과 '깨짐' 을 구분합니다."""
    p = _rules_path()
    if not p.exists():
        return [], True
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception:      # noqa: BLE001
        return [], False
    return (data, True) if isinstance(data, list) else ([], False)


def add_rule(keyword: str) -> dict:
    keyword = (keyword or "").strip()
    if len(keyword) < 2:
        raise ValueError("두 글자 이상의 표현을 입력하세요.")
    rules, ok = _read_rules()
    if not ok:                             # 깨진 사전을 빈 사전으로 덮어쓰지 않는다
        moved = _quarantine(_rules_path())
        _append_log({"ts": datetime.datetime.now().isoformat(timespec="seconds"),
                     "kind": "rules_recovered",
                     "moved_to": moved.name if moved else None})
    if any(r.get("keyword") == keyword for r in rules):
        return stats()
    rules.append({"keyword": keyword, "date": datetime.date.today().isoformat()})
    _write_json(_rules_path(), rules)
    lexicon.load_user_rules()          # 다음 분석부터 바로 적용
    _append_log({"ts": datetime.datetime.now().isoformat(timespec="seconds"),
                 "kind": "rule_add", "keyword": keyword,
                 "rules_after": len(rules)})
    return stats()


def stats() -> dict:
    d = _dir()
    n_data = 0
    if _dataset_path().exists():
        n_data = sum(1 for _ in _dataset_path().open(encoding="utf-8"))
    # 예시는 '쌓인 것(pool)' 과 '실제로 프롬프트에 들어가는 것(injected)' 이 다르다.
    # 풀은 계속 커지지만 주입 칸 수는 상수로 고정이라 분석 비용은 일정하다.
    pool = _examples_pool()
    n_ex = len(pool)
    n_inj = len(_pick(pool, "exclude")) + len(_pick(pool, "include"))
    n_rules = 0
    if _rules_path().exists():
        try:
            n_rules = len(json.loads(_rules_path().read_text(encoding="utf-8")))
        except Exception:
            pass
    n_arch = len(list((d / "analyses").glob("*.json")))
    return {"dataset": n_data, "examples": n_ex, "injected": n_inj,
            "rules": n_rules, "archives": n_arch}
