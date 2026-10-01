"""
merge.py ─ 규칙 사전 결과와 모델 결과를 합쳐 "최종 목록" 을 만드는 3단계
=====================================================================

[이 파일이 하는 일]
  1) 모델이 돌려준 인용구(quote)를 원문 안의 정확한 글자 위치(span)로 바꾼다.
  2) 모델이 매출·이익률 같은 경영 지표를 기술 후보로 올린 오탐은 C 등급으로 내린다.
  3) 모델이 놓쳤지만 규칙 사전 점수가 높은 문단은 B 등급으로 구제한다(안전망).
  4) 같은 문단 안에서 겹치는 구간을 하나로 합쳐 "칠할 구간(Mark)" 목록을 만든다.

      Finding(모델) + Hit(규칙) ──▶ resolve() ──▶ (최종 Finding 목록, Mark 목록)

[초보자를 위한 설명 ─ span 이란?]
  "문단의 몇 번째 글자부터 몇 번째 글자까지" 를 뜻하는 (시작, 끝) 숫자 쌍입니다.
  예) text = "조명 입사각을 22~28° 구간에서", quote = "22~28°" 이면 span = (8, 14).
  mark.py 는 이 숫자로 런을 잘라 형광펜을 칠합니다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from . import config, lexicon
from .analyze import Finding
from .extract import Deck, Segment

# 등급 비교용 순위. 겹치는 구간은 더 높은 등급을 따릅니다.
GRADE_RANK = {"C": 0, "B": 1, "A": 2}

# 규칙 사전이 강하게 걸렸는데 모델이 아무 말도 하지 않은 문단을 구제하는 기준 점수 ('빠짐없이' 에서만).
# 온디바이스 소형 모델은 묵시적 표현을 조용히 넘기는 일이 잦아서 이 안전망이 필요합니다.
# 값을 올리면 구제가 줄고(정밀도↑), 내리면 구제가 늘어납니다(재현율↑).
RESCUE_SCORE = 8

# '확실한 것만'(strict) 에서 모델이 놓쳐도 살리는 공개 신호의 최소 가중치 — 전시·시연(8)·논문·학회(8)·보도자료(7) 만.
# 출시·양산(6)·제안서 제출(6) 은 예정·계획 문장이 많아 strict 에서는 모델이 직접 골랐을 때만 남긴다.
STRICT_RESCUE_WEIGHT = 7

_WS = re.compile(r"\s+")
_SENT_END = re.compile(r"[.!?。\n·]")


def _readable_span(text: str, span: tuple[int, int],
                   whole_below: int = 70) -> tuple[int, int]:
    """규칙 사전이 잡은 좁은 구간을 사람이 읽을 수 있는 "문장 단위" 로 넓힌다.

    'ms   →   11 ' 처럼 잘린 조각이 그대로 인용구가 되면 검토자가 뜻을 알 수 없습니다.
    문단이 짧으면(70자 이하) 통째로, 길면 구간을 감싸는 문장까지.
    """
    if len(text.strip()) <= whole_below:
        return (0, len(text))
    s, e = span
    left = 0
    for m in _SENT_END.finditer(text, 0, s):
        left = m.end()
    right = len(text)
    m = _SENT_END.search(text, e)
    if m:
        right = m.start() + 1
    if right - left > 160:  # 너무 길어지면 원래 구간 주변만
        left, right = max(0, s - 20), min(len(text), e + 20)
    return (left, right)


def _find_span(text: str, quote: str) -> tuple[int, int] | None:
    """모델이 돌려준 인용구를 원문에서 찾아 (시작, 끝) 위치로 바꾼다.

    모델이 공백을 조금 다르게 적는 일이 흔해서, 1차로 그대로 찾고
    2차로 "공백을 전부 뺀 상태" 로 다시 찾습니다. 그래도 없으면 None.
    """
    if not quote:
        return None
    i = text.find(quote)
    if i >= 0:
        return (i, i + len(quote))

    # 공백을 무시하고 다시 시도
    stripped = _WS.sub("", quote)
    if not stripped:
        return None
    idx_map: list[int] = []          # 공백 뺀 문자열의 i번째 글자 → 원문에서의 위치
    packed_chars: list[str] = []
    for pos, ch in enumerate(text):
        if not ch.isspace():
            packed_chars.append(ch)
            idx_map.append(pos)
    packed = "".join(packed_chars)
    j = packed.find(stripped)
    if j < 0:
        return None
    return (idx_map[j], idx_map[j + len(stripped) - 1] + 1)


@dataclass
class Mark:
    """실제로 색을 칠할 구간 하나. mark.py 로 넘어간다."""

    seg_id: int
    slide_no: int
    span: tuple[int, int] | None   # None 이면 문단 전체
    grade: str
    disclosure_risk: bool


def _locate_quote(text: str, quote: str) -> tuple[int, int] | None:
    """모델 인용구의 자리: 그대로 → 공백 무시 → 숫자·공백 변형과 조각 합(memory.locate) 순으로 찾는다.

    모델이 인용구를 한두 글자 고쳐 쓰는 일이 잦아(2026-09-21 측정 37건 중 5건), 못 찾으면 문단 전체가
    칠해져 "막 칠한" 것처럼 보였다. 검토 학습이 쓰는 너그러운 찾기를 여기에도 쓴다.
    """
    span = _find_span(text, quote)
    if span is not None:
        return span
    try:
        from . import memory                        # 호출 시점에 가져온다 (순환 import 방지)
        return memory.locate(text, quote)
    except Exception:  # noqa: BLE001
        return None


def resolve(deck: Deck, findings: list[Finding],
            hits_by_seg: dict[int, list[lexicon.Hit]], mode: str = "broad",
            disclosure: bool = True) -> tuple[list[Finding], list[Mark]]:
    """이 파일의 진입점. (최종 Finding 목록, Mark 목록) 을 돌려준다.

    mode : "broad" 빠짐없이 (2026-09-30b 까지의 동작) / "strict" 확실한 것만 —
           ① 인용구 자리를 못 찾은 후보는 문단 전체를 칠하지 않고 버린다
           ② 모델이 낸 B(수단이 안 적힌 후보)는 공개 관련이 아니면 버린다
           ③ 사전 점수 안전망은 쓰지 않고, 확실한 공개 신호(STRICT_RESCUE_WEIGHT 이상)만 살린다
           ④ config.STRICT_MIN_CHARS 보다 짧은 문단의 모델 후보는 버린다 (표 조각)
    disclosure : False 면 공개 관련 표시 끔 (2026-10-01b) — 모델이 낸 공개 후보는 A 면 공개 표시만 지우고 보통 후보로 두고,
           그 밖(공개 여부만 말하는 문장)은 버린다. 공개 규칙(RISK_*)은 prescreen 에서 이미 빠져 구제·공개 보강에도 쓰이지 않는다.
    """
    strict = mode == "strict"
    seg_map: dict[int, Segment] = {s.seg_id: s for s in deck.segments}

    # ── 1) 모델 결과의 인용구를 원문 좌표로 변환하고, 경영 정보 오탐을 걸러낸다 ──
    resolved: list[Finding] = []
    for f in findings:
        seg = seg_map.get(f.seg_id)
        if seg is None:
            continue
        if not disclosure and f.disclosure_risk and f.source == "llm":
            if f.grade != "A":
                continue                                   # 공개 여부만 말하는 문장 — 표시 끔
            f.disclosure_risk = False                      # 수단이 적힌 A 는 보통 후보로 남긴다
        if strict and f.source == "llm":
            if len(seg.text) < config.STRICT_MIN_CHARS:
                continue
            if f.grade == "B" and not f.disclosure_risk:
                continue
        f.span = _locate_quote(seg.text, f.quote)
        if f.span is None:
            if strict and f.source == "llm":
                continue                                   # 자리를 못 찾으면 확실하지 않은 것 — 버린다
            # 인용구를 찾지 못하면 문단 전체를 칠하고, 인용구는 원문으로 교체한다
            f.quote = seg.text
        # 모델이 "특허 출원 0건 · 검토 미착수" 같은 특허 행정 상태를 '공개' 로 돌려보내면 뺀다.
        # 규칙 사전의 진짜 공개 신호(전시·논문·출시…)가 같은 구간에 있으면 그대로 둔다.
        if (f.disclosure_risk and f.source == "llm"
                and lexicon.is_ip_status(seg.text, f.span)
                and not any(h.disclosure and not (h.span[1] <= f.span[0] or h.span[0] >= f.span[1])
                            for h in hits_by_seg.get(f.seg_id, []) if f.span)):
            continue
        # 재무·조직 지표에 "효과만 기재" 원칙을 확대 적용한 판정은 등급을 내린다.
        # 모델이 재현율을 높이라는 지시를 매출·이익률까지 밀고 나가는 경향이 있다.
        if not f.disclosure_risk and lexicon.is_business_noise(
            seg.text, hits_by_seg.get(f.seg_id, []), f.span
        ):
            f.grade = "C"
            f.reason = (f.reason + " / 경영 지표 문장으로 판단해 등급을 낮춤").strip()
        resolved.append(f)

    # ── 2) 규칙 사전이 강하게 걸렸는데 모델이 놓친 문단 구제 (안전망) ──
    covered = {f.seg_id for f in resolved}
    for seg_id, hits in hits_by_seg.items():
        if not hits or seg_id in covered:
            continue
        seg = seg_map.get(seg_id)
        if seg is None:
            continue
        # "논문·특허 등 공개 실적 없음" 처럼 뒤집힌 공개 신호는 공개로 세지 않는다
        hits = [h for h in hits if not (h.disclosure and lexicon.negated(seg.text, h.span))]
        if not hits:
            continue
        if strict:
            # 확실한 공개 신호만 살린다 (전시·시연 / 논문·학회 / 보도자료). 점수 안전망은 쓰지 않는다
            hits = [h for h in hits if h.disclosure and h.weight >= STRICT_RESCUE_WEIGHT]
            if not hits:
                continue
            risky = True
        else:
            sc = lexicon.score(seg.text, hits)
            risky = lexicon.has_disclosure_risk(hits)
            if sc < RESCUE_SCORE and not risky:
                continue
            if lexicon.is_business_noise(seg.text, hits) and not risky:
                continue
        top = max(hits, key=lambda h: h.weight)      # 가장 가중치 높은 규칙을 대표로
        wide = _readable_span(seg.text, top.span)
        resolved.append(
            Finding(
                seg_id=seg_id,
                slide_no=seg.slide_no,
                quote=seg.text[wide[0]:wide[1]].strip(),
                grade="B",
                category=top.category,
                implicit=top.category in config.IMPLICIT_CATEGORIES,   # 모델 답과 같은 기준
                disclosure_risk=risky,
                reason=f"{top.hint} (규칙 사전 감지 · 모델 미판정 구간)",
                source="lexicon",
                span=wide,
                rule_hints=[h.rid for h in hits],
            )
        )

    # ── 3) 규칙 사전의 공개 신호로 disclosure_risk 보강 (뒤집힌 신호는 제외) ──
    for f in resolved:
        if not f.disclosure_risk and lexicon.has_disclosure_risk(hits_by_seg.get(f.seg_id, [])):
            seg = seg_map.get(f.seg_id)
            if seg and f.span:
                for h in hits_by_seg[f.seg_id]:
                    if h.category == "공개이력" and h.disclosure and not lexicon.negated(seg.text, h.span) and not (
                        h.span[1] <= f.span[0] or h.span[0] >= f.span[1]
                    ):
                        f.disclosure_risk = True
                        break

    # 모델이 스스로 매긴 C(배경·참고 수준)는 출원 후보가 아니므로 결과에서 뺀다.
    # 단 공개 신호가 붙었으면 남긴다 — 그 구간은 C(공개 관련정보)로 표시된다.
    resolved = [f for f in resolved if f.grade != "C" or f.disclosure_risk]

    return sort_findings(resolved), build_marks(deck, resolved)


def sort_findings(resolved: list[Finding]) -> list[Finding]:
    """슬라이드 → 문단 → 위치 순으로 정렬."""
    resolved.sort(key=lambda f: (f.slide_no, f.seg_id,
                                 f.span[0] if f.span else 0,
                                 -GRADE_RANK[f.grade]))
    return resolved


def build_marks(deck: Deck, resolved: list[Finding]) -> list[Mark]:
    """── 4) 칠할 구간 계획 — 겹치는 구간은 하나로 합치고 더 높은 등급을 따른다 ──
    (pipeline 이 검토 기억을 적용해 후보를 바꾼 뒤 다시 부를 수 있도록 따로 떼어 둠)"""
    seg_map: dict[int, Segment] = {s.seg_id: s for s in deck.segments}
    marks: list[Mark] = []
    by_seg: dict[int, list[Finding]] = {}
    for f in resolved:
        seg = seg_map.get(f.seg_id)
        if seg is None or not seg.markable:        # 차트 안 글자 등은 칠할 수 없음
            continue
        by_seg.setdefault(f.seg_id, []).append(f)

    for seg_id, group in by_seg.items():
        seg = seg_map[seg_id]
        whole = [f for f in group if f.span is None]
        spans = sorted((f for f in group if f.span is not None),
                       key=lambda f: f.span)
        if whole:
            # 문단 전체를 칠해야 하는 항목이 하나라도 있으면 전체를 한 번만 칠한다
            best = max(whole + spans, key=lambda f: GRADE_RANK[f.grade])
            marks.append(Mark(seg_id, seg.slide_no, None, best.grade,
                              any(f.disclosure_risk for f in group)))
            continue
        # 겹치거나 맞닿은 구간을 왼쪽부터 병합
        cur_s, cur_e, cur_g, cur_r = None, None, "C", False
        for f in spans:
            s, e = f.span
            if cur_s is None:
                cur_s, cur_e, cur_g, cur_r = s, e, f.grade, f.disclosure_risk
                continue
            if s <= cur_e:  # 겹치거나 맞닿음
                cur_e = max(cur_e, e)
                if GRADE_RANK[f.grade] > GRADE_RANK[cur_g]:
                    cur_g = f.grade
                cur_r = cur_r or f.disclosure_risk
            else:
                marks.append(Mark(seg_id, seg.slide_no, (cur_s, cur_e), cur_g, cur_r))
                cur_s, cur_e, cur_g, cur_r = s, e, f.grade, f.disclosure_risk
        if cur_s is not None:
            marks.append(Mark(seg_id, seg.slide_no, (cur_s, cur_e), cur_g, cur_r))

    return marks
