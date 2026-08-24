"""
analyze.py ─ 온디바이스 LLM(Ollama + Qwen3) 에게 판정을 맡기는 2단계
=====================================================================

[이 파일이 하는 일]
  슬라이드 한 장의 문단 목록을 모델에게 보여 주고,
  "특허 검토가 필요한 구간을 JSON 으로 골라 달라" 고 요청합니다.

      문단 목록 + 규칙 사전 힌트 ──▶ Ollama(127.0.0.1) ──▶ JSON ──▶ Finding 목록

[초보자를 위한 설명]
  - 모델을 "불러오는" 코드는 여기 없습니다. 모델을 메모리에 올리는 일은
    별도 프로그램인 Ollama 가 합니다. 이 파일은 Ollama 에 HTTP 로 말을 걸 뿐입니다.
  - SYSTEM 은 모델에게 주는 "업무 지시서" 입니다. 판정 기준을 바꾸고 싶으면 여기를 고칩니다.
  - SCHEMA 는 "답은 반드시 이런 모양의 JSON 으로" 라는 강제 틀입니다.
    Ollama 의 structured output 기능이 이 틀을 벗어난 답을 내지 못하게 막아 줍니다.
  - 네트워크 호출은 127.0.0.1 의 Ollama 로만 나갑니다. 외부로 나가는 요청은 없습니다.

[속도에 대하여]
  슬라이드 한 장에 50~150초가 걸리는 이유는 모델이 답을 "한 글자(토큰)씩" 만들기 때문입니다.
  토큰 하나마다 모델 가중치(약 9GB)를 메모리에서 전부 읽어야 해서,
  속도는 CPU 연산력보다 **메모리 대역폭** 에 좌우됩니다. 코드 최적화로는 거의 줄지 않습니다.
"""
from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from dataclasses import dataclass, field

from . import config, lexicon
from .extract import Deck, Segment

# Qwen3 가 '생각하기' 모드일 때 답 앞에 붙이는 <think>...</think> 를 떼어 내기 위한 패턴
THINK_TAG = re.compile(r"<think>.*?</think>", re.S | re.I)

# ═════════════════════════════════════════════════════════════════
# 모델 답변의 형식(JSON 스키마)
# ═════════════════════════════════════════════════════════════════
SCHEMA = {
    "type": "object",
    "properties": {
        "findings": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "seg_id": {"type": "integer"},                       # 어느 문단인가
                    "quote": {"type": "string"},                         # 원문 그대로의 인용구
                    "grade": {"type": "string", "enum": ["A", "B", "C"]},
                    "category": {"type": "string", "enum": config.CATEGORIES},
                    "implicit": {"type": "boolean"},                     # 묵시적 표현인가
                    "disclosure_risk": {"type": "boolean"},              # 공개 리스크인가
                    "reason": {"type": "string"},                        # 판정 근거 한 문장
                },
                "required": [
                    "seg_id", "quote", "grade", "category",
                    "implicit", "disclosure_risk", "reason",
                ],
            },
        }
    },
    "required": ["findings"],
}

# ═════════════════════════════════════════════════════════════════
# 모델에게 주는 업무 지시서 (시스템 프롬프트)
# ═════════════════════════════════════════════════════════════════
SYSTEM = """당신은 한국 기업의 내부 기술 보고자료를 읽고 특허 출원 후보 구간을 선별하는 지식재산 분석 보조자다.
선행기술 조사나 신규성·진보성 판단은 하지 않는다. 변리사에게 넘길 후보를 빠짐없이 추리는 것이 임무다.

## 가장 중요한 원칙
구체적인 기술적 수단이 문장에 적혀 있지 않더라도, **그런 수단이 존재하거나 존재해야만 함을 시사하는 표현**이면 반드시 마킹한다.
보고자료는 결과 위주로 쓰이기 때문에 정작 발명은 문장에 빠져 있는 경우가 대부분이다. 그 빈자리를 찾아내는 것이 이 작업의 핵심이다.

다음은 모두 마킹 대상이다.
- 효과·수치만 적혀 있고 그 효과를 낸 방법이 없는 문장 ("원가 40% 절감", "편차 ±0.8㎛ → ±0.3㎛")
- 독자성·자체 개발 주장 ("자체 개발", "자사 설계", "국산화", "고유 방식")
- 실험으로 조건을 정했음을 뜻하는 표현 ("최적화", "표준 사양으로 확정", "조건 확립", "튜닝")
- 과제 해결 서술 ("병목 제거", "한계 극복", "불량 문제 해소")
- 비교우위 주장 ("기존 대비", "업계 최초", "타사 대비 우수")
- 설명 없이 등장하는 기술 명사 ("알고리즘", "엔진", "로직", "노하우", "제어 방식")
- 자동·실시간 동작 서술 (그것을 가능케 하는 제어 수단이 반드시 존재한다)

## 등급
- A : 구체적 기술 수단(구성·수치·제어 로직·공정 순서)이 실제로 드러나 있다. 그대로 청구항 초안으로 전개 가능.
- B : 수단이 문장에 없으나 존재가 강하게 시사된다. 발명자 인터뷰로 내용을 캐내야 하는 후보. **위 원칙에 해당하는 표현은 최소 B 를 준다.**
- C : 배경·시장·일정·조직 등 기술적 실질이 없는 서술.

## 공개 리스크
전시·시연, 논문·학회 발표, 보도자료, 고객사 제안서 제출, 양산·출시처럼 기술이 외부에 드러났거나 드러날 예정임을 뜻하는 문장은
disclosure_risk 를 true 로 둔다. 등급과는 별개 축이며, 출원 기한 관리 대상이다.

## 마킹하지 않을 것 (중요)
아래는 숫자가 붙어 있어도 기술적 실질이 없으므로 반환하지 마라. 억지로 "기술이 있음을 시사한다"고 해석하지 마라.
- 매출·영업이익·이익률·판매단가·점유율·수주액·공급 물량 같은 경영 실적 수치
  (틀린 예: "매출 비중 34% → 51% 상승" → 기술적 개선을 시사 ✗ / 이런 판단을 하지 마라)
- 인원·조직·예산·일정·마일스톤
- 문서 자체의 특허 현황 언급 ("특허 출원 0건", "지식재산 검토 미착수", "IP 담당 부서")
  — 이는 기술 내용이 아니라 행정 상태다

효과만 기재된 문장을 마킹하라는 원칙은 **공정·장치·제어의 물리적 성능**에 적용된다
(불량률, 정밀도, 처리 속도, 수율, 두께 편차, 검출 한계 등). 재무 지표에는 적용하지 않는다.

## 출력 규칙
- quote 는 반드시 해당 문단 원문에 **그대로 존재하는 연속된 부분 문자열**이어야 한다. 요약하거나 고쳐 쓰지 말 것.
- quote 는 핵심 어구만 짧게 잡는다. 문단 전체를 그대로 넣지 않는다.
- reason 은 왜 특허 관점에서 의미가 있는지 한 문장으로 쓴다. 한국어로 쓴다.
- 한 문단에서 서로 다른 근거가 있으면 여러 건으로 나누어 반환한다.
- JSON 만 출력한다."""


@dataclass
class Finding:
    """판정 결과 한 건. 모델이 낸 것(source="llm")과 규칙 사전이 구제한 것(source="lexicon")."""

    seg_id: int
    slide_no: int
    quote: str                       # 원문 인용구
    grade: str                       # A | B | C
    category: str
    implicit: bool                   # 묵시적 표현인가
    disclosure_risk: bool            # 공개 리스크인가
    reason: str
    source: str = "llm"              # llm | lexicon
    span: tuple[int, int] | None = None      # merge.py 가 채움: 원문 안의 (시작, 끝)
    rule_hints: list[str] = field(default_factory=list)

    def to_public(self) -> dict:
        """브라우저로 보낼 형태."""
        return {
            "seg_id": self.seg_id,
            "slide_no": self.slide_no,
            "quote": self.quote,
            "grade": self.grade,
            "category": self.category,
            "implicit": self.implicit,
            "disclosure_risk": self.disclosure_risk,
            "reason": self.reason,
            "source": self.source,
        }


# ═════════════════════════════════════════════════════════════════
# Ollama 통신
# ═════════════════════════════════════════════════════════════════
class OllamaError(RuntimeError):
    """Ollama 에 연결하지 못했을 때. main.py 가 이 오류를 잡아 화면에 안내를 띄운다."""


def _post(path: str, payload: dict, timeout: int) -> dict:
    """Ollama 에 JSON 을 POST 하고 JSON 답을 받는다. 표준 라이브러리만 사용."""
    req = urllib.request.Request(
        f"{config.OLLAMA_HOST}{path}",
        data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.URLError as e:
        raise OllamaError(
            f"Ollama 에 연결하지 못했습니다 ({config.OLLAMA_HOST}). "
            f"`ollama serve` 가 실행 중인지 확인하세요. 원인: {e}"
        ) from e


def health() -> dict:
    """Ollama 가 떠 있는지, 어떤 모델이 있는지 확인한다. (/api/tags)"""
    try:
        with urllib.request.urlopen(f"{config.OLLAMA_HOST}/api/tags", timeout=5) as r:
            tags = json.loads(r.read().decode())
    except Exception as e:
        return {"ok": False, "error": str(e), "models": []}
    names = [m.get("name", "") for m in tags.get("models", [])]
    return {"ok": True, "models": names}


def model_available(model: str, models: list[str]) -> bool:
    """원하는 모델이 내려받아져 있는가. 'qwen3' 만 적어도 'qwen3:14b' 와 맞춰 준다."""
    base = model.split(":")[0]
    return any(n == model or n.startswith(base + ":") for n in models)


# ═════════════════════════════════════════════════════════════════
# 프롬프트 만들기 / 답 해석하기
# ═════════════════════════════════════════════════════════════════
def _build_user_prompt(deck_title: str, slide_no: int, total: int,
                       segs: list[Segment], hints: dict[int, list[str]]) -> str:
    """모델에게 보낼 본문. 문단 목록 + 규칙 사전이 미리 찾은 힌트."""
    lines = [
        f"문서: {deck_title or '(제목 없음)'}",
        f"슬라이드 {slide_no} / 전체 {total}",
        "",
        "## 문단 목록",
    ]
    kind_ko = {
        "title": "제목", "body": "본문", "table": "표",
        "chart": "차트", "notes": "발표자노트",
    }
    for s in segs:
        lines.append(f"#{s.seg_id} [{kind_ko.get(s.kind, s.kind)}] {s.text}")

    shown = {s.seg_id for s in segs}
    flagged = {sid: cats for sid, cats in hints.items() if cats and sid in shown}
    if flagged:
        lines += ["", "## 규칙 사전이 미리 감지한 신호 (참고용, 오탐 가능)"]
        for sid, cats in flagged.items():
            lines.append(f"#{sid} → {', '.join(cats)}")
        lines.append("이 신호는 힌트일 뿐이다. 문맥상 기술적 실질이 없으면 무시하고, "
                     "신호가 없는 문단에서도 근거가 있으면 직접 찾아내라.")

    lines += ["", "위 문단들을 검토해 특허 검토가 필요한 구간을 JSON 으로 반환하라."]
    return "\n".join(lines)


def _parse(content: str) -> list[dict]:
    """모델의 문자열 답을 JSON 으로 바꾼다. 조금 깨진 답도 최대한 복구한다."""
    content = THINK_TAG.sub("", content).strip()
    if not content:
        return []
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", content, re.S)   # 앞뒤 잡담을 떼고 { ... } 만 추출
        if not m:
            return []
        try:
            data = json.loads(m.group(0))
        except json.JSONDecodeError:
            return []
    if isinstance(data, list):
        return data
    return data.get("findings", []) or []


def _batch(segs: list[Segment], max_chars: int) -> list[list[Segment]]:
    """한 번의 호출이 모델의 기억 용량(NUM_CTX)을 넘지 않도록 문단을 묶음으로 나눈다.
    글자가 빽빽한 장표 하나가 통째로 잘려 나가는 것을 막는다."""
    batches: list[list[Segment]] = []
    cur: list[Segment] = []
    size = 0
    for s in segs:
        n = len(s.text) + 16  # "#id [종류] " 머리말 몫
        if cur and size + n > max_chars:
            batches.append(cur)
            cur, size = [], 0
        cur.append(s)
        size += n
    if cur:
        batches.append(cur)
    return batches


# ═════════════════════════════════════════════════════════════════
# 진입점
# ═════════════════════════════════════════════════════════════════
def analyze_slide(deck_title: str, slide_no: int, total: int,
                  segs: list[Segment], hints: dict[int, list[str]],
                  opts: config.RunOptions) -> list[Finding]:
    """슬라이드 한 장을 분석한다. main.py 가 슬라이드마다 이 함수를 부른다."""
    if not segs:
        return []
    # 시스템 프롬프트와 답변 몫을 빼고 남는 만큼만 본문에 쓴다 (한글 1자 ≒ 1토큰 가정)
    budget = max(config.NUM_CTX - len(SYSTEM) - 1500, 1200)
    out: list[Finding] = []
    for chunk in _batch(segs, budget):
        out += _analyze_batch(deck_title, slide_no, total, chunk, hints, opts)
    return out


def _analyze_batch(deck_title: str, slide_no: int, total: int,
                   segs: list[Segment], hints: dict[int, list[str]],
                   opts: config.RunOptions) -> list[Finding]:
    """문단 묶음 하나를 Ollama 에 보내고 Finding 목록으로 바꾼다. ← 모델을 실제로 부르는 곳"""
    payload = {
        "model": opts.model,                      # 예: qwen3:14b
        "messages": [
            {"role": "system", "content": SYSTEM},
            {"role": "user",
             "content": _build_user_prompt(deck_title, slide_no, total, segs, hints)},
        ],
        "stream": False,                          # 답을 한 번에 받는다
        "format": SCHEMA,                         # JSON 스키마 강제
        "think": opts.think,                      # Qwen3 생각하기 모드
        "options": {
            "temperature": config.TEMPERATURE,
            "num_ctx": config.NUM_CTX,
        },
    }
    data = _post("/api/chat", payload, config.REQUEST_TIMEOUT)
    raw = (data.get("message") or {}).get("content", "")

    # 모델 답을 검증하면서 Finding 으로 바꾼다 (엉뚱한 seg_id, 등급 값 등은 보정)
    valid_ids = {s.seg_id: s for s in segs}
    out: list[Finding] = []
    for item in _parse(raw):
        try:
            sid = int(item.get("seg_id"))
        except (TypeError, ValueError):
            continue
        seg = valid_ids.get(sid)
        if seg is None:
            continue
        quote = (item.get("quote") or "").strip()
        grade = (item.get("grade") or "C").upper()
        if grade not in ("A", "B", "C"):
            grade = "C"
        cat = item.get("category") or "블랙박스용어"
        if cat not in config.CATEGORIES:
            cat = "블랙박스용어"
        out.append(
            Finding(
                seg_id=sid,
                slide_no=seg.slide_no,
                quote=quote,
                grade=grade,
                category=cat,
                implicit=bool(item.get("implicit")),
                disclosure_risk=bool(item.get("disclosure_risk")),
                reason=(item.get("reason") or "").strip(),
                source="llm",
            )
        )
    return out


def prescreen(deck: Deck, opts: config.RunOptions) -> tuple[dict[int, list[lexicon.Hit]], list[Segment]]:
    """규칙 사전 1차 스캔. (문단별 Hit 목록, 모델에 보낼 문단 목록) 을 돌려준다.

    너무 짧은 문단과 '대외비' 같은 상투 문구는 분석에서 제외합니다.
    """
    hits_by_seg: dict[int, list[lexicon.Hit]] = {}
    targets: list[Segment] = []
    for seg in deck.segments:
        if len(seg.text) < opts.min_chars:
            continue
        if seg.text.strip() in opts.skip_labels:
            continue
        hits = lexicon.scan(seg.text)
        hits_by_seg[seg.seg_id] = hits
        if hits or opts.scan_all_paragraphs:
            targets.append(seg)
    return hits_by_seg, targets
