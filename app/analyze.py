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
  슬라이드 한 장에 몇 분이 걸리는 이유는 모델이 답을 "한 글자(토큰)씩" 만들기 때문입니다.
  토큰 하나마다 모델 가중치(qwen3:8b 약 5GB)를 메모리에서 전부 읽어야 해서,
  속도는 CPU 연산력보다 **메모리 대역폭** 에 좌우됩니다. 실측(예시 문서 5장)으로 시간의 9할이
  '답을 쓰는' 시간이었으므로, 이 도구는 **답을 짧게** 받도록 설계했습니다 (2026-09-21):
    · JSON 키를 한 글자(i·q·g·c·d·r)로, 판정 사유는 12자 이내, 인용구는 25자 이내
    · '묵시적 표현인가' 는 모델에게 묻지 않고 분류(config.IMPLICIT_CATEGORIES)로 정함
    · 같은 문단·같은 유형은 한 건만, C(기술적 실질 없음)는 반환하지 않음
  모델이 보낸 토큰 수와 초당 속도는 analyze_slide(stats=) 로 집계되어 화면·실행 기록에 남습니다.
"""
from __future__ import annotations

import http.client
import json
import re
import socket
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from typing import Callable

from . import config, lexicon
from .extract import Deck, Segment

# Qwen3 가 '생각하기' 모드일 때 답 앞에 붙이는 <think>...</think> 를 떼어 내기 위한 패턴
THINK_TAG = re.compile(r"<think>.*?</think>", re.S | re.I)

# ═════════════════════════════════════════════════════════════════
# 모델 답변의 형식(JSON 스키마)
# ═════════════════════════════════════════════════════════════════
# 키를 한 글자로 둔 이유: 모델이 답을 한 토큰씩 쓰므로 키 이름도 매 건 시간을 먹는다.
# 예시 문서 실측에서 한 건에 약 120 토큰 중 절반이 키 이름·상투적 사유·중복 플래그였다.
#   i = 문단 번호   q = 원문 인용구   g = 등급(A/B/C)   c = 분류   d = 공개 관련 문장인가   r = 사유(짧게)
# 옛 형식(seg_id, quote, grade, category, implicit, disclosure_risk, reason)도 _norm_item 이 그대로 받는다.
SCHEMA = {
    "type": "object",
    "properties": {
        "findings": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "i": {"type": "integer"},                            # 어느 문단인가 (#번호)
                    "q": {"type": "string"},                             # 원문 그대로의 인용구
                    "g": {"type": "string", "enum": ["A", "B", "C"]},    # 등급
                    "c": {"type": "string", "enum": config.CATEGORIES},  # 분류
                    "d": {"type": "boolean"},                            # 공개 관련 문장인가
                    "r": {"type": "string"},                             # 판정 사유 (12자 이내)
                },
                "required": ["i", "q", "g", "c", "d", "r"],
            },
        }
    },
    "required": ["findings"],
}

# 모델 답 한 건을 내부 이름으로 맞춘다. 새 형식(한 글자 키)과 옛 형식 모두 받는다.
_KEY_ALIASES = {
    "seg_id": ("i", "seg_id"), "quote": ("q", "quote"), "grade": ("g", "grade"),
    "category": ("c", "category"), "disclosure_risk": ("d", "disclosure_risk"),
    "reason": ("r", "reason"), "implicit": ("m", "implicit"),
}


def _norm_item(item: dict) -> dict:
    out = {}
    for name, keys in _KEY_ALIASES.items():
        for k in keys:
            if k in item:
                out[name] = item[k]
                break
    return out

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

## 공개 관련정보
전시·시연, 논문·학회 발표, 보도자료, 고객사 제안서 제출, 양산·출시처럼 기술이 외부에 드러났거나 드러날 예정임을 뜻하는 문장은
d 를 true 로 둔다. 이렇게 표시된 구간은 결과에서 C(공개 관련정보) 로 안내되며,
공개 전에 출원을 끝냈어야 하는 대상이다. 그 밖의 문장은 d 를 false 로 둔다.
"특허 출원 0건", "출원 미착수", "IP 검토 예정" 처럼 **특허 행정 상태**를 말하는 문장은 공개가 아니다 — 반환하지 마라.

## 마킹하지 않을 것 (중요)
아래는 숫자가 붙어 있어도 기술적 실질이 없으므로 반환하지 마라. 억지로 "기술이 있음을 시사한다"고 해석하지 마라.
- 매출·영업이익·이익률·판매단가·점유율·수주액·공급 물량 같은 경영 실적 수치
  (틀린 예: "매출 비중 34% → 51% 상승" → 기술적 개선을 시사 ✗ / 이런 판단을 하지 마라)
- 인원·조직·예산·일정·마일스톤
- 문서 자체의 특허 현황 언급 ("특허 출원 0건", "지식재산 검토 미착수", "IP 담당 부서")
  — 이는 기술 내용이 아니라 행정 상태다

효과만 기재된 문장을 마킹하라는 원칙은 **공정·장치·제어의 물리적 성능**에 적용된다
(불량률, 정밀도, 처리 속도, 수율, 두께 편차, 검출 한계 등). 재무 지표에는 적용하지 않는다.

## 출력 규칙 (답은 짧을수록 좋다 — 한 글자마다 시간이 든다)
JSON 한 개만 출력한다: {"findings": [{"i": 문단 번호, "q": 인용구, "g": 등급, "c": 분류, "d": 공개 여부, "r": 사유}, ...]}
- i : 문단 앞의 # 번호 (숫자만).
- q : 반드시 해당 문단 원문에 **그대로 존재하는 연속된 부분 문자열**. 요약하거나 고쳐 쓰지 말 것.
      핵심 어구만 **25자 이내**로 짧게 잡는다. 문장 전체를 넣지 않는다.
- g : A 또는 B. C(기술적 실질 없음)는 반환하지 않는다 — 단 공개 관련 문장은 d:true 로 반환한다.
- c : 분류 이름 하나.
- r : 특허 관점의 근거를 **12자 이내** 핵심어로만 쓴다. 예: "수단 없는 효과 수치", "독자 설계 주장", "학회 발표 이력". 문장으로 풀어 쓰지 않는다.
- 한 문단에 서로 **다른 분류**의 근거가 있으면 건을 나누고, **같은 분류**의 근거가 여럿이면 가장 대표적인 것 한 건만 반환한다."""


def _system_prompt() -> str:
    """업무 지시서 + (있다면) 검토 반영 탭에서 쌓인 '사내 확정 사례' 블록.

    사람이 검토완료본으로 교정한 오탐·누락 사례가 review_data/examples.json 에
    쌓이면, 여기서 자동으로 프롬프트에 붙어 다음 분석부터 반영됩니다.
    """
    try:
        from . import review
        block = review.examples_block()
    except Exception:
        block = ""
    return SYSTEM + ("\n\n" + block if block else "")


@dataclass
class Finding:
    """판정 결과 한 건. 모델이 낸 것(source="llm")과 규칙 사전이 구제한 것(source="lexicon")."""

    seg_id: int
    slide_no: int
    quote: str                       # 원문 인용구
    grade: str                       # A | B | C
    category: str
    implicit: bool                   # 묵시적 표현인가
    disclosure_risk: bool            # 이미 공개된 내용인가
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


# Ollama 는 같은 PC(127.0.0.1)에 있으므로 회사 프록시(HTTP_PROXY 환경 변수·인터넷 옵션)를 거치면 안 된다.
# 프록시를 무시하는 전용 열기 도구 — health()·_server_gone() 이 쓴다. (_chat 은 http.client 를 직접 써서 원래 프록시를 안 탄다)
_DIRECT = urllib.request.build_opener(urllib.request.ProxyHandler({}))


class Aborted(Exception):
    """사용자의 '중단' 신호로 모델 호출을 끊었을 때. pipeline.py 가 Cancelled 로 바꿔 올린다."""


# 진행 상황 콜백: (지금까지 받은 답변 글자 수) → None.  모델이 아직 문단을 읽는 중이면 0.
ProgressFn = Callable[[int], None]

_GUARD_TICK = 0.5        # 초. 중단 신호를 이만큼 간격으로 확인한다
_PING_EVERY = 15         # 초. 서버 생존 확인(/api/tags) 간격
_PING_FAILS = 3          # 연속 이만큼 '연결 거부' 면 서버가 죽은 것으로 본다


def _server_gone() -> bool:
    """Ollama 프로세스가 사라졌는지 확인. 느려서 늦게 답하는 것은 '살아 있음' 으로 본다."""
    try:
        with _DIRECT.open(f"{config.OLLAMA_HOST}/api/tags", timeout=5):
            return False
    except urllib.error.URLError as e:
        # 연결 거부·리셋 = 프로세스 없음. 그 밖(시간 초과 등)은 바쁜 것일 수 있어 살아 있다고 본다
        return isinstance(e.reason, (ConnectionRefusedError, ConnectionResetError))
    except (ConnectionRefusedError, ConnectionResetError):
        return True
    except Exception:  # noqa: BLE001
        return False


def _chat(payload: dict, cancel: threading.Event | None = None,
          progress: ProgressFn | None = None) -> dict:
    """/api/chat 을 스트리밍으로 부르고, 조각을 이어 붙여 한 번에 받은 것과 같은 모양의 dict 로 돌려준다.

    [시간 제한을 두지 않는 대신 이렇게 지킨다]
      · 회사 PC(CPU 만, 16GB)에서는 슬라이드 하나에 10분이 넘게 걸릴 수 있어 고정 제한을 없앴다.
      · cancel 신호가 켜지면 감시 스레드가 소켓을 끊어 **즉시** 멈춘다 (Aborted). 모델이 아직
        문단을 읽는 중(아무 바이트도 안 온 상태)이어도 끊긴다.
      · 감시 스레드가 15초마다 /api/tags 로 서버 생존을 확인해, 프로세스가 사라졌으면 끊는다 (OllamaError).
      · config.REQUEST_TIMEOUT(PM_TIMEOUT) 이 0 보다 크면 "그 초 동안 아무 바이트도 안 올 때" 만 끊는다.
        답이 한 글자라도 계속 오는 동안은 시간이 흘러도 끊지 않는다 (무응답 제한이지 총량 제한이 아님).
      · progress 콜백으로 지금까지 받은 답변 글자 수를 알려 화면이 "살아 있음" 을 보여 준다.
    """
    u = urllib.parse.urlsplit(config.OLLAMA_HOST)
    https = u.scheme == "https"
    conn_cls = http.client.HTTPSConnection if https else http.client.HTTPConnection
    timeout = config.REQUEST_TIMEOUT if config.REQUEST_TIMEOUT > 0 else None
    conn = conn_cls(u.hostname or "127.0.0.1", u.port or (443 if https else 80), timeout=timeout)
    body = json.dumps({**payload, "stream": True}).encode()
    path = (u.path.rstrip("/") or "") + "/api/chat"

    state = {"done": False, "reason": ""}      # reason: "cancel" | "server_down"
    nchars = 0

    def _kill(reason: str) -> None:
        """다른 스레드에서 소켓을 끊어, 막혀 있는 읽기를 오류로 깨운다."""
        if state["done"]:
            return
        state["reason"] = reason
        try:
            sock = conn.sock
            if sock is not None:
                sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        try:
            conn.close()
        except Exception:  # noqa: BLE001
            pass

    def _guard() -> None:
        """중단 신호 확인 + 서버 생존 확인 + 2초마다 진행 콜백."""
        fails = 0
        last_ping = last_tick = time.monotonic()
        while not state["done"]:
            if cancel is not None and cancel.wait(_GUARD_TICK):
                _kill("cancel")
                return
            if cancel is None:
                time.sleep(_GUARD_TICK)
            now = time.monotonic()
            if now - last_ping >= _PING_EVERY:
                last_ping = now
                fails = fails + 1 if _server_gone() else 0
                if fails >= _PING_FAILS:
                    _kill("server_down")
                    return
            if progress is not None and now - last_tick >= 2:
                last_tick = now
                progress(nchars)

    if cancel is not None and cancel.is_set():
        raise Aborted()
    threading.Thread(target=_guard, daemon=True).start()

    content: list[str] = []
    final: dict = {}
    try:
        conn.request("POST", path, body=body, headers={"Content-Type": "application/json"})
        if state["reason"]:                      # 요청을 보내는 사이에 중단됨
            raise OSError("aborted")
        resp = conn.getresponse()                # 모델이 문단을 다 읽고 첫 글자를 쓸 때까지 여기서 기다린다
        if resp.status != 200:
            detail = resp.read(300).decode("utf-8", "replace")
            raise OllamaError(f"Ollama 응답 오류 {resp.status}: {detail}")
        last_report = 0.0
        for line in resp:                        # 한 줄 = JSON 조각 하나
            if not line.strip():
                continue
            obj = json.loads(line)
            piece = (obj.get("message") or {}).get("content") or ""
            if piece:
                content.append(piece)
                nchars += len(piece)
                if progress is not None and time.monotonic() - last_report >= 0.5:
                    last_report = time.monotonic()
                    progress(nchars)
            if obj.get("error"):
                raise OllamaError(f"Ollama 오류: {obj['error']}")
            if obj.get("done"):
                final = obj
                break
    except OllamaError:
        raise
    except (OSError, http.client.HTTPException, ValueError) as e:
        if state["reason"] == "cancel":
            raise Aborted() from None
        if state["reason"] == "server_down":
            raise OllamaError(
                f"Ollama 서버가 사라져 분석을 멈췄습니다 ({config.OLLAMA_HOST}). "
                "작업 표시줄의 Ollama 를 다시 실행하거나 run.bat 을 다시 켠 뒤 시도하세요."
            ) from e
        if isinstance(e, socket.timeout) or (isinstance(e, TimeoutError)):
            raise OllamaError(
                f"Ollama 가 {timeout}초 동안 아무 응답도 보내지 않아 멈췄습니다. "
                "느린 PC 라면 PM_TIMEOUT=0 (제한 없음, 기본값) 으로 두세요."
            ) from e
        raise OllamaError(
            f"Ollama 에 연결하지 못했습니다 ({config.OLLAMA_HOST}). "
            f"`ollama serve` 가 실행 중인지 확인하세요. 원인: {e}"
        ) from e
    finally:
        state["done"] = True
        try:
            conn.close()
        except Exception:  # noqa: BLE001
            pass
    if cancel is not None and cancel.is_set():
        raise Aborted()
    if progress is not None:
        progress(nchars)
    return {**final, "message": {**(final.get("message") or {}), "content": "".join(content)}}


def health() -> dict:
    """Ollama 가 떠 있는지, 어떤 모델이 있는지 확인한다. (/api/tags)"""
    try:
        with _DIRECT.open(f"{config.OLLAMA_HOST}/api/tags", timeout=5) as r:
            tags = json.loads(r.read().decode())
    except Exception as e:
        return {"ok": False, "error": str(e), "models": []}
    names = [m.get("name", "") for m in tags.get("models", [])]
    return {"ok": True, "models": names}


def model_available(model: str, models: list[str]) -> bool:
    """원하는 모델이 내려받아져 있는가.

    태그까지 정확히 맞아야 한다 — 'qwen3:8b' 가 있다고 'qwen3:4b-instruct' 가 있는 것은 아니다
    (화면의 정밀/빠름 선택이 이 값으로 설치 여부를 표시한다). 태그 없이 'qwen3' 만 적으면 어느 태그든 있으면 된다.
    """
    if model in models or model + ":latest" in models:
        return True
    if ":" not in model:
        return any(n.split(":")[0] == model for n in models)
    return False


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
def new_stats() -> dict:
    """모델 호출 집계 상자. analyze_slide(stats=) 에 넘기면 호출마다 더해진다.

    prompt_tokens / output_tokens : 모델이 읽은 토큰 · 쓴 토큰 (Ollama 가 세어 준 값)
    prompt_sec / output_sec       : 읽는 데 · 쓰는 데 걸린 초
    calls                         : 모델 호출 횟수
    """
    return {"prompt_tokens": 0, "output_tokens": 0, "prompt_sec": 0.0, "output_sec": 0.0, "calls": 0}


def add_stats(total: dict, part: dict) -> None:
    for k in ("prompt_tokens", "output_tokens", "prompt_sec", "output_sec", "calls"):
        total[k] = total.get(k, 0) + part.get(k, 0)


def analyze_slide(deck_title: str, slide_no: int, total: int,
                  segs: list[Segment], hints: dict[int, list[str]],
                  opts: config.RunOptions,
                  cancel: threading.Event | None = None,
                  progress: ProgressFn | None = None,
                  stats: dict | None = None,
                  system: str | None = None) -> list[Finding]:
    """슬라이드 한 장을 분석한다. pipeline.py 가 슬라이드마다 이 함수를 부른다.

    cancel   : '중단' 신호. 켜지면 진행 중인 모델 호출을 끊고 Aborted 를 올린다.
    progress : 지금까지 받은 답변 글자 수를 알리는 콜백 (화면의 "살아 있음" 표시용).
    stats    : new_stats() 로 만든 집계 상자. 넘기면 토큰 수·소요 초가 더해진다 (속도 기록용).
    system   : 이 슬라이드용 지시서 (검토 학습의 '유사 사례' 블록이 붙은 것). None 이면 공통 지시서.
    """
    if not segs:
        return []
    sys_prompt = system or _system_prompt()
    # 시스템 프롬프트와 답변 몫을 빼고 남는 만큼만 본문에 쓴다 (한글 1자 ≒ 1토큰 가정)
    # SYSTEM 만이 아니라 뒤에 붙는 '사내 확정 사례' 블록까지 포함한 실제 길이를 뺀다.
    # (블록을 빼지 않으면 예산을 실제보다 크게 잡아 모델 답변 몫이 모자랄 수 있다)
    budget = max(config.NUM_CTX - len(sys_prompt) - 1500, 1200)
    out: list[Finding] = []
    for chunk in _batch(segs, budget):
        out += _analyze_batch(deck_title, slide_no, total, chunk, hints, opts,
                              system_override=system, cancel=cancel, progress=progress, stats=stats)
    return out


def _analyze_batch(deck_title: str, slide_no: int, total: int,
                   segs: list[Segment], hints: dict[int, list[str]],
                   opts: config.RunOptions,
                   system_override: str | None = None,
                   cancel: threading.Event | None = None,
                   progress: ProgressFn | None = None,
                   stats: dict | None = None) -> list[Finding]:
    """문단 묶음 하나를 Ollama 에 보내고 Finding 목록으로 바꾼다. ← 모델을 실제로 부르는 곳

    system_override 는 성능 측정(evaluate.py)이 "최초 설정 프롬프트" 로
    채점할 때만 넘깁니다. 평소 분석에서는 None 이라 _system_prompt() 를 씁니다.
    """
    payload = {
        "model": opts.model,                      # 예: qwen3:8b
        "messages": [
            {"role": "system", "content": system_override or _system_prompt()},
            {"role": "user",
             "content": _build_user_prompt(deck_title, slide_no, total, segs, hints)},
        ],
        "format": SCHEMA,                         # JSON 스키마 강제 (_chat 이 stream=True 로 조각을 받아 이어 붙인다)
        "think": opts.think,                      # Qwen3 생각하기 모드
        "options": {
            "temperature": config.TEMPERATURE,
            "num_ctx": config.NUM_CTX,
        },
    }
    data = _chat(payload, cancel=cancel, progress=progress)
    raw = (data.get("message") or {}).get("content", "")
    if stats is not None:                         # Ollama 가 마지막 조각에 세어 보내는 토큰 수·소요 시간
        stats["prompt_tokens"] += int(data.get("prompt_eval_count") or 0)
        stats["output_tokens"] += int(data.get("eval_count") or 0)
        stats["prompt_sec"] += (data.get("prompt_eval_duration") or 0) / 1e9
        stats["output_sec"] += (data.get("eval_duration") or 0) / 1e9
        stats["calls"] += 1

    # 모델 답을 검증하면서 Finding 으로 바꾼다 (엉뚱한 seg_id, 등급 값 등은 보정)
    valid_ids = {s.seg_id: s for s in segs}
    out: list[Finding] = []
    for raw_item in _parse(raw):
        if not isinstance(raw_item, dict):
            continue
        item = _norm_item(raw_item)
        try:
            sid = int(item.get("seg_id"))
        except (TypeError, ValueError):
            continue
        seg = valid_ids.get(sid)
        if seg is None:
            continue
        quote = (item.get("quote") or "").strip()
        grade = str(item.get("grade") or "C").upper()
        if grade not in ("A", "B", "C"):
            grade = "C"
        cat = item.get("category") or "블랙박스용어"
        if cat not in config.CATEGORIES:
            cat = "블랙박스용어"
        # 묵시 여부는 모델에게 묻지 않고 분류로 정한다 (옛 형식 답에 값이 있으면 그 값을 존중)
        implicit = item["implicit"] if "implicit" in item else (cat in config.IMPLICIT_CATEGORIES)
        out.append(
            Finding(
                seg_id=sid,
                slide_no=seg.slide_no,
                quote=quote,
                grade=grade,
                category=cat,
                implicit=bool(implicit),
                disclosure_risk=bool(item.get("disclosure_risk")),
                reason=str(item.get("reason") or "").strip(),
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
