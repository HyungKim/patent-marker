"""
tests/mock_llm.py ─ "가짜 Ollama" 로 모델 연동 경로 전체를 점검 (몇 초)
=====================================================================
진짜 모델 대신, 정해진 규칙으로 답하는 작은 HTTP 서버를 이 파일 안에서 띄웁니다.
프롬프트 구성 → 응답 파싱 → 인용구 위치 확정 → 병합 → PPTX 마킹이 끊기지 않는지 확인합니다.

    실행:  python tests/mock_llm.py
           python tests/mock_llm.py samples/회사보고자료_예시.pptx out_mock.pptx
           python tests/mock_llm.py --serve        ← 가짜 서버만 띄워 두기 (웹 화면·명령행 점검용)
           python tests/mock_llm.py --serve 20     ← 답을 20초 늦게 (느린 PC 흉내 — 경과 표시·[중단] 점검)

다른 테스트에서:  import mock_llm; mock_llm.start()  → 포트 11599 에 가짜 서버가 뜬다
"""
from __future__ import annotations

import json
import re
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

PORT = 11599   # 진짜 Ollama(11434) 와 겹치지 않는 포트

# 모델이 실제로 잡아야 하는 유형들을 흉내 낸다.
# 인용구는 반드시 원문에 그대로 있어야 하므로, 문단 텍스트에서 직접 골라낸다.
# (네 번째 값 '묵시 여부' 는 도구가 config.IMPLICIT_CATEGORIES 로 정하는 값과 같게 적어 둔 참고용)
PICKERS = [
    (re.compile(r"자체\s*제작|자사\s*설계|독자"), "B", "독자성주장", True,
     "독자 설계 주장 — 사내에 미공개 구성이 존재함을 시사"),
    (re.compile(r"\d+\s*%\s*(?:절감|향상|개선)"), "B", "효과만기재", True,
     "정량 효과만 기재, 수단 미기재 — 발명자 인터뷰 필요"),
    (re.compile(r"\d+\s*[~∼]\s*\d+\s*°"), "A", "수치·범위한정", False,
     "각도 범위 한정 — 수치한정 청구항 소재"),
    (re.compile(r"\d+nm[^,]{0,20}\d+nm|200Hz"), "A", "구성·구조", False,
     "파장·주기가 특정된 구성 — 장치 청구항 소재"),
    (re.compile(r"시연|논문|학술대회"), "B", "공개이력", False,
     "외부 공개 이력 — 공지예외주장 기한 확인 필요"),
    (re.compile(r"표준\s*사양으로\s*확정|최적화"), "B", "최적화·조건확립", True,
     "실험으로 도출한 조건 — 값이 미기재된 수치한정 소재"),
]


def _fake_vec(text: str, dims: int = 64) -> list[float]:
    import hashlib
    t = "".join(text.split()).lower()
    v = [0.0] * dims
    for i in range(max(len(t) - 1, 1)):
        h = int(hashlib.md5(t[i:i + 2].encode("utf-8")).hexdigest(), 16)
        v[h % dims] += 1.0
    n = sum(x * x for x in v) ** 0.5 or 1.0
    return [x / n for x in v]


class Handler(BaseHTTPRequestHandler):
    """Ollama 의 /api/tags, /api/chat, /api/embed 를 흉내 내는 최소 구현."""

    protocol_version = "HTTP/1.1"     # 진짜 Ollama 처럼 chunked 스트리밍
    delay = 0.0                       # 테스트용: 답하기 전에 이만큼 기다린다 (느린 PC 흉내)

    def log_message(self, *a):
        pass

    def do_GET(self):
        # bge-m3 도 '설치됨' 으로 답한다 — 검토 학습의 뜻 기준 검색 경로를 시험하기 위해
        self._send({"models": [{"name": "qwen3:8b"}, {"name": "qwen3:14b"}, {"name": "bge-m3"}]})

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        if self.path.rstrip("/").endswith("/api/embed"):
            # 가짜 임베딩: 글자 두 개 조각을 64칸에 흩뿌린 벡터 — 비슷한 글일수록 비슷한 벡터가 된다
            inp = body.get("input")
            texts = [inp] if isinstance(inp, str) else list(inp or [])
            self._send({"model": body.get("model"), "embeddings": [_fake_vec(t) for t in texts]})
            return
        if Handler.delay:
            time.sleep(Handler.delay)
        user = body["messages"][-1]["content"]
        findings = []
        for line in user.splitlines():
            m = re.match(r"#(\d+) \[[^\]]+\] (.*)", line)
            if not m:
                continue
            sid, text = int(m.group(1)), m.group(2)
            for pat, grade, cat, implicit, reason in PICKERS:
                hit = pat.search(text)
                if not hit:
                    continue
                # 진짜 모델과 같은 축약 형식(i·q·g·c·d·r). implicit 은 도구가 분류로 정하므로 보내지 않는다.
                findings.append({
                    "i": sid, "q": hit.group(0), "g": grade, "c": cat,
                    "d": cat == "공개이력", "r": reason[:12],
                })
                break
        text = json.dumps({"findings": findings}, ensure_ascii=False)
        # 진짜 Ollama 가 마지막 조각에 실어 보내는 집계 흉내 (토큰 ≒ 글자 수 / 2, 속도는 고정값)
        tally = {"prompt_eval_count": len(user) // 2, "prompt_eval_duration": int(0.2e9),
                 "eval_count": max(1, len(text) // 2), "eval_duration": int(max(1, len(text) // 2) / 20 * 1e9)}
        if body.get("stream"):
            # 진짜 Ollama 처럼 조각(NDJSON)으로 나눠 보낸다 — analyze._chat 의 이어 붙이기 경로를 검사
            cut = max(1, len(text) // 3)
            pieces = [text[i:i + cut] for i in range(0, len(text), cut)]
            self._send_stream([{"message": {"role": "assistant", "content": pc}, "done": False} for pc in pieces]
                              + [{"message": {"role": "assistant", "content": ""}, "done": True, **tally}])
        else:
            self._send({"message": {"content": text}, **tally})

    def _send(self, obj):
        raw = json.dumps(obj).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _send_stream(self, objs):
        try:
            self.send_response(200)
            self.send_header("Content-Type", "application/x-ndjson")
            self.send_header("Transfer-Encoding", "chunked")
            self.end_headers()
            for obj in objs:
                raw = json.dumps(obj).encode() + b"\n"
                self.wfile.write(f"{len(raw):x}\r\n".encode() + raw + b"\r\n")
                self.wfile.flush()
            self.wfile.write(b"0\r\n\r\n")
        except (BrokenPipeError, ConnectionResetError):
            pass                      # 클라이언트가 [중단] 으로 먼저 끊은 경우 — 정상

    def handle(self):
        try:
            super().handle()
        except (BrokenPipeError, ConnectionResetError):
            pass


_SERVER: ThreadingHTTPServer | None = None


def start() -> str:
    """가짜 Ollama 를 백그라운드 스레드로 띄우고 주소를 돌려준다 (이미 떠 있으면 그대로)."""
    global _SERVER
    host = f"http://127.0.0.1:{PORT}"
    if _SERVER is None:
        try:
            # 진짜 Ollama 처럼 요청을 동시에 받는다 (/api/chat 이 느려도 /api/tags 는 바로 답함)
            _SERVER = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
        except OSError:
            # 다른 창에서 `python tests/mock_llm.py --serve` 로 이미 띄워 둔 경우 — 그대로 쓴다
            import urllib.request
            with urllib.request.urlopen(f"{host}/api/tags", timeout=3) as r:
                if b"qwen3" not in r.read():
                    raise
            return host
        threading.Thread(target=_SERVER.serve_forever, daemon=True).start()
    return host


def main() -> None:
    host = start()

    from app import config

    config.OLLAMA_HOST = host                       # 가짜 서버로 향하게

    from app import analyze, extract, mark, merge

    analyze.config.OLLAMA_HOST = config.OLLAMA_HOST

    if len(sys.argv) > 1 and sys.argv[1] == "--serve":
        if len(sys.argv) > 2:                      # --serve 20 : 답하기 전 20초 기다림 (느린 PC 흉내)
            Handler.delay = float(sys.argv[2])
        print(f"가짜 Ollama 대기 중: {host}   (Ctrl+C 로 종료, 지연 {Handler.delay}초)")
        print(f"  다른 창에서:  PM_OLLAMA_HOST={host} python -m app.main")
        try:
            threading.Event().wait()
        except KeyboardInterrupt:
            return

    src = sys.argv[1] if len(sys.argv) > 1 else "samples/회사보고자료_예시.pptx"
    dst = sys.argv[2] if len(sys.argv) > 2 else "out_mock.pptx"

    print("health:", analyze.health())

    deck = extract.extract(src)
    opts = config.RunOptions()
    hits_by_seg, targets = analyze.prescreen(deck, opts)

    by_slide: dict[int, list] = {}
    for seg in targets:
        by_slide.setdefault(seg.slide_no, []).append(seg)

    all_f = []
    for n in range(1, deck.slide_count + 1):
        segs = by_slide.get(n, [])
        if not segs:
            continue
        hints = {s.seg_id: sorted({h.category for h in hits_by_seg.get(s.seg_id, [])})
                 for s in segs}
        hints = {k: v for k, v in hints.items() if v}
        got = analyze.analyze_slide("테스트", n, deck.slide_count, segs, hints, opts)
        print(f"  슬라이드 {n}: 모델 응답 {len(got)}건")
        all_f += got

    resolved, marks = merge.resolve(deck, all_f, hits_by_seg)
    located = sum(1 for f in resolved if f.span is not None)
    print(f"\n총 후보 {len(resolved)}건 · 인용구 위치 확정 {located}건 "
          f"({located / max(len(resolved), 1) * 100:.0f}%) · 하이라이트 {len(marks)}구간")

    stats = mark.apply(deck, resolved, marks)   # 형광펜 + 첫 슬라이드 범례만
    deck.prs.save(dst)
    print(f"저장: {dst}  {stats}\n")

    for f in resolved:
        if f.source != "llm":
            continue
        flag = "C" if f.disclosure_risk else " "
        imp = "묵시" if f.implicit else "명시"
        print(f"  p{f.slide_no} [{f.grade}]{flag} {imp} {f.category:<14} {f.quote[:38]!r}")


if __name__ == "__main__":
    main()
