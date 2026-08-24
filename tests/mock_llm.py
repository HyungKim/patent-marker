"""
tests/mock_llm.py ─ "가짜 Ollama" 로 모델 연동 경로 전체를 점검 (몇 초)
=====================================================================
진짜 모델 대신, 정해진 규칙으로 답하는 작은 HTTP 서버를 이 파일 안에서 띄웁니다.
프롬프트 구성 → 응답 파싱 → 인용구 위치 확정 → 병합 → PPTX 마킹이 끊기지 않는지 확인합니다.

    실행:  python tests/mock_llm.py
           python tests/mock_llm.py samples/회사보고자료_예시.pptx out_mock.pptx
"""
from __future__ import annotations

import json
import re
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

PORT = 11599   # 진짜 Ollama(11434) 와 겹치지 않는 포트

# 모델이 실제로 잡아야 하는 유형들을 흉내 낸다.
# 인용구는 반드시 원문에 그대로 있어야 하므로, 문단 텍스트에서 직접 골라낸다.
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


class Handler(BaseHTTPRequestHandler):
    """Ollama 의 /api/tags, /api/chat 을 흉내 내는 최소 구현."""

    def log_message(self, *a):
        pass

    def do_GET(self):
        self._send({"models": [{"name": "qwen3:14b"}]})

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
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
                findings.append({
                    "seg_id": sid, "quote": hit.group(0), "grade": grade,
                    "category": cat, "implicit": implicit,
                    "disclosure_risk": cat == "공개이력", "reason": reason,
                })
                break
        self._send({"message": {"content": json.dumps({"findings": findings},
                                                      ensure_ascii=False)}})

    def _send(self, obj):
        raw = json.dumps(obj).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)


srv = HTTPServer(("127.0.0.1", PORT), Handler)
threading.Thread(target=srv.serve_forever, daemon=True).start()

from app import config  # noqa: E402

config.OLLAMA_HOST = f"http://127.0.0.1:{PORT}"     # 가짜 서버로 향하게

from app import analyze, extract, mark, merge  # noqa: E402

analyze.config.OLLAMA_HOST = config.OLLAMA_HOST

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

stats = mark.apply(deck, resolved, marks, add_summary=True, tag_marks=True)
deck.prs.save(dst)
print(f"저장: {dst}  {stats}\n")

for f in resolved:
    if f.source != "llm":
        continue
    flag = "⚠" if f.disclosure_risk else " "
    imp = "묵시" if f.implicit else "명시"
    print(f"  p{f.slide_no} [{f.grade}]{flag} {imp} {f.category:<14} {f.quote[:38]!r}")
