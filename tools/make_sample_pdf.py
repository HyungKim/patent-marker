"""
tools/make_sample_pdf.py ─ 예시 PPTX 의 글자를 그대로 담은 간이 PDF 만들기 (점검·연습용)
=====================================================================
    python tools/make_sample_pdf.py                       → samples/회사보고자료_예시.pdf
    python tools/make_sample_pdf.py 원본.pptx 결과.pdf     → 지정한 파일

파워포인트를 PDF 로 저장한 파일과 같은 구조(쪽마다 글상자·글자 좌표)를 가진 **글자만 있는** PDF 를 만듭니다.
tests/test_pdf.py 가 PDF 경로를 점검할 때 이 파일을 쓰고, 없으면 스스로 만듭니다.

- 저장소에 PDF 를 넣지 않기 위한 도구입니다 (GitHub 웹으로는 이진 파일을 올리기 어려워서). 파이썬 표준 라이브러리 + python-pptx 만 씁니다.
- 글꼴은 Adobe 표준 한국어 CID 글꼴(HYSMyeongJo-Medium)을 **파일에 넣지 않고 이름만** 적습니다. Acrobat·Edge·Chrome·미리보기는
  시스템의 한글 글꼴로 대신 보여 줍니다. 모양은 소박하지만 글자·좌표는 정확해서 점검용으로 충분합니다.
- 실제 연습은 PowerPoint 에서 samples/회사보고자료_예시.pptx 를 "다른 이름으로 저장 → PDF" 로 만든 파일이 가장 실전과 같습니다.
- 그 글꼴에 없는 글자(가운뎃점 ·)는 비슷한 글자로 바꿔 넣습니다 (SUBST).
- build(..., highlights=[...]) 로 "작성자가 칠해 둔 형광펜" 을 흉내 낼 수 있습니다. 파워포인트가 PDF 로 저장할 때처럼
  글줄 뒤에 색 네모를 그려 넣습니다 (tests/test_prehl.py 가 원본 형광펜 정리를 점검할 때 씀).
      {"slide": 2, "para": 5, "color": "FFFF00"}                       글줄에 딱 붙은 형광펜
      {"slide": 2, "para": 7, "color": "000080", "text": "FFFFFF"}     진한 형광펜 + 흰 글자
      {"slide": 2, "para": 9, "color": "DEEBF7", "pad": 8}             여백을 둔 채움 (표 칸·도형 — 형광펜이 아님)
"""
from __future__ import annotations

import sys
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import extract  # noqa: E402  (python-pptx 로 문단 읽기)

PAGE_W, PAGE_H = 960.0, 540.0        # 파워포인트 16:9 (13.333 × 7.5 인치)
MARGIN = 40.0
FONT = "HYSMyeongJo-Medium"          # Adobe-Korea1 표준 글꼴 — 뷰어가 시스템 글꼴로 대신 보여 줌
SUBST = {"·": "•"}         # UniKS-UCS2-H 에 없는 글자 → 비슷한 글자 (가운뎃점 → 불릿)


def _w(ch: str, size: float) -> float:
    """글자 폭 — /W 에 적는 값과 같게: ASCII 500, 나머지 1000 (1000 = 글꼴 크기)."""
    return size * (0.5 if ord(ch) < 0x80 else 1.0)


def wrap(text: str, size: float, maxw: float) -> list[str]:
    """줄바꿈: 공백이 있으면 공백에서, 없으면(한글) 글자 단위로 꺾는다."""
    lines: list[str] = []
    cur = ""
    cur_w = 0.0
    for ch in text:
        w = _w(ch, size)
        if cur and cur_w + w > maxw:
            cut = cur.rfind(" ")
            if cut > len(cur) * 0.5:                  # 뒤쪽에 공백이 있으면 거기서 끊는다
                lines.append(cur[:cut])
                cur = cur[cut + 1:]
            else:
                lines.append(cur)
                cur = ""
            cur_w = sum(_w(c, size) for c in cur)
        cur += ch
        cur_w += w
    if cur:
        lines.append(cur)
    return lines


def _hex(s: str) -> str:
    return "".join(f"{ord(c):04X}" for c in s)


def layout(pptx: Path) -> tuple[int, list[list[tuple[float, float, float, list[str]]]]]:
    """슬라이드마다 (x, 윗선 y, 글꼴 크기, 줄 목록) 의 문단 목록. 한 장이 한 쪽에 다 들어가도록 글꼴 크기를 고른다."""
    deck = extract.extract(str(pptx))
    pages = []
    for slide_no in range(1, deck.slide_count + 1):
        segs = [s for s in deck.segments
                if s.slide_no == slide_no and s.kind in ("title", "body", "table") and s.text.strip()]
        items: list[tuple[float, float, float, list[str]]] = []
        for size in (16.0, 14.0, 12.0, 11.0, 10.0, 9.0, 8.0, 7.0, 6.0):
            y = PAGE_H - MARGIN
            items = []
            fits = True
            for s in segs:
                fs = size + 4 if s.kind == "title" else size
                text = "".join(SUBST.get(c, c) for c in s.text)
                lines = wrap(text, fs, PAGE_W - 2 * MARGIN)
                h = len(lines) * fs * 1.35
                if y - h < MARGIN:
                    fits = False
                    break
                items.append((MARGIN, y, fs, lines))
                y -= h + fs * 0.7                    # 문단 사이 간격 (pdfminer 가 다른 글상자로 나누도록 넉넉히)
            if fits:
                break
        pages.append(items)
    return deck.slide_count, pages


def _rgb(hexstr: str) -> str:
    return " ".join(f"{int(hexstr[i:i + 2], 16) / 255:.4f}" for i in (0, 2, 4))


def build(pptx: Path, out: Path, highlights: list[dict] | None = None) -> tuple[int, int]:
    """PDF 파일을 쓴다. (쪽 수, 문단 수) 를 돌려준다. highlights 는 파일 머리말 참고."""
    n_pages, pages = layout(pptx)
    marks = {(h["slide"], h["para"]): h for h in (highlights or [])}
    objs: list[bytes] = []

    def add(body: bytes) -> int:
        objs.append(body)
        return len(objs)

    add(b"")                                            # 1: Catalog (나중에 채움)
    add(b"")                                            # 2: Pages
    fd = add(f"<< /Type /FontDescriptor /FontName /{FONT} /Flags 4 /FontBBox [-92 -250 1010 898] /ItalicAngle 0 "
             f"/Ascent 880 /Descent -120 /CapHeight 720 /StemV 80 >>".encode())
    cid = add(f"<< /Type /Font /Subtype /CIDFontType0 /BaseFont /{FONT} "
              f"/CIDSystemInfo << /Registry (Adobe) /Ordering (Korea1) /Supplement 1 >> "
              f"/FontDescriptor {fd} 0 R /DW 1000 /W [1 95 500] >>".encode())
    font = add(f"<< /Type /Font /Subtype /Type0 /BaseFont /{FONT}-UniKS-UCS2-H /Encoding /UniKS-UCS2-H "
               f"/DescendantFonts [{cid} 0 R] >>".encode())
    page_ids = []
    n_paras = 0
    for page_no, items in enumerate(pages, 1):
        # 색은 파워포인트·macOS 가 쓰는 방식(색 공간 지정 뒤 sc)으로 적고 글자 덩어리는 q … Q 로 감싼다 —
        # 저장·복원 뒤에도 형광펜 색을 제대로 읽는지(pdfdoc._Interpreter) 점검할 수 있게
        parts = ["/DeviceRGB cs 0 0 0 sc"]
        for k, (x, y, fs, lines) in enumerate(items, 1):
            n_paras += 1
            h = marks.get((page_no, k))
            if h:
                pad = float(h.get("pad", 0))
                parts.append(f"{_rgb(h['color'])} sc")
                for i, line in enumerate(lines):
                    base = y - fs - i * fs * 1.35                      # 이 줄의 글자 밑선
                    w = sum(_w(c, fs) for c in line)
                    parts.append(f"{x - pad:.2f} {base - 0.22 * fs - pad:.2f} {w + 2 * pad:.2f} {1.32 * fs + 2 * pad:.2f} re f")
            parts.append("q")
            parts.append(f"{_rgb(h['text']) if h and h.get('text') else '0 0 0'} sc")
            parts.append(f"BT /F1 {fs:g} Tf {fs * 1.35:g} TL 1 0 0 1 {x:g} {y - fs:g} Tm")
            for i, line in enumerate(lines):
                parts.append(("" if i == 0 else "T* ") + f"<{_hex(line)}> Tj")
            parts.append("ET")
            parts.append("Q")
        data = zlib.compress("\n".join(parts).encode("ascii"))
        cs = add(b"<< /Length " + str(len(data)).encode() + b" /Filter /FlateDecode >>\nstream\n" + data + b"\nendstream")
        pg = add(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_W:g} {PAGE_H:g}] "
                 f"/Resources << /Font << /F1 {font} 0 R >> >> /Contents {cs} 0 R >>".encode())
        page_ids.append(pg)
    objs[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[1] = (f"<< /Type /Pages /Count {len(page_ids)} /Kids [" + " ".join(f"{p} 0 R" for p in page_ids) + "] >>").encode()

    buf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for i, body in enumerate(objs, 1):
        offsets.append(len(buf))
        buf += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(buf)
    buf += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        buf += f"{off:010d} 00000 n \n".encode()
    buf += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(bytes(buf))
    return n_pages, n_paras


def main() -> None:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "samples" / "회사보고자료_예시.pptx"
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "samples" / "회사보고자료_예시.pdf"
    n_pages, n_paras = build(src, dst)
    print(f"저장: {dst}  ({n_pages}쪽 · 문단 {n_paras}개 · {dst.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
