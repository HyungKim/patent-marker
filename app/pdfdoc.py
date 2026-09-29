"""
pdfdoc.py ─ PDF 보고서(파워포인트를 PDF 로 변환한 자료) 읽기 · 마킹 · 검토본 읽기
=====================================================================

[이 파일이 하는 일]
  PPTX 와 똑같은 흐름(읽기 → 규칙 사전 → 모델 판정 → 병합 → 마킹 → 검토 반영)을 PDF 에도 적용합니다.
  글자를 뽑는 데는 pdfminer.six, 형광펜을 넣고 읽는 데는 pypdf 를 씁니다 — 둘 다 순수 파이썬이라
  Windows 에서도 pip 로 그대로 설치됩니다.

      extract(path)        PDF → PdfDeck (문단 목록 + 글자마다의 좌표)   ← extract.extract() 가 .pdf 면 여기로 넘김
      apply_and_save(...)  판정 결과를 **형광펜 주석**(Highlight annotation)으로 넣어 저장 + 1쪽 범례 메모
      read_reviewed(path)  검토자가 형광펜을 고쳐 저장한 PDF 에서 형광펜 구간·색을 읽음 (review.read_reviewed 가 .pdf 면 여기로)

[PPTX 와 다른 점]
  · PDF 는 글자 자체를 고칠 수 없으므로 '【출원검토필요】 문구 삽입' 옵션은 적용되지 않습니다 (형광펜 + 메모만).
  · 형광펜은 PDF 주석이라 Acrobat Reader·Edge·Chrome 등 어느 뷰어에서나 보이고(모양 스트림을 함께 넣어 미리보기 창에서도 보임), 검토자는 뷰어의 형광펜 도구로
    추가하거나(누락) 주석을 지워서(오탐) 검토합니다. 색 규약은 PPTX 와 같습니다: 노랑 = A, 청록/하늘/파랑 = B, 빨강/분홍 = C(공개).
  · 각 형광펜 주석의 '메모' 에 등급·분류·사유가 들어 있어, 뷰어의 주석 목록에서 바로 읽을 수 있습니다.
  · 문단 단위는 PDF 의 글상자(pdfminer 의 TextBox) 입니다. PPT 문단 하나가 보통 글상자 하나로 나옵니다.
  · 회전된 페이지(/Rotate)는 좌표가 어긋날 수 있습니다 — PPT 변환본은 회전이 없어 해당하지 않습니다.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTAnno, LTChar, LTTextBox, LTTextLine
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage
from pypdf import PdfReader, PdfWriter
from pypdf.annotations import Highlight, Rectangle, Text
from pypdf.generic import ArrayObject, DictionaryObject, FloatObject, NameObject, NumberObject, StreamObject, TextStringObject

from . import config, prehl
from .extract import Segment

logging.getLogger("pdfminer").setLevel(logging.ERROR)     # 글꼴 정보 경고가 검은 창을 덮지 않게

# 글자를 문단으로 묶는 기준. word_margin 은 글자 사이 틈이 이만큼(글자 폭 비율) 넘으면 공백으로 보는 값 —
# 한글은 자간이 좁아 0.15 면 헛공백이 거의 안 생긴다 (예시 문서로 확인: PPTX 문단 106개 전부 일치).
LAP = LAParams(char_margin=2.0, line_margin=0.5, word_margin=0.15, boxes_flow=0.5)
_HANGUL = re.compile(r"[가-힣]")


@dataclass
class PdfDeck:
    """PDF 한 권 = 경로 + 문단 목록 + 문단마다 글자별 좌표."""

    path: str
    page_count: int
    segments: list[Segment]
    boxes: dict[int, list] = field(default_factory=dict)     # seg_id → [(x0,y0,x1,y1) 또는 None] (글자마다)
    pages: dict[int, tuple[float, float]] = field(default_factory=dict)   # 쪽 → (너비, 높이)
    is_pdf: bool = True
    baked: dict[int, list] = field(default_factory=dict)     # 쪽 → 덮을 '굳은 형광펜' 네모 목록 (prehl.find_baked)
    baked_kept: int = 0                                       # 형광펜 같지만 확신이 없어 그대로 둘 개수
    prehl: dict = field(default_factory=dict)                 # 원본 형광펜 정리 결과

    @property
    def slide_count(self) -> int:
        return self.page_count

    def by_slide(self, slide_no: int) -> list[Segment]:
        return [s for s in self.segments if s.slide_no == slide_no]

    def get(self, seg_id: int) -> Segment | None:
        for s in self.segments:
            if s.seg_id == seg_id:
                return s
        return None


# ═════════════════════════════════════════════════════════════════
# 1. 읽기
# ═════════════════════════════════════════════════════════════════
def _join_lines(lines: list[list[tuple[str, tuple | None]]]) -> list[tuple[str, tuple | None]]:
    """글상자 안의 줄들을 하나로 잇는다. 한글끼리 이어지는 줄바꿈은 공백 없이(단어 중간에서 꺾인 것),
    그 밖(영문·숫자)은 공백 하나로 잇는다."""
    out: list[tuple[str, tuple | None]] = []
    for line in lines:
        if not line:
            continue
        if out:
            prev, nxt = out[-1][0], line[0][0]
            if not (_HANGUL.match(prev) and _HANGUL.match(nxt)) and not prev.isspace():
                out.append((" ", None))
        out.extend(line)
    return out


class _Interpreter(PDFPageInterpreter):
    """pdfminer 의 그림 상태 저장(q)이 색 공간을 빠뜨려, 복원(Q) 뒤의 색이 마지막 성분 하나로만 읽힌다
    (파워포인트·macOS 가 만든 PDF 에서 노랑 1 1 0 이 0 으로 읽힘). 색 공간도 함께 저장하도록 고친다.
    글자 추출에는 영향이 없고, 형광펜 색·글자색을 읽을 때만 쓰인다."""

    def get_current_state(self):
        ctm, textstate, gstate = super().get_current_state()
        gstate.scs = self.graphicstate.scs
        gstate.ncs = self.graphicstate.ncs
        return ctm, textstate, gstate


def _layout_pages(path: str):
    """쪽마다 pdfminer 배치 결과(LTPage)를 내놓는다. high_level.extract_pages 와 같되 위의 해석기를 쓴다."""
    with open(path, "rb") as fp:
        rsrc = PDFResourceManager(caching=True)
        device = PDFPageAggregator(rsrc, laparams=LAP)
        interp = _Interpreter(rsrc, device)
        for page in PDFPage.get_pages(fp):
            interp.process_page(page)
            yield device.get_result()


def extract(path: str, detect_baked: bool = True) -> PdfDeck:
    """PDF 파일 경로 → PdfDeck. extract.extract() 와 같은 모양의 문단 목록을 만든다.
    detect_baked 면 쪽 내용에 그림으로 굳은 형광펜도 함께 찾아 둔다 (덮는 것은 저장할 때)."""
    segments: list[Segment] = []
    boxes: dict[int, list] = {}
    pages: dict[int, tuple[float, float]] = {}
    baked: dict[int, list] = {}
    baked_kept = 0
    n = 0
    page_no = 0
    for page_no, page in enumerate(_layout_pages(str(path)), 1):
        pages[page_no] = (float(page.width), float(page.height))
        if detect_baked:
            try:
                cover, kept = prehl.find_baked(page)
            except Exception:  # noqa: BLE001  (형광펜 찾기가 실패해도 분석은 계속)
                cover, kept = [], 0
            if cover:
                baked[page_no] = cover
            baked_kept += kept
        box_no = 0
        for el in page:
            if not isinstance(el, LTTextBox):
                continue
            lines: list[list[tuple[str, tuple | None]]] = []
            for line in el:
                if not isinstance(line, LTTextLine):
                    continue
                cur: list[tuple[str, tuple | None]] = []
                for ch in line:
                    if isinstance(ch, LTChar):
                        cur.append((ch.get_text(), (float(ch.x0), float(ch.y0), float(ch.x1), float(ch.y1))))
                    elif isinstance(ch, LTAnno):
                        t = ch.get_text()
                        if t == "\n":
                            continue                       # 줄 끝 — _join_lines 가 처리
                        cur.append((t, None))
                if cur:
                    lines.append(cur)
            chars = _join_lines(lines)
            # 앞뒤 공백 정리 (좌표 목록도 같이 잘라 글자 번호가 어긋나지 않게)
            while chars and chars[0][0].isspace():
                chars.pop(0)
            while chars and chars[-1][0].isspace():
                chars.pop()
            text = "".join(c for c, _ in chars)
            if not text.strip():
                continue
            box_no += 1
            n += 1
            segments.append(Segment(seg_id=n, slide_no=page_no, kind="body",
                                    addr=f"p{page_no}/글상자{box_no}", text=text, markable=True))
            boxes[n] = [b for _, b in chars]
    return PdfDeck(path=str(path), page_count=page_no, segments=segments, boxes=boxes, pages=pages,
                   baked=baked, baked_kept=baked_kept)


# ═════════════════════════════════════════════════════════════════
# 2. 마킹 — 형광펜 주석 + 1쪽 범례 메모
# ═════════════════════════════════════════════════════════════════
def _line_rects(char_boxes: list, span: tuple[int, int] | None) -> list[tuple[float, float, float, float]]:
    """구간의 글자 좌표를 줄 단위 네모로 묶는다 (같은 줄 = y 가 2pt 안에서 같음)."""
    lo, hi = span if span else (0, len(char_boxes))
    rects: list[list[float]] = []
    for b in char_boxes[lo:hi]:
        if b is None:
            continue
        x0, y0, x1, y1 = b
        if rects and abs(rects[-1][1] - y0) <= 2.0 and x0 >= rects[-1][0] - 1.0:
            r = rects[-1]
            r[0], r[1], r[2], r[3] = min(r[0], x0), min(r[1], y0), max(r[2], x1), max(r[3], y1)
        else:
            rects.append([x0, y0, x1, y1])
    return [(r[0], r[1] - 1.0, r[2], r[3] + 1.0) for r in rects]


def _quads(rects) -> ArrayObject:
    """네모 목록 → QuadPoints (네모마다 좌상·우상·좌하·우하 8개 숫자)."""
    arr = ArrayObject()
    for x0, y0, x1, y1 in rects:
        for v in (x0, y1, x1, y1, x0, y0, x1, y0):
            arr.append(FloatObject(v))
    return arr


def _union(rects) -> tuple[float, float, float, float]:
    return (min(r[0] for r in rects), min(r[1] for r in rects), max(r[2] for r in rects), max(r[3] for r in rects))


def _rgb(hexv: str) -> tuple[float, float, float]:
    return tuple(int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _form(writer: PdfWriter, bbox, content: str, multiply: bool = False) -> DictionaryObject:
    """주석의 '모양 스트림'(/AP) 을 만든다.

    Acrobat·Edge·Chrome 은 모양이 없어도 형광펜을 스스로 그리지만, 탐색기 미리보기·일부 구형 뷰어는 /AP 가 없으면
    아무것도 보여 주지 않는다. 그래서 색칠 명령을 직접 넣어 어느 뷰어에서나 같은 모양이 나오게 한다.
    multiply=True 면 곱하기 혼합(Multiply)으로 칠해 글자가 형광펜 밑에서도 또렷이 보인다 (Acrobat 형광펜과 같은 방식).
    """
    st = StreamObject()
    st[NameObject("/Type")] = NameObject("/XObject")
    st[NameObject("/Subtype")] = NameObject("/Form")
    st[NameObject("/BBox")] = ArrayObject([FloatObject(v) for v in bbox])
    if multiply:
        gs = DictionaryObject({NameObject("/Type"): NameObject("/ExtGState"), NameObject("/BM"): NameObject("/Multiply")})
        st[NameObject("/Resources")] = DictionaryObject({NameObject("/ExtGState"): DictionaryObject({NameObject("/GS"): gs})})
    st.set_data(content.encode("latin-1"))
    return DictionaryObject({NameObject("/N"): writer._add_object(st)})


def _f(v: float) -> str:
    return f"{v:.2f}"


def apply_and_save(deck: PdfDeck, findings: list, marks: list, dst: Path, tag_marks: bool = False,
                   strip_highlights: bool = False) -> dict:
    """판정 결과를 형광펜 주석으로 넣고 dst 에 저장한다. (tag_marks 는 PDF 에선 적용 불가 — 메모로 대신)

    strip_highlights 면 먼저 원본의 형광펜 주석을 지우고, 쪽에 굳은 형광펜은 바탕색으로 덮는다 (prehl.clean_pdf).
    """
    reader = PdfReader(deck.path)
    writer = PdfWriter(clone_from=reader)
    deck.prehl = prehl.clean_pdf(writer, deck.baked, deck.baked_kept, strip_highlights, config.PDF_COVER_BAKED)
    seg_map = {s.seg_id: s for s in deck.segments}
    by_seg: dict[int, list] = {}
    for f in findings:
        by_seg.setdefault(f.seg_id, []).append(f)

    painted = 0
    for m in marks:
        seg = seg_map.get(m.seg_id)
        if seg is None:
            continue
        rects = _line_rects(deck.boxes.get(m.seg_id, []), m.span)
        if not rects:
            continue
        grade = "C" if m.disclosure_risk else m.grade
        box = _union(rects)
        annot = Highlight(rect=box, quad_points=_quads(rects),
                          highlight_color=config.GRADE_COLOR[grade], title_bar="출원검토필요")
        r, g, b = _rgb(config.GRADE_COLOR[grade])
        annot[NameObject("/AP")] = _form(writer, box, f"/GS gs {_f(r)} {_f(g)} {_f(b)} rg " + " ".join(
            f"{_f(x0)} {_f(y0)} {_f(x1 - x0)} {_f(y1 - y0)} re f" for x0, y0, x1, y1 in rects), multiply=True)
        annot[NameObject("/F")] = NumberObject(4)          # 인쇄에도 나오게
        # 메모: 이 구간에 걸린 후보들의 등급·분류·사유 — 뷰어의 주석 목록에서 바로 읽힌다
        notes = []
        for f in by_seg.get(m.seg_id, []):
            if m.span is None or f.span is None or not (f.span[1] <= m.span[0] or f.span[0] >= m.span[1]):
                g = "C" if f.disclosure_risk else f.grade
                notes.append(f"[{g}] {config.GRADE_LABEL[g]} · {f.category}" + (f" · {f.reason}" if f.reason else ""))
        annot[NameObject("/Contents")] = TextStringObject("출원검토필요\n" + "\n".join(dict.fromkeys(notes)))
        writer.add_annotation(page_number=seg.slide_no - 1, annotation=annot)
        painted += 1

    legend = 0
    if findings and config.SLIDE_LEGEND and deck.page_count:
        w, h = deck.pages.get(1, (595.0, 842.0))
        x, y = 10.0, h - 10.0
        note = Text(rect=(x, y - 20, x + 20, y), open=False,
                    text=(f"{config.LEGEND_TITLE}\n"
                          f"노랑 = A {config.GRADE_LABEL['A']}\n"
                          f"청록 = B {config.GRADE_LABEL['B']}\n"
                          f"빨강 = C {config.GRADE_LABEL['C']}\n"
                          "검토: 형광펜을 지우면 '아님', 새로 칠하면 '추가' 로 반영됩니다."),
                    title_bar=config.LEGEND_NAME)
        # 메모 아이콘 모양: 노란 쪽지 + 글줄 세 개 (아이콘을 스스로 그리지 않는 뷰어용)
        nx, ny = x, y - 20
        note[NameObject("/AP")] = _form(writer, (nx, ny, nx + 20, ny + 20), (
            f"1.00 0.85 0.30 rg 0.30 0.30 0.30 RG 1 w {_f(nx + 0.5)} {_f(ny + 0.5)} 19 19 re B "
            f"1.2 w {_f(nx + 4)} {_f(ny + 14)} m {_f(nx + 16)} {_f(ny + 14)} l S "
            f"{_f(nx + 4)} {_f(ny + 10)} m {_f(nx + 16)} {_f(ny + 10)} l S "
            f"{_f(nx + 4)} {_f(ny + 6)} m {_f(nx + 12)} {_f(ny + 6)} l S"))
        note[NameObject("/F")] = NumberObject(4)
        writer.add_annotation(page_number=0, annotation=note)
        # 색 견본 세 칸 (뷰어 종류와 무관하게 보이는 네모 주석) — 메모에 등급 이름
        for i, g in enumerate(("A", "B", "C")):
            rx = x + 26 + i * 22
            box = (rx, y - 16, rx + 18, y - 2)
            sq = Rectangle(rect=box, interior_color=config.GRADE_COLOR[g], title_bar=config.LEGEND_NAME)
            sq[NameObject("/Contents")] = TextStringObject(f"{g} · {config.GRADE_LABEL[g]}")
            sq[NameObject("/C")] = ArrayObject([FloatObject(0.4)] * 3)
            r, gg, b = _rgb(config.GRADE_COLOR[g])
            sq[NameObject("/AP")] = _form(writer, box, (
                f"{_f(r)} {_f(gg)} {_f(b)} rg 0.40 0.40 0.40 RG 1 w "
                f"{_f(box[0] + 0.5)} {_f(box[1] + 0.5)} {_f(box[2] - box[0] - 1)} {_f(box[3] - box[1] - 1)} re B"))
            sq[NameObject("/F")] = NumberObject(4)
            writer.add_annotation(page_number=0, annotation=sq)
        legend = 1

    dst.parent.mkdir(parents=True, exist_ok=True)
    with open(dst, "wb") as fh:
        writer.write(fh)
    return {"runs_painted": painted, "marks": len(marks), "tags": 0, "legend": legend, "prehl": deck.prehl}


# ═════════════════════════════════════════════════════════════════
# 3. 검토본 읽기 — 형광펜 주석 → 문단별 구간·색
# ═════════════════════════════════════════════════════════════════
def _hex(color) -> str:
    try:
        vals = [float(v) for v in color]
    except Exception:  # noqa: BLE001
        return config.GRADE_COLOR["A"]
    if len(vals) == 3:
        return "".join(f"{max(0, min(255, round(v * 255))):02X}" for v in vals)
    if len(vals) == 1:                     # 회색조
        g = max(0, min(255, round(vals[0] * 255)))
        return f"{g:02X}{g:02X}{g:02X}"
    return config.GRADE_COLOR["A"]


def _highlights(reader: PdfReader) -> dict[int, list[tuple[list, str]]]:
    """쪽마다 [(네모 목록, 색 hex), ...]. 형광펜(Highlight) 주석만."""
    out: dict[int, list] = {}
    for pno, page in enumerate(reader.pages, 1):
        for a in page.get("/Annots") or []:
            try:
                obj = a.get_object()
            except Exception:  # noqa: BLE001
                continue
            if obj.get("/Subtype") != "/Highlight":
                continue
            q = obj.get("/QuadPoints")
            rects = []
            if q:
                vals = [float(v) for v in q]
                for i in range(0, len(vals) - 7, 8):
                    xs, ys = vals[i:i + 8:2], vals[i + 1:i + 8:2]
                    rects.append((min(xs), min(ys), max(xs), max(ys)))
            elif obj.get("/Rect"):
                r = [float(v) for v in obj["/Rect"]]
                rects.append((min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3])))
            if rects:
                out.setdefault(pno, []).append((rects, _hex(obj.get("/C") or [])))
    return out


def read_reviewed(path: str) -> dict:
    """검토완료 PDF → review.read_reviewed() 와 같은 모양: {fingerprint, slide_count, paras:[{slide_no, text, spans, tags}]}"""
    from .review import fingerprint_texts

    deck = extract(path, detect_baked=False)
    hl = _highlights(PdfReader(path))
    paras = []
    for seg in deck.segments:
        spans = []
        char_boxes = deck.boxes.get(seg.seg_id, [])
        for rects, color in hl.get(seg.slide_no, []):
            hit = [False] * len(char_boxes)
            for i, b in enumerate(char_boxes):
                if b is None:
                    continue
                cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2
                if any(r[0] - 1 <= cx <= r[2] + 1 and r[1] - 1 <= cy <= r[3] + 1 for r in rects):
                    hit[i] = True
            # 걸린 글자들을 연속 구간으로 (공백·줄바꿈처럼 좌표 없는 글자는 사이에 끼어 있어도 이어진 것으로)
            i = 0
            while i < len(hit):
                if not hit[i]:
                    i += 1
                    continue
                j = i
                while j < len(hit) and (hit[j] or char_boxes[j] is None):
                    j += 1
                while j > i and (char_boxes[j - 1] is None or not hit[j - 1]):
                    j -= 1
                if j - i >= 1 and seg.text[i:j].strip():
                    spans.append({"start": i, "end": j, "color": color})
                i = j + 1
        spans.sort(key=lambda s: s["start"])
        paras.append({"slide_no": seg.slide_no, "text": seg.text, "spans": spans, "tags": []})
    return {"fingerprint": fingerprint_texts([s.text for s in deck.segments]),
            "slide_count": deck.page_count, "paras": paras}
