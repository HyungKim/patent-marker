"""
pickdialog.py ─ 파일 열기 창을 띄우는 작은 보조 프로그램
=====================================================================

local.pick_files() 가 별도 프로세스로 실행합니다. (서버 프로세스 안에서 창을 띄우면
서버가 멈추기 때문에 따로 띄웁니다.)

    python app/pickdialog.py single [시작폴더]   → 파일 하나
    python app/pickdialog.py multi  [시작폴더]   → 여러 개

고른 경로들을 JSON 목록으로 한 줄 출력합니다. 취소하면 []  (ASCII 만 써서 인코딩 문제가 없습니다).
tkinter 가 없는 Python 이면 ImportError 로 끝나고, 부르는 쪽이 안내문을 보여 줍니다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


def main() -> int:
    import tkinter as tk                       # 표준 라이브러리 (Windows python.org 설치판에 포함)
    from tkinter import filedialog

    multiple = len(sys.argv) > 1 and sys.argv[1] == "multi"
    start = sys.argv[2] if len(sys.argv) > 2 and Path(sys.argv[2]).is_dir() else str(Path.home())

    root = tk.Tk()
    root.withdraw()                            # 빈 기본 창은 숨기고 파일 창만 보인다
    root.attributes("-topmost", True)          # 브라우저 뒤로 숨지 않게 맨 앞으로
    root.update()

    kw = dict(parent=root, initialdir=start,
              filetypes=[("PowerPoint 파일", "*.pptx *.potx"), ("모든 파일", "*.*")])
    if multiple:
        picked = root.tk.splitlist(filedialog.askopenfilenames(title="검토완료 PPTX 선택 (여러 개 가능)", **kw))
    else:
        one = filedialog.askopenfilename(title="분석할 PPTX 선택", **kw)
        picked = [one] if one else []
    root.destroy()

    # Windows 의 tk 는 C:/a/b 처럼 슬래시로 주므로 Path 를 거쳐 C:\a\b 로 통일한다
    sys.stdout.write(json.dumps([str(Path(p)) for p in picked if p]) + "\n")
    sys.stdout.flush()
    return 0


if __name__ == "__main__":
    sys.exit(main())
