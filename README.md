# 특허 마킹 도구 (온디바이스 · 오프라인)

회사 보고자료(PPTX)를 읽어 **특허 출원 검토가 필요한 구간에 형광펜과 `【특허검토필요】` 표시**를 넣어 돌려주는 도구.
AI(Qwen3)를 쓰지만 **내 컴퓨터 안에서만** 돌고, 문서는 외부로 나가지 않습니다.

> **Windows 로 옮겨서 설치하려면 → [docs/00_Windows_따라하기_가이드.md](docs/00_Windows_따라하기_가이드.md) (이 문서대로만 하면 됩니다)**
> 처음이라면 → [docs/01_프로그램_구조.md](docs/01_프로그램_구조.md) (구조 설명) → [docs/02_설치_및_이전_안내.md](docs/02_설치_및_이전_안내.md) (설치·이전)
> 발표용 → [docs/특허마킹도구_구조설명.pptx](docs/특허마킹도구_구조설명.pptx) (15장)   ·   결과 모양 미리 보기 → [samples/회사보고자료_예시_마킹결과.pptx](samples/회사보고자료_예시_마킹결과.pptx)

---

## 빠른 시작 (Windows)

0. GitHub 에서 받은 경우: 초록 **Code → Download ZIP** → `C:\` 에 풀고 폴더명을 `patent_marker` 로 변경.
   (저장소에는 15GB `offline_bundle` 이 없으므로 아래 `setup.bat` 온라인 설치를 사용합니다)
1. `setup.bat` 더블클릭 ─ 처음 한 번. Python·Ollama·모델(9GB)을 자동 설치합니다. (인터넷 필요)
   - 인터넷이 없는 PC 라면 `setup_offline.bat` (미리 만든 `offline_bundle/` 필요)
2. `run.bat` 더블클릭 ─ 브라우저가 열립니다. PPTX 를 올리고 **분석 시작**.
   - GPU 가 없거나 메모리 16GB 이하면 `run_8b.bat` (가벼운 qwen3:8b 로 실행)
3. 끝나면 **마킹된 PPTX 내려받기**.

## 빠른 시작 (macOS)

```bash
bash setup.sh
```

```bash
bash run.sh
```

---

## 결과물에 들어가는 것

| 표시 | 모양 | 뜻 |
|---|---|---|
| 형광펜 + 밑줄 | 노랑(A) / 하늘(B) / 살구(⚠공개) | 특허 단서 구간 |
| `【특허검토필요】` | 형광펜 바로 뒤, 빨간 작은 글씨 | 인쇄·흑백에서도 보이는 표시 |
| 배지 | 슬라이드 오른쪽 위 "특허검토필요 N건" | 넘겨 보며 한눈에 파악 |
| 발표자 노트 | 슬라이드별 [등급·유형] "인용구" → 근거 | 판정 이유 |
| 요약 슬라이드 | 문서 맨 끝 | 전체 목록·집계 |

등급: **A** 구체적 기술 수단이 드러남(즉시 출원 검토) · **B** 수단은 미기재이나 존재가 시사됨(발명자 인터뷰) · **⚠** 전시·논문 등 공개 리스크

---

## 폴더 구성

```
app/        프로그램 본체 (config → extract → lexicon → analyze → merge → mark, main)
docs/       구조 설명 · 설치/이전 안내 · 발표자료(pptx)
tests/      점검 스크립트 (모델 없이 돌릴 수 있는 것 포함)
samples/    연습용 보고서 · 실제 모델로 마킹한 결과 예시
tools/      오프라인 설치 꾸러미 만들기
.vscode/    VS Code 설정 (작업 메뉴 · F5 실행)
setup.*     처음 한 번 설치     run.*   실행 (run_8b.bat = 가벼운 모델)     setup_offline.*   인터넷 없이 설치
```

---

## 자주 바꾸는 것

| 바꾸고 싶은 것 | 파일 | 항목 |
|---|---|---|
| 모델 (가볍게) | `app/config.py` | `MODEL = "qwen3:8b"` |
| 표시 문구 | `app/config.py` | `TAG_TEXT` |
| 형광펜 색 | `app/config.py` | `GRADE_COLOR` |
| 단서 표현 추가 (업종 바뀔 때) | `app/lexicon.py` | `RULES`, `NOISE` |
| 판정 기준 | `app/analyze.py` | `SYSTEM` |

---

## 점검

```bash
.venv/bin/python tests/test_filter.py     # 1초  · 오탐 필터
.venv/bin/python tests/smoke.py           # 몇 초 · 모델 없이 읽기→마킹
.venv/bin/python tests/mock_llm.py        # 몇 초 · 가짜 모델로 연동 경로
.venv/bin/python tests/e2e.py             # 수 분 · 진짜 모델
```
(Windows 는 `.venv\Scripts\python` 로 바꿔 실행)

---

## 한계

- 변리사 검토 전 **1차 스크리닝**입니다. 선행기술 조사, 신규성·진보성 판단은 하지 않습니다.
- 슬라이드 안 **이미지 속 글자**는 읽지 못합니다 (OCR 없음).
- 차트 내부 글자는 문맥 참고용으로만 읽고 칠하지 않습니다.
- `.ppt`(구형)는 `.pptx` 로 변환 후 사용.
- 형광펜은 **PowerPoint 2016 이상**에서 표시됩니다. 다른 뷰어에서는 밑줄과 `【특허검토필요】` 문구로 확인하세요.
- 재현율을 높게 잡아서 매출·이익률 같은 경영 수치가 가끔 후보로 올라옵니다. `lexicon.NOISE` 로 눌러 두었지만 업종 용어가 다르면 보강이 필요합니다.
