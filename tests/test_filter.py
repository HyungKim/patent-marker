"""
tests/test_filter.py ─ 경영지표 오탐 필터 점검 (LLM 없이, 1초)
=====================================================================
실제 실행에서 모델이 "매출 비중 34% → 51%" 같은 재무 수치를 특허 후보로 올렸던
사례를 고정 케이스로 박아 두고, lexicon.is_business_noise() 가 이를 걸러내는지 확인합니다.

    실행:  python tests/test_filter.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import lexicon  # noqa: E402

# (문단 원문, 모델이 인용한 구간, 걸러져야 하는가)
CASES = [
    # 실행 로그에서 B 등급으로 새어 나왔던 경영 지표 → 걸러져야 함(True)
    ("전극 검사장비 매출 비중이 34% → 51%로 상승. 조립·물류 장비 중심 구조에서 "
     "검사 솔루션 중심으로 이동 중.",
     "전극 검사장비 매출 비중이 34% → 51%로 상승", True),
    ("2026년 상반기 영업이익률 12.7%로 개선. 개선분의 상당 부분은 슬라이드 4의 "
     "광학 모듈 원가 절감 효과에서 기인함.",
     "2026년 상반기 영업이익률 12.7%로 개선", True),
    ("장비 평균 판매단가 연 7~9% 하락. 하드웨어 원가 절감과 소프트웨어 차별화가 수익성의 관건.",
     "장비 평균 판매단가 연 7~9% 하락", True),

    # 같은 문단의 기술 문장은 살아남아야 함(False)
    ("전극 검사장비 매출 비중이 34% → 51%로 상승. 조립·물류 장비 중심 구조에서 "
     "검사 솔루션 중심으로 이동 중.",
     "조립·물류 장비 중심 구조에서 검사 솔루션 중심으로 이동 중", False),
    ("2026년 상반기 영업이익률 12.7%로 개선. 개선분의 상당 부분은 슬라이드 4의 "
     "광학 모듈 원가 절감 효과에서 기인함.",
     "광학 모듈 원가 절감 효과", False),

    # 순수 기술 문장은 어떤 경우에도 걸러지면 안 됨(False)
    ("450nm 청색광과 660nm 적색광을 200Hz로 교번 점등하고, 파장별 영상을 합성해 "
     "표면 결함과 내부 기공을 한 번의 스캔으로 동시 검출.",
     "450nm 청색광과 660nm 적색광을 200Hz로 교번 점등", False),
    ("검사 결과를 코터 슬롯다이 제어기에 200ms 주기로 반영하는 폐루프 구성. "
     "두께 편차 ±0.8㎛ → ±0.3㎛로 축소.",
     "두께 편차 ±0.8㎛ → ±0.3㎛로 축소", False),
    ("오검출률 4.2% → 0.7%", "4.2% → 0.7%", False),
]

fails = 0
for text, quote, expect in CASES:
    i = text.find(quote)
    assert i >= 0, f"인용구가 원문에 없음: {quote!r}"
    span = (i, i + len(quote))
    got = lexicon.is_business_noise(text, lexicon.scan(text), span)
    ok = got == expect
    fails += 0 if ok else 1
    print(f"{'PASS' if ok else 'FAIL'}  기대={expect!s:<5} 실제={got!s:<5} {quote[:44]!r}")

print(f"\n{len(CASES) - fails}/{len(CASES)} 통과")
sys.exit(1 if fails else 0)
