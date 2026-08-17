# Quick Start Guide

## 설치

```bash
pip install yfinance openpyxl --break-system-packages
```

## 사용 방법 (순서대로)

### 1️⃣ 기업명 → Ticker 확인

```bash
python scripts/ticker_lookup.py "Apple"
# 또는 매핑에 없는 기업은 Ticker 직접 입력
python scripts/ticker_lookup.py AAPL
```

### 2️⃣ 재무 데이터 수집(캐시)

```bash
python scripts/fetch_financials.py AAPL --period both --quarters 12 --years 5
python scripts/fetch_extra_info.py AAPL
```

### 3️⃣ 엑셀 생성

```bash
# 연간+분기 둘 다
python scripts/build_workbook.py AAPL "Apple Inc." --period both --outdir /mnt/user-data/outputs

# 연간만
python scripts/build_workbook.py AAPL "Apple Inc." --period annual --outdir /mnt/user-data/outputs

# 분기만
python scripts/build_workbook.py AAPL "Apple Inc." --period quarterly --outdir /mnt/user-data/outputs
```

### 4️⃣ 여러 기업 비교

```bash
# 비교 대상 전부 먼저 캐시 생성
python scripts/fetch_financials.py MSFT --period both
python scripts/fetch_financials.py GOOGL --period both

python scripts/build_comparison_workbook.py AAPL:Apple MSFT:Microsoft GOOGL:Google --outdir /mnt/user-data/outputs
```

## 입출력 예시

### 입력
```
"Apple 재무제표 뽑아줘"
"AAPL 년 분석해줘"
"테슬라 분기 분석해줘"
"Apple, Microsoft 비교해줘"
```

### 출력 (--period both 기준)
```
Apple Inc._20260817.xlsx
├─ 분기_재무제표   (보통 4~5개 분기)
├─ 연간_재무제표   (보통 4개년)
├─ 지표_분기       (11개 지표 + 차트 4개)
├─ 지표_연간       (11개 지표 + 차트 4개)
├─ 투자분석        (4개 섹션: 재무비율/위험신호/주가지표/간이투자판단)
└─ 원본데이터      (숨김, 수식 참조 소스)
```

## 지원 기업

### 매핑 목록에 있는 38개 기업(한글/영문 라벨 합쳐 83개)
```
빅테크: AAPL, MSFT, GOOGL, AMZN, META, TSLA, NVDA, INTC, AMD, NFLX, ADBE, CRM, ORCL, IBM, QCOM, AVGO
금융: JPM, BAC, WFC, GS, C, MS, V, MA, PYPL, BRK-B
소비재/기타: KO, PEP, MCD, NKE, SBUX, DIS, COST, WMT, JNJ, PFE, BA, XOM
```

### 매핑에 없는 기업
Ticker를 직접 입력하면 yfinance로 실존 여부를 확인합니다: `python scripts/ticker_lookup.py TICKER`

## 생성되는 지표 (11개)

- 매출액, 영업이익, 당기순이익 (금액)
- 매출총이익률(%), 영업이익률(%), 순이익률(%)
- ROE(%), ROA(%)
- 부채비율(%), 유동비율(%), 자기자본비율(%)

> v1.0.0 문서는 "19개 지표(건전성8+수익성6+성장성6+활동성6)"라고 광고했지만 실제 구현은 8개뿐이었습니다. v2.0.0은 실제 구현된 11개만 정직하게 안내합니다. 성장성·활동성 지표는 아직 없습니다 — 필요하면 `build_workbook.py`의 `INDICATOR_ROWS`에 추가합니다.

## 캐싱

```
cache/
├── company_{ticker}.json
├── financials_{ticker}_quarterly.json
├── financials_{ticker}_annual.json
└── extra_info_{ticker}.json
```

캐시가 있으면 `fetch_financials.py`/`fetch_extra_info.py`가 재조회 없이 그대로 재사용합니다(`--force`로 강제 재조회 가능). `build_workbook.py`는 이 캐시를 **실제로 읽습니다**(v1.0.0은 캐시를 무시하고 매번 새로 API를 불렀습니다).

## 트러블슈팅

| 문제 | 원인/해결 |
|---|---|
| Ticker를 찾을 수 없음 | 정확한 Ticker 심볼을 직접 입력(예: AAPL) |
| 재무제표 일부 계정이 비어있음 | yfinance가 그 계정을 아예 제공 안 하거나 필드명이 다를 수 있음. 지어내지 않고 빈 칸으로 둠 |
| 분기가 12개가 아니라 4~5개만 나옴 | yfinance 무료 데이터의 정상적인 제약입니다. 버그 아님 |
| `build_workbook.py` 실행 시 "캐시가 없습니다" 경고 | `fetch_financials.py`를 먼저 실행해야 함 |
| pip install 실패 | 이 환경(claude.ai 채팅)과 실제 실행 환경(Cowork)의 네트워크 설정이 다를 수 있음. Cowork에서 직접 실행 필요 |

## 다음 단계 (아직 안 만든 것)

1. ⏳ investment-thesis-writer 스타일 정성분석 통합
2. ⏳ 버핏멍거_가치평가 스타일 DCF/오너어닝 계산
3. ⏳ 성장성·활동성 지표 추가(현재 재무비율 11개만 구현)

더 자세한 정보는 `SKILL.md` 참고.
