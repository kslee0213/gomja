# 파이프라인 아키텍처

## 단일 기업 흐름

```
사용자 요청
    ↓
ticker_lookup.py       (기업명 → Ticker → SEC CIK, SEC 공식 매핑 사용)
    ↓
fetch_financials.py    (SEC EDGAR CompanyFacts API → cache/secfacts_{ticker}.json, 원본 그대로)
fetch_extra_info.py    (yfinance 주가만 → cache/price_{ticker}.json)
    ↓
build_workbook.py      (secfacts 캐시에서 duration/instant 필터링해 원본데이터 시트 생성 →
                         나머지 시트는 전부 원본데이터를 참조하는 수식.
                         PER/PBR/PSR은 SEC 재무값 + 주가로 직접 수식 계산)
    ↓
엑셀 파일(6개 시트) 생성 완료
```

## SEC CompanyFacts 파싱 흐름 (핵심)

```
raw_facts (secfacts_{ticker}.json의 "raw" 필드)
    ↓
_collect_duration_dates()  — IS/CF에서 실제 "연간(FY)"/"분기" 날짜 집합을 먼저 구함
    ↓
extract_periods(key)       — 각 계정의 (end-start) 일수로 분기/연간 구분(duration),
                              또는 위에서 구한 날짜 집합에 매칭(instant/BS)
    ↓
build_frequency_payload()  — {sj: {key: {period: val}}} 형태로 재구성, 최근 N개만
```

⚠️ instant(BS) 계정을 duration 계정의 날짜 집합과 매칭시키지 않으면, 아직 10-K가 없는 진행 중인 분기말 스냅샷이 "연간" 쪽에 잘못 섞여 들어간다(실제로 재현된 버그, 수정 완료).

## 다중 기업 비교 흐름

```
여러 기업 요청
    ↓
[기업마다: fetch_financials.py(SEC) + fetch_extra_info.py(yfinance, 선택)]
    ↓
build_comparison_workbook.py
  (각 기업의 secfacts 캐시를 실제로 읽어 시계열별로 나란히 배치.
   "12분기 비교"와 "최근5년 비교"는 서로 다른 실제 데이터를 보여준다.)
    ↓
비교 엑셀 생성 완료
```

## 캐싱 구조

```
cache/
├── sec_company_tickers.json      (ticker_lookup.py — SEC 전체 티커→CIK 매핑, 하루 단위 재사용)
├── company_{ticker}.json         (ticker_lookup.py — 개별 기업 CIK)
├── secfacts_{ticker}.json        (fetch_financials.py — SEC CompanyFacts 원본 그대로)
└── price_{ticker}.json           (fetch_extra_info.py — yfinance 주가만)
```

## DART 버전과의 설계 차이

| 항목 | dart-kospi-financials | us-stocks-financials (v3.0.0) |
|---|---|---|
| 재무제표 소스 | DART Open API(공식, API 키 필요) | SEC EDGAR(공식, 키 불필요·이메일만) |
| 주가 소스 | KRX Open API(별도 인증키) | yfinance(비공식, 최소 사용) |
| 계정 식별 | XBRL account_id(한국 표준) | XBRL us-gaap 태그(미국 표준) — **같은 원리** |
| 분기 파생 계산 | 필요(2분기=반기−1분기 등) | duration 필터링만 필요(SEC가 이미 분기/누적을 다 제공, 일수로 골라 쓰면 됨) |
| 데이터 커버리지 | 12분기/5개년 거의 항상 채움 | 회사가 XBRL 공시를 시작한 시점부터 전체 이력(보통 넉넉함) |
| 통화/단위 | 원화, 억원 단위 | 달러, 백만달러 단위(단 EPS는 예외) |

## 향후 확장 (아직 안 함)

- investment-thesis-writer 스타일 정성분석
- 버핏멍거_가치평가 스타일 DCF·오너어닝·해자 체크리스트
- 성장성·활동성 지표 추가(현재 11개만 구현)
