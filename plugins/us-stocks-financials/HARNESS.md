# 파이프라인 아키텍처

## 단일 기업 흐름

```
사용자 요청
    ↓
ticker_lookup.py       (기업명 → Ticker, 실존 여부 검증)
    ↓
fetch_financials.py    (BS/IS/CF 수집 → cache/financials_{ticker}_{quarterly|annual}.json)
fetch_extra_info.py    (주가/배당/섹터 수집 → cache/extra_info_{ticker}.json)
    ↓
build_workbook.py      (캐시를 실제로 읽어 원본데이터 시트 생성 →
                         나머지 시트는 전부 원본데이터를 참조하는 수식)
    ↓
엑셀 파일(6개 시트) 생성 완료
```

## 다중 기업 비교 흐름

```
여러 기업 요청
    ↓
[기업마다 순차 처리 — yfinance 레이트리밋 고려해 너무 빠르게 연속 호출하지 않는다]
- AAPL: fetch_financials + fetch_extra_info
- MSFT: fetch_financials + fetch_extra_info
- GOOGL: fetch_financials + fetch_extra_info
    ↓
build_comparison_workbook.py
  (각 기업의 캐시된 재무제표를 실제로 읽어 시계열별로 나란히 배치.
   "12분기 비교"와 "최근5년 비교"는 서로 다른 실제 데이터를 보여준다 — v1.0.0의
   두 시트가 동일했던 버그를 v2.0.0에서 수정했다.)
    ↓
비교 엑셀 생성 완료
```

## 캐싱 구조

```
cache/
├── company_{ticker}.json               (ticker_lookup.py)
├── financials_{ticker}_quarterly.json  (fetch_financials.py — frequency가 파일명에 반드시 포함)
├── financials_{ticker}_annual.json     (fetch_financials.py)
└── extra_info_{ticker}.json            (fetch_extra_info.py)
```

- **build_workbook.py는 이 캐시를 실제로 읽는다.** v1.0.0은 캐시 파일들을 만들기만 하고 build_workbook.py가 이를 전혀 참조하지 않아, 캐싱 시스템 자체가 워크북 생성과 단절되어 있었다(문서에는 "재실행 70~80% 단축"이라고 되어 있었지만 사실이 아니었다).
- 캐시가 없으면 build_workbook.py는 조용히 넘어가지 않고 명확한 경고를 stderr에 남긴다.

## DART 버전과의 설계 차이

| 항목 | dart-kospi-financials | us-stocks-financials |
|---|---|---|
| 데이터 소스 | DART Open API(공식, API 키 필요) | yfinance(비공식 스크레이핑, 키 불필요) |
| 계정 식별 | XBRL account_id(표준화됨, 회사별 계정명이 달라도 안정적) | yfinance 인덱스 라벨(문자열 그대로, 라이브러리가 어느 정도 정규화해줌) |
| 분기 파생 계산 | 필요(2분기=반기−1분기 등, 누적 보고 관행 때문) | 불필요(10-Q가 이미 분기 단독 값을 보고) |
| 데이터 커버리지 | 12분기/5개년을 거의 항상 채울 수 있음 | 보통 4~5분기/4개년까지만 (yfinance 무료 데이터 한계) |
| 통화/단위 | 원화, 억원 단위 | 달러, 백만달러 단위 |
| 주가 연동 | KRX Open API(별도 인증키) | yfinance `.info`(같은 라이브러리, 추가 키 불필요, 다만 현재 시점 스냅샷만) |

## 향후 확장 (아직 안 함)

- investment-thesis-writer 스타일 정성분석(사업내용 원문 + 웹리서치 + 정성 텍스트)
- 버핏멍거_가치평가 스타일 DCF·오너어닝·해자 체크리스트
- 성장성·활동성 지표 추가(현재 재무비율 11개만 구현, DART 버전은 34개 METRIC_RULES)
- ~~thesis-auditor 통합~~ (해당 스킬은 다른 대화에서 이미 삭제됨 — 이 계획 항목 제거)
