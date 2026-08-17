---
name: us-stock-financial-extractor
description: >
  Use this skill when the user asks to fetch, extract, or analyze US-listed
  (NASDAQ/NYSE) company financial statements, such as "애플 재무제표 뽑아줘",
  "테슬라 년 분석해줘", "AAPL 분기 분석해줘", or requests to compare multiple US
  companies like "마이크로소프트, 구글 비교해줘". Financial statements come from
  SEC EDGAR's official free XBRL API (no key required, no scraping risk);
  current stock price comes from yfinance (minimal use only). This is the US
  counterpart to dart-kospi-financials, built on the same design principle
  (formula-based audit trail via a raw-data sheet, embedded charts).
metadata:
  version: "3.0.0"
---

# US Stock Financial Extractor

기업명(또는 Ticker)을 입력받아 재무제표를 수집하고, 분기·연간 재무제표와 재무지표, 간이 투자분석을 담은 엑셀 파일을 만든다.

> ⚠️ **v2.0.0(yfinance 전면 사용) → v3.0.0(SEC EDGAR 전환)**. v2.0.0은 재무제표까지 yfinance로 가져왔는데, yfinance는 비공식 스크레이핑이라 Yahoo가 클라우드 IP를 403으로 차단하는 문제가 실제로 발생했다(다른 세션에서 재현됨). v3.0.0은 재무제표를 **SEC(미국 증권거래위원회)가 직접 제공하는 공식 무료 API(data.sec.gov)**로 전환했다 — DART Open API와 원리가 같다(XBRL 표준계정 태그 기반). yfinance는 SEC가 제공하지 않는 "현재 주가" 하나만 위해 최소한으로만 남겼다.

## 데이터 소스 구조 (중요)

| 데이터 | 소스 | 특징 |
|---|---|---|
| 재무제표(BS/IS/CF) | **SEC EDGAR** (`data.sec.gov`) | 공식 정부 API, 완전 무료, **API 키 불필요**(단 User-Agent에 연락처 이메일 필요 — 인증키 아님), 차단 위험 낮음(비공식 스크레이핑이 아니라 진짜 API) |
| 현재 주가·시가총액 | yfinance (Yahoo Finance) | 여전히 비공식이라 403/429 위험이 남아있음. 이 부분만 실패해도 재무제표는 영향받지 않는다(투자분석의 PER/PBR 등 주가 연동 지표만 빈다) |

## 0. 사전 준비

```bash
pip install requests openpyxl --break-system-packages   # SEC EDGAR용, 대부분 이미 설치돼 있음
pip install yfinance --break-system-packages              # 주가 전용, 선택이지만 강력 권장
```

### 0-0. 작업 시작 전 필수 확인

1. **연락처 이메일**: SEC EDGAR는 API 키는 없지만 User-Agent 헤더에 유효한 이메일이 반드시 있어야 한다(SEC 정책 — 없으면 403 위험). 대화에 없으면 먼저 사용자에게 물어본다: "SEC API 호출 시 User-Agent에 넣을 연락처 이메일을 알려주세요(예: you@example.com)."
2. **네트워크 허용 목록**: `data.sec.gov`가 Cowork 네트워크 허용 목록에 없으면 접근이 막힌다(이 채팅 환경에서 직접 확인됨 — `data.sec.gov`가 차단되어 있었다). 접근이 안 되면 DART 때처럼 이 채팅 환경과 사용자의 Cowork 실행 환경이 다를 수 있다는 점을 안내하고, Cowork 설정에서 이 도메인을 추가해달라고 요청한다.
3. yfinance는 선택이다 — 없어도 재무제표(핵심 산출물)는 정상 생성되고, 주가 연동 지표만 빈다.

### ⚠️ yfinance(주가)가 403/429로 막힐 때의 대응

`fetch_extra_info.py` 실행 중 403/429가 나면:
1. **절대로 조용히 다른 방식(웹 검색 기반 수동 리포트 등)으로 전환하지 않는다.** 재무제표(SEC EDGAR)는 정상 작동하니, 이 부분은 그대로 진행하고 "주가 연동 지표만 못 채웠다"고 정확히 알린다.
2. 재시도는 1~2회만 짧게 시도한다.
3. `data.sec.gov`가 아니라 `query1/2.finance.yahoo.com`이 막힌 것이므로, SEC 기반 재무제표 작업 자체를 멈추지 않는다.

## 1. 기업명 → Ticker → CIK 변환

```
python scripts/ticker_lookup.py <기업명 또는 Ticker> --contact <연락처이메일>
```

- `COMMON_TICKERS`에 38개 고유 기업(한글/영문 라벨 합쳐 83개)이 매핑되어 있다. 없으면 입력을 Ticker로 간주한다.
- SEC의 공식 전체 티커→CIK 매핑(`sec.gov/files/company_tickers.json`, 역시 키 불필요)을 한 번 받아 캐시해두고 재사용한다(회사마다 개별 조회 불필요).
- CIK를 못 찾으면 명확히 실패를 알리고 정확한 Ticker를 요청한다.

## 2. 재무제표 수집 (SEC EDGAR)

```
python scripts/fetch_financials.py <ticker> <10자리CIK> --contact <연락처이메일>
python scripts/fetch_extra_info.py <ticker>   # 주가(yfinance, 선택)
```

- `fetch_financials.py`는 `https://data.sec.gov/api/xbrl/companyfacts/CIK{10자리}.json`을 **통째로** 받아 `cache/secfacts_{ticker}.json`에 원본 그대로 저장한다(가공은 build_workbook.py가 읽을 때 한다 — 나중에 새 계정을 추가하고 싶어도 재조회 불필요).
- **레이트리밋**: 초당 10회. 기업 하나당 요청 1회로 끝나므로 순차 처리하면 문제없다. 여러 기업을 연속 조회할 때 `time.sleep(0.15)` 정도의 간격을 둔다(스크립트에 이미 포함됨).
- 캐시가 있으면 재사용한다(`--force`로 강제 재조회).

### duration 계정(손익계산서/현금흐름표) 파싱 주의

같은 계정(예: 매출액)에 3개월(분기단독)/6개월/9개월/12개월(연간) 값이 전부 섞여서 응답에 들어있다. `(end-start)` 일수로 80~100일=분기, 350~380일=연간을 걸러낸다(`fetch_financials.py`의 `_duration_days()`). **재무상태표(instant) 항목은 이 duration 계정들의 실제 연간/분기 날짜와 매칭되는 것만 각 세트(quarterly/annual)에 넣는다** — 그렇지 않으면 아직 10-K가 없는 진행 중인 분기말 BS 스냅샷이 "연간" 쪽에 잘못 섞여 들어가는 버그가 생긴다(실제로 재현·수정됨, `references/sec_edgar_xbrl_reference.md` 참고).

## 3. 엑셀 생성

```
python scripts/build_workbook.py <ticker> "<기업명>" --period both --quarters 12 --years 5 --outdir /mnt/user-data/outputs
```

### 생성되는 시트

| 시트 | 내용 |
|---|---|
| `분기_재무제표` | 있는 만큼의 분기(SEC 공시 이력에 따라 다름, 보통 여러 개년치 다 있을 수 있음 — yfinance의 4~5개 제약과 달리 SEC는 회사가 XBRL 공시를 시작한 이후 전체 이력을 제공) |
| `연간_재무제표` | 있는 만큼의 연도 |
| `지표_분기` / `지표_연간` | 11개 재무비율 + 라인차트 4개 |
| `투자분석` | A.재무비율 요약 B.위험신호 C.주가 연동 지표(PER/PBR/PSR을 **직접 수식으로 계산** — yfinance의 완제품 값을 그대로 안 믿음) D.간이 투자판단 |
| `원본데이터` | SEC/yfinance에서 뽑은 raw 값(숨김). 다른 모든 시트가 이 시트를 수식으로 참조 |

- PER = 현재가 ÷ EPS(SEC), PBR = 현재가 ÷ (자본총계÷상장주식수), PSR = 현재가 ÷ (매출액÷상장주식수) — 전부 감사 가능한 수식.

## 4. 여러 기업 비교

```
python scripts/build_comparison_workbook.py AAPL:Apple MSFT:Microsoft --quarters 12 --years 5 --outdir /mnt/user-data/outputs
```

비교 대상 기업은 먼저 각각 `fetch_financials.py`로 캐시를 만들어둬야 한다. `12분기 비교`/`최근5년 비교` 시트는 각 기업의 실제 캐시된 시계열을 읽어 서로 다른 값을 보여준다(v1.0.0의 동일값 버그는 v2.0.0에서 이미 수정, v3.0.0은 데이터 소스만 SEC로 교체).

## 5. 저장 및 전달

완료 후 요약: 실제로 채워진 분기/연도 개수, 못 찾은 계정 목록(SEC 태그 후보가 다 실패한 경우), 주가 지표 중 못 가져온 항목(yfinance 실패 시)을 안내한다.

## 참고

### v3.0.0 (SEC EDGAR 전환)

1. **데이터 소스 교체**: yfinance(재무제표) → SEC EDGAR CompanyFacts API. 계정 매칭을 문자열 인덱스 이름 추측(yfinance)에서 XBRL 표준 태그(SEC, DART와 같은 원리)로 전환해 더 견고해졌다.
2. **실제로 잡은 버그**: BS(재무상태표, instant) 항목을 quarterly/annual 구분 없이 그대로 양쪽에 넣었다가, 아직 10-K가 없는 진행 중인 분기말 스냅샷이 "연간" 쪽에 섞여 들어가는 문제가 실제 테스트로 재현됐다. IS/CF(duration)의 실제 날짜 집합과 매칭시켜 수정.
3. **실제로 잡은 버그 2**: EPS(주당 지표)에 다른 계정과 동일하게 백만달러 단위 환산(÷1,000,000)을 적용해서 PER이 45,000,000배로 나오는 버그가 있었다. `PER_SHARE_KEYS`로 예외 처리.
4. **PER/PBR/PSR을 직접 수식으로 계산**: yfinance의 이미 계산된 `trailingPE` 등을 그대로 믿지 않고, SEC 재무제표 값(EPS/자본총계/매출액)과 주가를 조합해 우리가 직접 수식을 만든다(감사 가능성 원칙 강화).
5. yfinance 의존 범위를 "재무제표 전체"에서 "현재 주가 하나"로 최소화 — 이제 Yahoo가 또 차단해도 핵심 산출물(재무제표)은 영향받지 않는다.
6. 이 환경(claude.ai 채팅)에서 `data.sec.gov` 자체도 네트워크 허용목록에 없어 라이브 검증은 못 했다. 웹 검색으로 확인한 XBRL 태그 구조로 구현했고, SEC CompanyFacts 응답 형식과 동일한 합성 데이터로 전체 파이프라인(분기/연간 분리, PER/PBR/PSR 계산, 비교 워크북)을 실행·재계산까지 검증했다.

### 알려진 한계

- CIK가 없는 비상장/신규 상장 기업, 또는 SEC 매핑에 아직 없는 회사는 지원 안 됨.
- 투자분석의 "간이 투자판단"은 재무비율 3개만 보는 단순 규칙이다(DART의 N섹션만큼 정교하지 않음).
- investment-thesis-writer·버핏멍거_가치평가 스타일의 정성분석/DCF 확장은 아직 없다.
