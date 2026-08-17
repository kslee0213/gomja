# SEC EDGAR XBRL 태그 참고

`fetch_financials.py`/`build_workbook.py`의 `ACCOUNTS` 딕셔너리가 쓰는 us-gaap 표준 태그와 근거.

이 환경(claude.ai 채팅)에서는 `data.sec.gov` 접근이 막혀 있어(네트워크 허용목록) 라이브 데이터로 직접 검증하지 못했다. 아래 태그는 웹 검색(2026-08-17 기준)과 XBRL 표준 관행으로 확인한 것이며, 실사용 중 특정 회사에서 계정이 비면 이 문서와 `ACCOUNTS`를 함께 갱신한다.

## API 구조

```
https://data.sec.gov/api/xbrl/companyfacts/CIK{10자리}.json
```

- 응답은 그 회사가 XBRL로 공시한 **모든** 계정의 **모든** 과거 값(10-K/10-Q 전체 이력)을 담고 있다.
- 구조: `facts.us-gaap.{태그}.units.USD[]` — 각 원소는 `{start, end, val, fy, fp, form, filed}`.
- **instant(재무상태표)**: `start` 없이 `end`만 있음(그 시점 잔액).
- **duration(손익계산서/현금흐름표)**: `start`와 `end` 둘 다 있음. **같은 계정에 3개월(분기단독)/6개월(반기누적)/9개월(3분기누적)/12개월(연간) 값이 전부 섞여서 들어있다** — `(end-start)` 일수로 걸러야 한다(`_duration_days()`).
  - 80~100일 → 분기 단독
  - 350~380일 → 연간(FY)

## 필수 헤더

```
User-Agent: {앱이름} {연락처이메일}
```
없으면 403이 날 수 있다(SEC 정책, 인증키는 아님).

## 계정 매핑

### 손익계산서 (duration, 3개월/12개월 필터링)

| 표준 키 | us-gaap 태그 후보 |
|---|---|
| 매출액 | `Revenues`, `RevenueFromContractWithCustomerExcludingAssessedTax`, `SalesRevenueNet` |
| 매출원가 | `CostOfRevenue`, `CostOfGoodsAndServicesSold` |
| 매출총이익 | `GrossProfit` |
| 영업이익 | `OperatingIncomeLoss` |
| 세전이익 | `IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest` 등 |
| 법인세비용 | `IncomeTaxExpenseBenefit` |
| 당기순이익 | `NetIncomeLoss`, `ProfitLoss` |
| 이자비용 | `InterestExpense` |
| EPS(희석) | `EarningsPerShareDiluted` — ⚠ 주당 값이라 백만달러 환산 대상 아님(`PER_SHARE_KEYS`) |

### 재무상태표 (instant)

| 표준 키 | us-gaap 태그 후보 |
|---|---|
| 유동자산 | `AssetsCurrent` |
| 자산총계 | `Assets` |
| 유동부채 | `LiabilitiesCurrent` |
| 부채총계 | `Liabilities` |
| 자본총계 | `StockholdersEquity`, `StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest` |
| 현금및현금성자산 | `CashAndCashEquivalentsAtCarryingValue` 등 |
| 재고자산 | `InventoryNet` |
| 매출채권 | `AccountsReceivableNetCurrent`, `ReceivablesNetCurrent` |
| 유형자산 | `PropertyPlantAndEquipmentNet` |

### 현금흐름표 (duration)

| 표준 키 | us-gaap 태그 후보 |
|---|---|
| 영업활동현금흐름 | `NetCashProvidedByUsedInOperatingActivities` 등 |
| 설비투자 | `PaymentsToAcquirePropertyPlantAndEquipment` |
| 감가상각비 | `DepreciationDepletionAndAmortization`, `DepreciationAmortizationAndAccretionNet`, `Depreciation` |
| 배당금지급 | `PaymentsOfDividends`, `PaymentsOfDividendsCommonStock` |

> DART 때와 달리, 미국 기업은 감가상각비를 현금흐름표 조정 항목으로 명시적으로 XBRL 태그를 붙이는 경우가 많아 API로 가져올 수 있는 확률이 더 높다(삼성전자처럼 주석에만 있어서 API로 못 가져오는 경우와 대비됨). 다만 회사마다 다를 수 있어 여전히 후보를 여러 개 둔다.

## instant(BS) 항목의 annual/quarterly 분류 — 실제로 잡은 버그

**최초 구현에서 BS 항목을 quarterly/annual 구분 없이 그대로 양쪽에 다 넣었다가, 아직 10-K(연간)가 없는 진행 중인 분기말 스냅샷이 "연간" 쪽에 섞여 들어가는 버그가 실제로 재현됐다.** (`_collect_duration_dates()`로 IS/CF의 실제 연간/분기 날짜 집합을 먼저 구하고, BS는 그 날짜와 매칭되는 것만 각 세트에 넣도록 수정.) 새로운 계정을 추가할 때도 이 원칙을 지킨다.

## 주가는 SEC에 없다

SEC EDGAR는 공시 데이터만 제공하고 시장 주가는 없다. `fetch_extra_info.py`가 yfinance로 현재가·시가총액·상장주식수만 최소한으로 가져오고, PER/PBR/PSR은 이 주가와 SEC 재무제표 값(EPS/자본총계/매출액)을 조합해 `build_workbook.py`가 직접 수식으로 계산한다(yfinance의 완제품 `trailingPE` 값을 그대로 믿지 않음 — 감사 가능성 원칙).
