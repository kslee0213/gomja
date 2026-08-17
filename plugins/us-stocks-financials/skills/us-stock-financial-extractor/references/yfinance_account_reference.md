# yfinance 계정 매핑 참고

`build_workbook.py`의 `ACCOUNTS` 딕셔너리가 실제로 쓰는 yfinance 키 이름과 근거.

⚠️ 이 환경(claude.ai 채팅)에서는 yfinance 설치·pypi.org 접근이 막혀 있어 라이브
데이터로 직접 검증하지 못했다. 아래 키 이름은 웹 검색(2026-08-16 기준)으로
확인한 것이며, yfinance 실사용 중 계정이 계속 비어 있으면 이 문서와
`ACCOUNTS`를 함께 갱신한다.

## 손익계산서 (income_statement / `ticker.financials`, `ticker.quarterly_financials`)

| 표준 키 | yfinance 후보 (우선순위 순) |
|---|---|
| 매출액 | `Total Revenue`, `Operating Revenue` |
| 매출원가 | `Cost Of Revenue`, `Reconciled Cost Of Revenue` |
| 매출총이익 | `Gross Profit` |
| 영업이익 | `Operating Income` |
| EBITDA | `EBITDA`, `Normalized EBITDA` |
| 세전이익 | `Pretax Income` |
| 법인세비용 | `Tax Provision` |
| 당기순이익 | `Net Income`, `Net Income Common Stockholders` |
| 이자비용 | `Interest Expense` |
| EPS(희석) | `Diluted EPS` |

## 재무상태표 (balance_sheet / `ticker.balance_sheet`, `ticker.quarterly_balance_sheet`)

| 표준 키 | yfinance 후보 |
|---|---|
| 유동자산 | `Current Assets` |
| 자산총계 | `Total Assets` |
| 유동부채 | `Current Liabilities` |
| 부채총계 | `Total Liabilities Net Minority Interest` ⚠ (`Total Liabilities` 아님) |
| 자본총계 | `Stockholders Equity` ⚠ (`Total Equity` 아님) |
| 현금및현금성자산 | `Cash And Cash Equivalents`, `Cash Cash Equivalents And Short Term Investments` |
| 재고자산 | `Inventory` |
| 매출채권 | `Accounts Receivable`, `Receivables` |
| 총차입금 | `Total Debt` |
| 유형자산 | `Net PPE`, `Gross PPE` |

## 현금흐름표 (cash_flow / `ticker.cashflow`, `ticker.quarterly_cashflow`)

| 표준 키 | yfinance 후보 |
|---|---|
| 영업활동현금흐름 | `Operating Cash Flow`, `Cash Flow From Continuing Operating Activities` |
| 잉여현금흐름 | `Free Cash Flow` |
| 설비투자 | `Capital Expenditure` (보통 음수로 표기됨) |
| 감가상각비 | `Depreciation And Amortization`, `Depreciation Amortization Depletion` |
| 배당금지급 | `Cash Dividends Paid`, `Common Stock Dividend Paid` (보통 음수) |

## v1.0.0에서 실제로 틀렸던 키 (참고용 — 재발 방지)

| v1.0.0이 쓴 잘못된 키 | 실제 yfinance 키 |
|---|---|
| `Total Equity` | `Stockholders Equity` |
| `Total Liabilities` | `Total Liabilities Net Minority Interest` |
| `Revenue` | `Total Revenue` |

## `.info` 딕셔너리 (fetch_extra_info.py가 쓰는 것 — 현재 시점 스냅샷)

`currentPrice`, `marketCap`, `trailingPE`, `forwardPE`, `priceToBook`, `priceToSalesTrailing12Months`, `trailingEps`, `bookValue`, `dividendYield`, `dividendRate`, `payoutRatio`, `returnOnEquity`, `returnOnAssets`, `debtToEquity`, `beta`, `fiftyTwoWeekHigh`, `fiftyTwoWeekLow`

이 값들은 Yahoo Finance가 이미 계산해서 제공하는 것이라, DART 버전처럼 원자료로부터 직접 수식으로 재계산하지 않는다(한계로 인지하고 있음 — SKILL.md "알려진 한계" 참고).
