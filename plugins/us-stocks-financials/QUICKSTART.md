# Quick Start Guide

## 설치

```bash
pip install requests openpyxl yfinance --break-system-packages
```

## 사용 방법 (순서대로)

### 1️⃣ 기업명 → Ticker → CIK 확인

```bash
python scripts/ticker_lookup.py "Apple" --contact you@example.com
```

### 2️⃣ 재무 데이터 수집

```bash
# 재무제표(SEC EDGAR, 필수) — CIK는 1단계 결과 사용
python scripts/fetch_financials.py AAPL 0000320193 --contact you@example.com

# 주가(yfinance, 선택이지만 권장)
python scripts/fetch_extra_info.py AAPL
```

### 3️⃣ 엑셀 생성

```bash
python scripts/build_workbook.py AAPL "Apple Inc." --period both --outdir /mnt/user-data/outputs
```

### 4️⃣ 여러 기업 비교

```bash
python scripts/fetch_financials.py MSFT 0000789019 --contact you@example.com
python scripts/build_comparison_workbook.py AAPL:Apple MSFT:Microsoft --outdir /mnt/user-data/outputs
```

## 출력 예시 (--period both 기준)

```
Apple Inc._20260817.xlsx
├─ 분기_재무제표   (SEC 공시 이력에 있는 만큼)
├─ 연간_재무제표   (SEC 공시 이력에 있는 만큼)
├─ 지표_분기       (11개 지표 + 차트 4개)
├─ 지표_연간       (11개 지표 + 차트 4개)
├─ 투자분석        (재무비율/위험신호/주가지표/간이투자판단)
└─ 원본데이터      (숨김, 수식 참조 소스)
```

## 트러블슈팅

| 문제 | 원인/해결 |
|---|---|
| SEC 403 에러 | User-Agent에 유효한 이메일이 들어갔는지 확인. 그래도 안 되면 `data.sec.gov`가 Cowork 네트워크 허용목록에 있는지 확인 |
| yfinance 403/429 | 재무제표(SEC)는 정상 진행됨 — 주가 연동 지표만 빈다. 절대 다른 방식(수동 리서치 등)으로 조용히 전환하지 않는다 |
| 계정이 비어있음 | SEC 태그 후보에 없는 표기를 쓰는 회사일 수 있음. `references/sec_edgar_xbrl_reference.md`에 후보 추가 |
| 티커/CIK를 못 찾음 | `ticker_lookup.py`로 정확한 CIK 확인 후 재시도 |

더 자세한 정보는 `SKILL.md` 참고.
