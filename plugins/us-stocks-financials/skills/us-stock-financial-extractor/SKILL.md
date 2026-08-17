---
name: us-stock-financial-extractor
description: >
  Use this skill when the user asks to fetch, extract, or analyze US-listed
  (NASDAQ/NYSE) company financial statements via yfinance, such as
  "애플 재무제표 뽑아줘", "테슬라 년 분석해줘", "AAPL 분기 분석해줘", or requests to
  compare multiple US companies like "마이크로소프트, 구글 비교해줘". No API key
  is required (yfinance/Yahoo Finance). This is the US counterpart to
  dart-kospi-financials, built on the same design (formula-based audit trail,
  raw-data sheet, embedded charts) but adapted for yfinance's data model and
  limitations.
metadata:
  version: "2.0.0"
---

# US Stock Financial Extractor

기업명(또는 Ticker)을 입력받아 yfinance로 재무제표를 수집하고, 분기·연간 재무제표와 재무지표, 간이 투자분석을 담은 엑셀 파일을 만든다.

> ⚠️ **v1.0.0에서 v2.0.0으로 전면 재작성됨.** v1.0.0은 문서가 광고한 기능(200+ 기업 매핑, 19개 지표, 지표 시트+차트, 투자분석 26개 섹션)의 상당수가 실제 코드에 없었고, 캐싱이 워크북 생성과 연결되지 않아 재실행 속도 개선도 실제로는 작동하지 않았으며, 비교 워크북의 "12분기"와 "최근5년" 시트가 완전히 동일한 값을 보여주는 버그가 있었다. v2.0.0은 dart-kospi-financials의 검증된 설계(원본데이터 시트 + 수식 참조로 감사 가능하게, 실제 라이브러리 반환값 확인 후 계정 매칭)를 기준으로 처음부터 다시 짰다. 자세한 변경 이력은 이 파일 하단 참고.

## yfinance의 근본적인 제약 — 반드시 사용자에게 정직하게 안내한다

DART(한국 전자공시)와 달리 yfinance는 **공식 API가 아니라 Yahoo Finance 웹 데이터를 스크레이핑하는 라이브러리**다. 이로 인한 제약을 작업 시작 전에 사용자에게 알려준다:

- **분기 데이터는 보통 4~5개 정도만 제공된다.** "12개 분기"를 요청해도 실제로 4~5개만 채워지는 게 정상이다(DART처럼 12개를 다 채울 수 있다고 가정하지 않는다).
- **연간 데이터는 보통 최근 4개년 정도만 제공된다.**
- **비공식 스크레이핑 특성상 가끔 필드가 비거나 형식이 바뀔 수 있다.** 계정을 못 찾으면 빈 칸으로 두고 지어내지 않는다(`ACCOUNTS`의 후보 목록에도 없으면 원본데이터 시트에 그 계정 자체가 안 나타난다).
- **레이트리밋**: 너무 빠르게 여러 번 호출하면 일시적으로 차단될 수 있다. 여러 기업을 연속 조회할 때는 사이에 약간의 지연을 둔다.
- API 키가 필요 없다는 것이 이 스킬의 장점이지만, 그만큼 데이터 안정성은 DART보다 낮다는 걸 감안하고 사용자에게 전달한다.

## 0. 사전 준비

```bash
pip install yfinance --break-system-packages   # (Cowork 환경에서 pypi.org 접근 필요)
```

- API 키 불필요.
- yfinance가 설치·접근 가능한지 먼저 확인한다(`python3 -c "import yfinance"`). 설치가 안 되거나 네트워크가 막혀 있으면, DART 스킬 때처럼 이 채팅 환경과 사용자의 Cowork 실행 환경이 다를 수 있다는 점을 사용자에게 안내한다.

## 1. 기업명 → Ticker 변환

```
python scripts/ticker_lookup.py <기업명 또는 Ticker>
```

- `ticker_lookup.py`의 `COMMON_TICKERS`에 등록된 건 **38개 고유 기업**(한글/영문 라벨 합쳐 83개)뿐이다. "200+ 기업 자동 매핑" 같은 과장된 표현을 쓰지 않는다.
- 매핑에 없으면 사용자가 입력한 문자열을 그대로 Ticker로 간주해 yfinance로 실존 여부를 검증한다. 존재하지 않으면 명확히 실패를 알리고 정확한 Ticker를 요청한다(존재하지도 않는 티커로 계속 진행하지 않는다).

## 2. 재무제표 수집 (캐시)

```
python scripts/fetch_financials.py <ticker> --period both --quarters 12 --years 5
python scripts/fetch_extra_info.py <ticker>
```

- `fetch_financials.py`는 `cache/financials_{ticker}_{quarterly|annual}.json`에 저장한다. **frequency가 파일명에 포함되어 있어야 한다** — v1.0.0은 이게 없어서 quarterly를 annual이 덮어쓰는 버그가 있었다.
- pandas Timestamp·numpy 타입은 `_json_safe()`를 거쳐 순수 파이썬 타입으로 변환 후 저장한다(표준 `json.dump()`가 이 타입들을 직렬화 못 해 실패하는 문제를 방지).
- 캐시가 이미 있으면 재사용한다(`--force`로 강제 재조회 가능).
- `fetch_extra_info.py`는 현재 주가·PER·PBR·배당 등 yfinance `.info`의 스냅샷 값을 `cache/extra_info_{ticker}.json`에 저장한다. 이 값들은 "현재 시점" 값이라 연도별로 달라지지 않는다는 점을 유의한다(과거 분기/연도별 PER 등은 계산하지 않는다 — 과거 시점 주가 데이터가 있어야 하는데 무료 yfinance로는 신뢰성 있게 재현하기 어렵다).

## 3. 엑셀 생성

```
python scripts/build_workbook.py <ticker> "<기업명>" --period both --quarters 12 --years 5 --outdir /mnt/user-data/outputs
```

`build_workbook.py`는 **1~2단계에서 만든 캐시를 실제로 읽는다**(v1.0.0은 캐시를 무시하고 매번 yfinance를 새로 호출했다). 캐시가 없으면 경고를 출력하고 그 부분 시트를 건너뛴다 — 반드시 fetch 스크립트를 먼저 실행한다.

### 생성되는 시트

| 시트 | 내용 |
|---|---|
| `분기_재무제표` | 있는 만큼의 분기(보통 4~5개). 원본데이터 시트를 참조하는 수식, 백만달러 단위 |
| `연간_재무제표` | 있는 만큼의 연도(보통 4개). 위와 동일 |
| `지표_분기` | 11개 재무비율(매출총이익률·영업이익률·순이익률·ROE·ROA·부채비율·유동비율·자기자본비율 등) + 라인차트 4개 |
| `지표_연간` | 위와 동일(연간 기준) |
| `투자분석` | A.재무비율 요약 B.위험신호(부채비율/유동비율 임계값) C.주가 연동 지표(현재가·PER·PBR·PSR·배당수익률, extra_info 캐시 기준) D.간이 투자판단(재무비율 3개만 보는 결정론적 등급 — 투자 조언 아님을 명시) |
| `원본데이터` | 캐시에서 읽은 raw 값(달러, 무환산). 숨김. 다른 모든 시트가 이 시트를 수식으로 참조한다 |

- **DART와 다르게 분기별 파생 계산(2분기=반기−1분기 같은)이 필요 없다.** 미국 10-Q(분기보고서)는 이미 그 분기 자체의 손익을 독립적으로 보고하므로, yfinance의 quarterly 데이터를 그대로 쓰면 된다.
- 계정 매칭은 `ACCOUNTS` 딕셔너리에 등록된 yfinance의 **실제 확인된 키 이름**(`Stockholders Equity`, `Total Liabilities Net Minority Interest`, `Total Revenue` 등)을 쓴다. v1.0.0은 `'Total Equity'`, `'Total Liabilities'`, `'Revenue'`처럼 존재하지 않는 키를 써서 지표가 전부 비어 있었다.

## 4. 여러 기업 비교

```
python scripts/build_comparison_workbook.py AAPL:Apple MSFT:Microsoft GOOGL:Google --quarters 12 --years 5 --outdir /mnt/user-data/outputs
```

- 비교 대상 기업들은 먼저 각각 `fetch_financials.py`로 캐시를 만들어둬야 한다.
- 시트: `12분기 비교`, `최근5년 비교`. **v1.0.0은 이 두 시트가 완전히 동일한 값(현재 시점 스냅샷)을 보여주는 버그가 있었다.** v2.0.0은 각 기업의 실제 분기/연간 캐시를 읽어 서로 다른 진짜 시계열 값을 보여주고, 지표별 라인 차트(기업=계열)도 실제로 그린다.
- 캐시가 없는 기업은 빈 칸으로 두고 시트 상단에 "⚠ 캐시 없음" 경고를 남긴다(조용히 건너뛰지 않는다).

## 5. 저장 및 전달

- 파일명: `{기업명}[_연간|_분기]_{YYYYMMDD}.xlsx` (both면 접미사 없음)
- 완료 후 요약: 실제로 채워진 분기/연도 개수(yfinance 데이터 제약으로 요청보다 적을 수 있음을 재차 안내), 못 찾은 계정 목록, 주가 지표 중 못 가져온 항목이 있으면 함께 안내한다.

## 참고

### v2.0.0 (전면 재작성)

v1.0.0에서 실제로 발견·수정한 문제 전체 목록:

1. **`build_workbook.py`의 `date.strftime('%Y-Q%q')`** — `%q`는 유효한 strftime 코드가 아니라 분기 헤더가 깨지는 버그. 제거하고 `YYYY-MM-DD[:7]` 형식으로 교체.
2. **`fetch_financials.py` 캐시 파일명 충돌** — quarterly/annual이 같은 파일명(`financials_{ticker}_{날짜}.json`)을 써서 `--period both` 실행 시 서로 덮어씀. frequency를 파일명에 포함해 분리.
3. **JSON 직렬화 실패 위험** — `DataFrame.to_dict()`에 pandas Timestamp·numpy 타입이 섞여 표준 `json.dump()`가 실패할 수 있었음. `_json_safe()` 재귀 변환 함수 추가.
4. **비교 워크북의 "12분기"="최근5년" 버그** — 두 시트가 시계열이 아니라 현재 시점 스냅샷만 반복 표시. 실제 캐시된 재무제표를 읽어 진짜 시계열로 재작성.
5. **캐싱과 워크북 생성이 연결 안 됨** — `build_workbook.py`가 캐시를 무시하고 매번 API를 새로 호출해, "재실행 70~80% 단축" 주장이 사실이 아니었음. 캐시를 실제로 읽도록 수정.
6. **`calculate_metrics.py`가 어디서도 호출되지 않는 죽은 코드** — 로직을 `build_workbook.py`의 `INDICATOR_ROWS`로 흡수하고 파일 자체를 제거.
7. **존재하지 않는 yfinance 키 사용** — `'Total Equity'`→`'Stockholders Equity'`, `'Total Liabilities'`→`'Total Liabilities Net Minority Interest'`, `'Revenue'`→`'Total Revenue'`로 실제 키 확인 후 수정. 웹 검색으로 yfinance의 실제 반환 필드명을 사전 확인했다(yfinance 설치·실행이 이 환경에서 불가능해 직접 라이브 검증은 못 했다 — 실사용 중 계정이 계속 비면 `ACCOUNTS`의 후보 목록을 추가 보완해야 할 수 있다).
8. **`except:`(bare except) 남발** — 실패 원인을 전부 숨기던 것을 구체적인 에러 메시지 출력으로 교체.
9. **문서-코드 불일치 전반** — "200+ 기업"(실제 38개), "19개 지표"(실제 11개), "투자분석 26개 섹션"(실제 4개 섹션) 등 과장된 문서를 실제 구현에 맞게 재작성. `thesis-auditor` 연동 계획도 제거(그 스킬은 별도 대화에서 이미 삭제됨).
10. **하드코딩된 값 → 수식 참조로 전환** — 모든 재무제표/지표 시트 셀이 "원본데이터" 시트를 참조하는 수식이 되도록 재설계(DART 스킬과 동일한 감사 가능성 원칙 적용).

### 알려진 한계 (v2.0.0에도 남아있는 것)

- yfinance 라이브 데이터로 직접 검증하지 못했다(이 환경에서 yfinance 설치·pypi 접근이 막혀 있었음). 실사용 중 계정 매칭이 안 되는 게 보이면 `build_workbook.py`의 `ACCOUNTS` 딕셔너리에 후보를 추가한다.
- 투자분석의 "간이 투자판단"은 재무비율 3개만 보는 매우 단순한 규칙이다. DART 스킬의 N섹션(재무건전성/수익성/성장성/저평가/주주환원 5개 카테고리, 5개년 충족연수 기준)만큼 정교하지 않다 — 필요하면 다음 단계로 확장한다.
- 과거 시점의 PER/PBR(연도별로 다른 값)은 계산하지 않는다. 현재 시점 스냅샷만 제공한다.
- investment-thesis-writer(정성분석)·버핏멍거_가치평가 스타일의 확장은 아직 없다. 필요해지면 DART 버전의 설계를 참고해 추가한다.
