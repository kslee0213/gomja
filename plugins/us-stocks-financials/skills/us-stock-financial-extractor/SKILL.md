---
name: us-stock-financial-extractor
description: >
  Use this skill when the user asks to fetch, extract, or analyze US-listed
  (NASDAQ/NYSE) company financial statements, such as "애플 재무제표 뽑아줘",
  "테슬라 년 분석해줘", "AAPL 분기 분석해줘", or requests to compare multiple US
  companies like "마이크로소프트, 구글 비교해줘". Financial statements come from
  SEC EDGAR's official free XBRL API (no key required); stock price/history comes
  from yfinance. This is the US counterpart to dart-kospi-financials with the
  SAME sheet structure and analysis content (재무제표, 지표+차트, 투자분석 A~N,
  투자판단 종합, 버핏멍거_가치평가, 가치평가_결론) and the same audit principle
  (formula-based audit trail via a raw-data sheet).
metadata:
  version: "4.0.0"
---

# US Stock Financial Extractor

기업명(또는 티커)을 입력받아 SEC EDGAR에서 재무제표를 수집하고, **dart-kospi-financials와 동일한 구성**의 분석 엑셀을 만든다. 버전 이력은 `CHANGELOG.md` 참고.

> **절대 규칙: "OO 년 분석해줘"는 확인 질문 없이 전체 파이프라인을 실행한다.** KR과 동일하게 "가치평가_결론+재무제표+투자분석+투자판단 종합+버핏멍거_가치평가를 전부 담은 파일 하나(7개 시트)"가 기본 산출물이다. 사용자가 "재무제표만"이라고 명시한 경우에만 4단계 이후를 건너뛴다.

## 0. 사전 준비

| 항목 | 내용 |
|---|---|
| 패키지 | `pip install requests openpyxl yfinance --break-system-packages` |
| 연락처 이메일 | SEC는 API 키 대신 User-Agent에 유효한 이메일을 요구한다(없으면 403 위험). 대화에 없으면 시작 전에 사용자에게 요청. |
| 네트워크 | `data.sec.gov`·`www.sec.gov`(재무제표, 필수), `query1/2.finance.yahoo.com`(주가, 권장)이 허용 목록에 있어야 한다. |
| yfinance 실패 시 | 재무제표(SEC)는 계속 진행하고 "주가 연동 지표만 못 채웠다"고 정확히 알린다. **조용히 웹 검색 수동 리포트로 전환하지 않는다.** 재시도는 1~2회만. |

## 1. 기업명 → Ticker → CIK

```
python scripts/ticker_lookup.py <기업명 또는 티커> --contact <이메일>
```
한글/영문 별칭 사전에 없으면 입력을 티커로 간주하고, SEC 공식 매핑(`company_tickers.json`, 캐시)에서 CIK를 찾는다. 못 찾으면 정확한 티커를 요청한다.

## 2. 데이터 수집

```
python scripts/fetch_financials.py <ticker> <10자리CIK> --contact <이메일>   # SEC 재무제표(필수)
python scripts/fetch_extra_info.py <ticker>                                  # 주가·월말 종가 7년·배당 이력(yfinance)
```

- CompanyFacts JSON을 통째로 `cache/secfacts_{ticker}.json`에 저장한다(재조회 불필요, `--force`로 갱신).
- `fetch_extra_info.py`는 현재가·시총·상장주식수에 더해 **월말 종가 7년 이력**과 **연도별 주당 배당 합계**를 가져온다 — 투자분석 L섹션(연도별 PER/PBR/PSR)과 N섹션 저평가 등급, 가치평가의 과거 PER 밴드에 쓰인다. 이력 조회가 실패해도 현재가만으로 계속 진행한다.

## 3. 엑셀 생성

```
python scripts/build_workbook.py <ticker> "<기업명>" --period {annual|quarterly|both} --outdir /mnt/user-data/outputs
```

`--period` 판단은 KR과 동일: "년"→annual(`_연간`), "분기"→quarterly(`_분기`), 없으면 both. 생성 시트(연간 기준)는 KR과 동일 구성이다:

| 시트 | 내용 |
|---|---|
| 분기_재무제표 / 연간_재무제표 | 손익·재무상태·현금흐름, 원본데이터 시트 참조 수식. 단위 백만달러(EPS는 달러). |
| 지표_분기 / 지표_연간 | 기본 지표 14개 + 비율 6개 + 임베드 차트 6개(KR과 동일 그룹) |
| 투자분석 (연간 전용) | A 회사 개황 / B 재무지표(건전성·수익성·성장성·활동성) / C 위험신호 6종 / D 청산가치(비율 셀참조) / E CCC / F FCF / G 현금흐름 3단 / H DuPont / I ROIC·NOPLAT(간이) / J 구성비 / K 외환손익 / L 주가 연동(연도별 종가·시총·PER·PBR·PSR + 현재가 블록) / M 배당 / N 투자판단 자동평가(A~E, KR과 동일 규칙) + 섹션별 차트 |
| 원본데이터 (숨김) | SEC 원본값(달러). 모든 시트의 유일한 소스. |

미국 공시 체계상 KR과 다른 점(지어내지 않고 명시): 과거 시가총액은 "현재 상장주식수 × 당시 종가" 근사치, 대주주·자기주식은 범위 밖(M섹션에 문구), 감가상각비는 SEC 태그로 잘 잡히는 편(오너어닝 계산 가능성이 KR보다 높음).

## 4. 투자판단 종합·가치평가·결론까지 이어서 (KR 5-3단계와 동일)

`--period annual`(또는 both)로 투자분석 시트를 만들었으면 멈추지 말고 이어서 실행한다:

1. 사업 내용 파악: DART 원문 대신 **웹 검색 + 최신 10-K의 Item 1(Business) 요약**을 쓴다(WebFetch로 SEC 사이트 열람 가능 시). 산업 동향·경쟁사·점유율·리스크·금리를 리서치하고 출처를 남긴다.
2. `content.json`(투자판단 종합) + `valuation_content.json`(버핏멍거_가치평가) 작성 — 스키마는 `investment-thesis-writer/SKILL.md` 참고. 정량 시트(N섹션 등급 등)와 모순되지 않게 쓴다.
3. 같은 파일에 시트 추가 (**`--outdir` 지정 금지, 단위 옵션 필수**):
   ```
   python plugins/dart-kospi-financials/skills/investment-thesis-writer/scripts/build_thesis_sheet.py \
       <xlsx> content.json --unit-label 백만달러
   python plugins/dart-kospi-financials/skills/investment-thesis-writer/scripts/build_valuation_sheet.py \
       <xlsx> valuation_content.json --unit-label 백만달러 --unit-multiplier 1000000
   ```
   ⚠ `--unit-multiplier 1000000`을 빠뜨리면 주당 내재가치가 100배 틀어진다(기본값이 KR 억원=1e8).
4. **가치평가_결론**:
   ```
   python plugins/valuation-verdict/skills/valuation-verdict/scripts/valuation_verdict.py \
       --source sec --cache-dir <이 스킬의 cache 폴더> --ticker <티커> \
       --xlsx <같은 xlsx> [--assumptions assumptions.json]
   ```
   가정(rf=미국채 10년, β, ERP 4~6%, 목표 PER)은 1번 리서치로 채워 `--assumptions`로 넘긴다. 결과 해석은 valuation-verdict SKILL.md 3~4번을 따른다.
5. 마지막에 `python /mnt/skills/public/xlsx/scripts/recalc.py <xlsx>`로 수식 오류를 확인한다. **`#N/A`는 "해당 기간 데이터 없음" 표시로 의도된 값이다**(분기 데이터가 성긴 과거 연도 등) — N/A 외의 오류(#VALUE!, #REF! 등)가 0인지 확인하고, N/A가 어느 시트에 몇 개인지 요약에 알린다.

## 5. 여러 기업 비교

```
python scripts/build_comparison_workbook.py AAPL:Apple MSFT:Microsoft --outdir /mnt/user-data/outputs
```
각 기업의 `fetch_financials.py` 캐시가 먼저 있어야 한다. `12분기 비교`/`최근5년 비교` 시트에 지표 17개 표+차트.

## 6. 저장 및 전달

완료 요약: 채워진 분기/연도 수, 못 찾은 계정 목록(`missing_indicators`), 주가 이력 사용 여부(`price_history_used`), 4단계까지 했다면 리서치 출처·가정 근거·**가치평가_결론의 한 줄 판정**.

### 구현 불변 규칙 (스크립트 수정 시 유지 — 실제 버그에서 나온 규칙)

1. **duration 필터**: 같은 계정에 3/6/9/12개월 값이 섞여 오므로 (end−start) 일수로 분기(80~100일)/연간(350~380일)을 거른다. **BS(instant)는 duration 계정의 실제 날짜 집합에 매칭되는 것만** 각 축에 넣는다(진행 분기말 스냅샷이 연간에 섞이는 버그 방지).
2. **4분기 파생**: 10-K에는 Q4 3개월 값이 보통 없다. 같은 회계연도 분기 3개가 있을 때만 Q4 = 연간−(Q1+Q2+Q3)로 파생하고, 3개가 안 모이면 지어내지 않는다. 주당 지표(EPS)는 파생하지 않는다.
3. **태그 후보 병합**: 첫 매칭 태그에서 멈추지 말고, 앞 태그가 비운 날짜만 뒤 태그로 보충한다(회사가 연도별로 다른 태그를 쓰는 경우 대비). 같은 (end, 기간)에 여러 공시가 있으면 **filed 최신** 값(정정 반영).
4. **EPS는 백만달러 환산 금지**(`PER_SHARE_KEYS`) — 위반 시 PER가 100만 배 틀어진다. 단위 폴백: USD → USD/shares.
5. PER/PBR/PSR은 완제품(trailingPE)을 믿지 않고 SEC 재무값+주가로 직접 수식 계산한다.
6. 빈 값은 0이 아니라 빈 칸/NA()(차트 끊김), 조정 계수는 셀 참조, 콤보 차트는 1차=막대+보조축=꺾은선, 실효세율은 세전이익≤0이면 NA·0~50% 클램프, 차트 열은 동적 계산 — KR "구현 불변 규칙"과 동일.
7. **N섹션 판정에서 순이익·자본이 0 이하인 연도의 PER/PBR은 제외**한다(음수 배수가 "저평가 충족"으로 오판되는 버그 방지).
8. 시트 라벨·레이아웃(지표 시트의 "지표" 헤더, 투자분석 L섹션의 종가/시가총액 행 등)은 **investment-thesis-writer·valuation-verdict가 그대로 읽는 계약**이다 — 바꾸면 그 두 스킬이 조용히 빈 시트를 만든다.

## 테스트

```
cd plugins/valuation-verdict/skills/valuation-verdict && python -m unittest discover -s tests -v
```
합성 SEC 픽스처(FAKE)가 이 플러그인의 수집·워크북 경로까지 함께 검증한다.
