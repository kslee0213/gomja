---
name: investment-thesis-writer
description: >
  Use this skill when the user asks for a qualitative/narrative investment thesis
  summary sheet, such as "OO 투자판단 종합해줘", "OO 사업역량 분석해줘",
  "OO 경쟁우위/리스크/성장성 예측 작성해줘", or "OO 투자 스토리 정리해줘". This adds
  a "투자판단 종합" sheet to an existing DART 재무분석 워크북, combining DART
  business-description text, live web research, and the already-computed
  "투자분석" sheet results into a one-page narrative investment thesis
  (industry analysis, sales/production process, competitive advantage, risk,
  predictable period, growth/profitability forecast, and final conclusion with
  a target holding period and expected annual return). Also use for
  value-investing valuation requests such as "OO 내재가치 계산해줘",
  "OO DCF 해줘", "OO 안전마진 확인해줘", "OO 오너어닝 계산해줘",
  "OO 해자 분석해줘" — these add a "버핏멍거_가치평가" sheet to the same workbook
  (오너 어닝, 그레이엄 내재가치·안전마진, 시장 내재 기대성장률 역산, ROIC vs WACC,
  다년도 DCF 내재가치 계산기, 경제적 해자 체크리스트).
metadata:
  version: "1.6.0"
---

# 투자판단 종합 시트 작성

`dart-financial-extractor` 스킬로 이미 만든 "투자분석" 워크북(연간 파일, `--period annual`)에 **"투자판단 종합"**이라는 새 시트를 하나 더 추가한다. 이 시트는 숫자 계산이 아니라 **글을 읽고 판단해서 쓰는 정성적 요약**이라, 이전 스킬과 성격이 다르다 — Claude가 직접 리서치하고 문장을 작성해야 완성된다.

## 언제 이 스킬을 쓰는가

- 사용자가 이미 `dart-financial-extractor`로 "OO 년 분석해줘"를 실행해서 "투자분석" 시트가 있는 xlsx 파일을 갖고 있어야 한다. 없으면 먼저 그 스킬로 만들라고 안내한다.
- "투자판단 종합", "사업역량", "경쟁우위", "투자 스토리" 같은 요청이 오면 이 스킬을 쓴다.

## 0. 사전 준비

| 항목 | 내용 |
|---|---|
| `DART_API_KEY` | 사업의 내용 원문 조회에 필요. 이미 대화에 있으면 재사용, 없으면 요청. |
| 원본 xlsx 경로 | 이미 만들어진 "OO_연간_YYYYMMDD.xlsx" 파일 경로. 사용자가 최근에 만들었으면 그 파일을 그대로 쓰고, 없으면 먼저 `dart-financial-extractor`의 절차로 만든다. |

## 1. DART 사업의 내용 원문 수집

```
python skills/investment-thesis-writer/scripts/fetch_business_description.py --api-key <DART_API_KEY> <corp_code>
```

- 최신 사업보고서에서 "II. 사업의 내용" 섹션 텍스트를 뽑아 `cache/bizdesc_{corp_code}.json`에 저장한다.
- 이 파싱은 회사마다 문서 구조가 달라 100% 정확하지 않을 수 있다. 텍스트가 비었거나 이상하면 (`text` 필드가 짧거나 재무 내용처럼 보이면) 사용자에게 알리고, 필요하면 DART 사이트 링크(`https://dart.fss.or.kr`)에서 직접 확인하도록 안내한다.
- 이 텍스트는 **참고 자료**일 뿐, 그대로 베껴 쓰지 않는다. 회사가 자기소개하는 문서라 우호적으로 쓰여 있을 수 있음을 감안한다.

## 2. 웹 리서치 (필수)

사업의 내용 원문은 보통 최신이 아니거나(사업보고서 제출 시점 기준) 회사 관점의 서술이라, 아래를 웹 검색으로 보강한다:

- 최근 산업 동향/전망 (해당 업종 성장률, 정책 변화 등)
- 경쟁사 현황 및 시장점유율 (가능하면 최신 수치)
- 최근 리스크 요인이 될 만한 뉴스(원자재 가격, 규제, 소송, 대규모 계약 등)
- 애널리스트 리포트나 업계 보고서가 검색되면 참고(단, 저작권 규정에 따라 원문을 그대로 인용하지 말고 요약·재구성한다)

검색은 "OO 시장 전망", "OO 경쟁사", "OO 시장점유율", "OO 리스크" 등으로 나눠서 여러 번 한다(한 번에 다 안 나옴). 찾은 출처 URL은 `sources` 필드에 남겨서 나중에 검증 가능하게 한다.

## 3. 기존 "투자분석" 시트 결과 반영

원본 xlsx의 "투자분석" 시트(B~N 섹션)에서 이미 계산된 수치— 특히 재무건전성/수익성/성장성 등급(N섹션), 위험 신호(C섹션), PER/PBR(L섹션) — 를 읽어서(openpyxl로 직접 셀 값을 읽거나, 앞서 대화에서 이미 알고 있으면 그 값을 재사용) 아래 텍스트 작성 시 근거로 활용한다. 정성적 서술과 정량적 결과가 서로 모순되지 않도록 한다(예: 재무건전성이 D등급인데 "재무구조가 매우 안정적"이라고 쓰면 안 됨).

## 4. content.json 작성

아래 스키마로 JSON 파일을 만든다(`build_thesis_sheet.py`의 `CONTENT_SCHEMA_EXAMPLE` 참고):

```json
{
  "company_name": "회사명", "stock_code": "종목코드",
  "industry_analysis": "...", "sales_process": "...", "competitive_advantage": "...",
  "production_process": "...", "production_process_extra": "...(선택)",
  "products": [{"name": "제품1", "description": "..."}, {"name": "제품2", "description": "..."}],
  "risk": "...", "predictable_period_text": "...", "growth_prediction_text": "...",
  "competitive_situation": "...", "profitability_prediction_text": "...",
  "final_conclusion_text": "...",
  "sustainable_period_years": 10, "expected_annual_return_pct": 12.0,
  "projection_years": 10,
  "revenue_growth_assumptions": [0.1, ...], "op_margin_assumptions": [0.1, ...], "net_margin_assumptions": [0.1, ...],
  "sources": ["https://..."]
}
```

작성 원칙:
- 각 텍스트 필드는 2~5문장 정도로, 첨부 참고 양식(성공투자노트)처럼 **결론부터 명확히** 쓴다(장황한 서론 없이 "~이다/~할 것으로 보인다"체).
- `revenue_growth_assumptions` 등 성장률 가정은 **근거 없이 임의로 지어내지 않는다** — 리서치한 산업 성장률, 회사의 과거 5개년 평균 성장률(투자분석 시트 B섹션 성장성 참고), 경쟁 환경을 종합해 정한 값이어야 하고, `growth_prediction_text`/`profitability_prediction_text`에 그 가정의 근거를 설명한다.
- `expected_annual_return_pct`는 현재 PER과 목표 PER(성장 시 재평가 가능한 수준) 가정, 이익 성장률을 결합해 추정한다. 계산 과정을 `final_conclusion_text`에 요약한다.
- 제품(`products`)은 사업보고서에서 언급된 주요 제품/서비스 1~3개를 넣는다(파악 안 되면 빈 배열로 두어도 됨 — 스크립트가 알아서 건너뛴다).

## 5. 시트 생성

```
python skills/investment-thesis-writer/scripts/build_thesis_sheet.py <원본_연간.xlsx 경로> <content.json 경로>
```

- **`--outdir`를 지정하지 않는 것이 기본(권장) 사용법이다.** 그러면 새 파일을 만들지 않고 원본 xlsx 자체에 "투자판단 종합" 시트를 추가해 같은 파일로 덮어쓴다 — 최종 결과물이 "재무제표+투자분석+투자판단 종합"을 전부 담은 파일 하나가 된다. `dart-financial-extractor`의 5-3단계에서 이어서 실행할 때는 항상 이 방식을 쓴다.
- 원본 파일을 보존하고 별도 파일로 받고 싶다는 요청이 명시적으로 있을 때만 `--outdir /mnt/user-data/outputs`를 붙인다(이 경우 `{회사명}_투자판단종합_{YYYYMMDD}.xlsx`라는 새 파일이 만들어진다).

- 기본 동작(`--outdir` 미지정)은 원본 파일 자체에 "투자판단 종합" 시트를 추가해 같은 파일명으로 덮어쓴다(같은 이름의 시트가 이미 있으면 그 시트만 교체). `--outdir`를 지정했을 때만 `{회사명}_투자판단종합_{YYYYMMDD}.xlsx`라는 새 파일이 별도로 만들어진다.
- 장기 재무추세 표는 과거 실적(최근 5개년, `지표_연간` 시트를 참조하는 수식이라 감사 가능) + 예측 구간(제시한 성장률 가정을 복리 계산하는 수식)으로 자동 구성된다. 하드코딩된 숫자가 아니라 전부 수식이므로, 사용자가 엑셀에서 가정치 셀만 바꾸면 전체 표가 다시 계산된다.
- 완료 후 사용자에게: (1) 어떤 웹 검색으로 리서치했는지 출처 요약, (2) 성장률 가정의 근거, (3) DART 사업의 내용 파싱이 잘 안 됐다면 그 사실을 알려준다.

## 6. "버핏멍거_가치평가" 시트 (v1.4.0부터)

첨부 참고자료("워렌버핏과 찰리멍거")의 계산들을 코드화한 **세 번째 시트**다. `투자판단 종합`이 완성된 뒤(또는 없어도 "투자분석" 시트만 있으면) 이어서 만들 수 있다. 5번 섹션과 마찬가지로 순수 계산은 전부 수식으로 채워지지만, 할인율·영구성장률·WACC·향후 오너 어닝 성장률 가정·해자 체크리스트 평가·종합 결론은 Claude가 판단해서 채워야 한다.

### 사전 준비

`skills/dart-financial-extractor/scripts/build_workbook.py`가 이미 "감가상각비" 계정을 "투자분석" 시트의 기초 참고값으로 채워두므로(v0.9.7부터), 별도 데이터 수집 없이 바로 진행할 수 있다. 다만 "투자분석" 시트가 오래된 버전(v0.9.6 이전)으로 만들어진 파일이면 감가상각비가 없어 오너 어닝 행이 비게 되니, 이 경우 사용자에게 최신 버전으로 재생성을 안내한다.

### content.json 작성

`skills/investment-thesis-writer/scripts/build_valuation_sheet.py`의 `CONTENT_SCHEMA_EXAMPLE` 참고. 핵심 필드:

- `wacc_pct`: 자본비용 가정(%). 베타 데이터가 없으므로 업종 평균(보통 6~10%)을 근거로 Claude가 제시한다.
- `discount_rate_pct`, `terminal_growth_pct`: DCF 할인율과 영구성장률. 할인율은 보통 WACC와 비슷하거나 약간 높게, 영구성장률은 장기 GDP 성장률(2~4%)을 넘지 않게 잡는다(그래야 TV가 발산하지 않는다).
- `projection_years`: DCF 예측 기간(보통 5~10년).
- `owner_earnings_growth_assumptions`: 연도별 오너 어닝 성장률 가정 리스트. 근거 없이 임의로 정하지 말고, "투자분석"의 과거 5개년 성장률과 "투자판단 종합"에서 이미 리서치한 산업 전망을 참고해 정한다.
- `moat_checklist`: 9개 항목(경쟁우위, 시장점유율, 경영진, 안전마진, 가격전가력, 신뢰성, 규제경험, 규모우위, 네트워크효과) 각각에 평가(강함/보통/약함/- )와 근거를 채운다. "투자판단 종합"에서 이미 리서치한 내용을 재사용하면 된다.
- `valuation_conclusion`: 안전마진(자산기반·성장주기반·DCF 세 가지가 나온다), ROIC-WACC 스프레드, 해자 체크리스트를 종합한 투자 의견.

### 실행

```
python skills/investment-thesis-writer/scripts/build_valuation_sheet.py <xlsx 경로> <content.json 경로>
```

`--outdir`를 지정하지 않으면(기본값) 같은 파일에 "버핏멍거_가치평가" 시트를 추가해 덮어쓴다 — `dart-financial-extractor`의 재무제표+투자분석, 그리고 `투자판단 종합`까지 전부 파일 하나에 모인다.

### 결과 해석 시 유의점

- 세 가지 안전마진(자산기반/성장주기반/DCF)이 서로 다른 값을 낼 수 있다. 이는 정상이며, 세 관점 중 어느 게 이 회사에 더 적합한지(자산이 무거운 회사면 자산기반, 성장주면 성장주기반, 현금흐름이 안정적이면 DCF) Claude가 판단해 종합 결론에 반영한다.
- DCF의 영구가치(TV)는 할인율과 영구성장률의 차이가 작을수록 급격히 커진다(발산 위험). `discount_rate_pct - terminal_growth_pct`가 3%p 미만이면 TV가 비정상적으로 크게 나올 수 있으므로 가정을 재검토하라고 안내한다.
- 완료 후 사용자에게: 세 안전마진 수치, ROIC-WACC 스프레드, 해자 체크리스트 요약, 그리고 DCF 가정(할인율·영구성장률·오너어닝 성장률)의 근거를 간단히 알려준다.

## 참고

### v1.6.0 (분기 연환산 모드 지원 제거 — 사용자 요청으로 원복)

`dart-financial-extractor`의 분기 연환산 기능(`연환산_재무제표`, `지표_연환산`, `투자분석_분기추정`)이 전체 제거되면서, 이 스킬의 `--period quarterly` 옵션도 함께 제거했다. `build_thesis_sheet.py`/`build_valuation_sheet.py`는 다시 연간 워크북(`지표_연간`/`투자분석`) 전용이다.

### v1.4.0 변경 사항

- **`build_valuation_sheet.py` 신규 추가**: "버핏멍거_가치평가" 시트(6번 섹션 참고)를 만든다. 오너 어닝, 그레이엄 내재가치·안전마진, 시장 내재 기대성장률 역산, ROIC vs WACC, 다년도 DCF, 해자 체크리스트를 계산한다.
- 단위 버그 2건을 실제로 테스트 중 발견해 수정함: (1) "상장주식수 역산" 계산에서 시가총액이 억원 단위인데 원 단위 종가와 그대로 나눠서 실제 주식수보다 1억분의 1로 작게 나오던 문제, (2) "시장 내재 기대성장률 = (PER−8.5)/2" 값이 이미 %포인트 단위인데 셀 서식을 %로 지정해서 100배로 표시되고, 그 값을 그대로 "실제 CAGR"(분수 단위)과 빼서 "괴리"를 계산해 단위가 안 맞던 문제. 두 값 모두 실제 계산해서 손으로 검산한 뒤 확인했다.

### v1.3.0 변경 사항

- **제품 분석 정렬**: 제품명(I열)과 설명(L열)이 같은 행에 오도록 수정(I13+L13, I21+L21). 예전엔 이름이 설명보다 아래 행에 떨어져 있었다.
- **재무추세표**: "연도"가 "연차"보다 먼저 나오도록 순서를 바꾸고, 표 시작 열을 D→C로 한 칸 당겼다. 시가총액·주가·PER·PBR 4행은 같은 워크북의 "투자분석" 시트 L섹션을 셀 참조 수식으로 채운다(과거 실적 연도만, 예측 구간은 미래 주가를 예측하지 않으므로 빈칸 유지).
- **테두리**: `apply_grid_border`/`apply_grid_border_range` 헬퍼로 모든 텍스트 박스(헤더+본문)와 재무추세표 전체에 얇은 테두리를 적용했다. 새 섹션을 추가할 때는 반드시 `section()` 헬퍼를 통해 만들거나, 직접 만드는 경우 `apply_grid_border_range()`를 꼭 호출해야 테두리가 붙는다.

- 이 스킬은 `dart-financial-extractor`가 만든 파일에 시트를 "추가"하는 것이라 그 스킬과 항상 짝을 이뤄 쓰인다(독립적으로 원본 재무 데이터를 처음부터 수집하지 않는다).
- 재무 추세 표의 예측 구간은 시나리오이지 사실이 아니다. 시트에도 이 사실을 명시하는 각주가 자동으로 들어가지만, 대화에서도 사용자에게 "추정치"임을 한 번 더 짚어준다.
