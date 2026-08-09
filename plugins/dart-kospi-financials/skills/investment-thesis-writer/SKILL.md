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
  a target holding period and expected annual return).
metadata:
  version: "1.1.0"
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

## 참고

- 이 스킬은 `dart-financial-extractor`가 만든 파일에 시트를 "추가"하는 것이라 그 스킬과 항상 짝을 이뤄 쓰인다(독립적으로 원본 재무 데이터를 처음부터 수집하지 않는다).
- 재무 추세 표의 예측 구간은 시나리오이지 사실이 아니다. 시트에도 이 사실을 명시하는 각주가 자동으로 들어가지만, 대화에서도 사용자에게 "추정치"임을 한 번 더 짚어준다.
