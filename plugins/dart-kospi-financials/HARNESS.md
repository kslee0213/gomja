# dart-kospi-financials — Claude 하네스(Harness) 구성 가이드

이 문서는 플러그인을 "Claude 하네스" 형태로 볼 때, 어떤 **스킬(판단·오케스트레이션)**과
어떤 **스크립트(결정론적 도구)**가 어떤 순서/책임으로 물리는지 정리한다.
(기준 버전: plugin v0.19.0 — dart-financial-extractor 0.19.0, investment-thesis-writer 1.6.0)

## 1. 핵심 설계 원칙: 결정론 ↔ 비결정론 분리

| 성격 | 담당 | 재현성 | 검증 방법 |
|---|---|---|---|
| **결정론(deterministic)** | `scripts/*.py` (API 호출, 분기계산, 엑셀 수식, 차트) | 같은 입력=같은 출력 | `recalc.py`(수식 오류 0 확인) + `missing_indicators`(계정 매칭 실패 목록) |
| **비결정론(non-deterministic)** | Claude가 하는 일: 웹 서치, 성장률·할인율 가정, 정성 텍스트 | 매번 달라짐 | SKILL.md의 "구현 불변 규칙" 준수 + 기본값 사용 시 "이 정보들은 정확하지 않습니다" 명시 + 사람 검토 |

> 핵심: "매 요청마다 웹 서치 결과가 달라진다"는 문제는 **없앨 수 없다.**
> 그래서 값을 지어내지 않고, 가정과 출처를 산출물에 그대로 열거해 사람이 검토할 수 있게 한다.

## 2. 스킬 2개 (역할 = 하네스의 '두뇌'/오케스트레이터)

| 스킬 | 책임 | 결정론? |
|---|---|---|
| `dart-financial-extractor` | DART 수집 → 분기계산 → 엑셀(재무제표/지표/투자분석) 작성 오케스트레이션 | 대부분 결정론(스크립트 위임). 계정명 매칭만 후보탐색 |
| `investment-thesis-writer` | 사업내용 원문 + **웹 서치** + 투자분석 결과 종합 → 정성 텍스트/가정 작성 → 시트 추가 | **비결정론 핵심** |

이 플러그인 밖의 별도 플러그인 `valuation-verdict`(스킬 `valuation-verdict`)가 파이프라인의 마지막 단계로
연결된다(v0.18.0부터 "OO 년 분석해줘" 기본 파이프라인에 포함). 이 플러그인의 `cache/`와 워크북을 입력받아
"가치평가_결론" 시트를 맨 앞에 추가한다.

> 참고: 과거에 검수 게이트 스킬 `thesis-auditor`를 두었으나, 분기 연환산 기능과 함께 삭제했다
> (dart-financial-extractor CHANGELOG 참조). 현재 검수는 위 표의 "검증 방법"대로 수행한다.

## 3. 스크립트 인벤토리 (역할 = 하네스의 '손·발' 도구)

### dart-financial-extractor/scripts
| 스크립트 | 입력 | 출력 | 단계 |
|---|---|---|---|
| `corp_code_lookup.py` | 기업명 | corp_code, 종목코드, company.json | 1 |
| `fetch_financials.py` | corp_code·연도·보고서코드 | 재무 원자료 JSON(cache) | 3 |
| `fetch_extra_disclosures.py` | corp_code·연도 | 배당·대주주·자기주식·주식총수(cache) | 4-1 |
| `fetch_stock_price.py` | 종목코드·날짜(KRX키) | 종가·시총·상장주식수 | 4-1(선택) |
| `build_workbook.py` | corp_code·기업명·period | 단일기업 엑셀(전 시트) | 5 |
| `build_comparison_workbook.py` | 여러 corp_code | 비교 엑셀 | 5(비교) |

### investment-thesis-writer/scripts
| 스크립트 | 입력 | 출력 |
|---|---|---|
| `fetch_business_description.py` | corp_code | 사업의 내용 원문(cache) |
| `build_thesis_sheet.py` | 기존 xlsx + content.json | "투자판단 종합" 시트 추가 |
| `build_valuation_sheet.py` | 기존 xlsx + content.json | "버핏멍거_가치평가" 시트 추가 |

### (별도 플러그인) valuation-verdict/scripts
| 스크립트 | 입력 | 출력 |
|---|---|---|
| `valuation_verdict.py` | `--source dart` + cache 폴더 + corp_code (+ `--xlsx`, `--assumptions`) | "가치평가_결론" 시트 추가, `verdict_{회사}.json`, 마크다운 요약 |
| `adapters.py` | dart/sec/xlsx/json 소스 | 정규화 입력(내부 모듈) |

## 4. 하네스 파이프라인 (전체 요청 처리 순서)

```
[사용자 요청] "OO 년 분석해줘" (+ 투자판단까지)
        │
        ▼
┌─────────────────────────── dart-financial-extractor ───────────────────────────┐
│ 1 corp_code_lookup.py        (결정론)                                            │
│ 2 대상 연도/보고서 산정       (결정론)                                            │
│ 3 fetch_financials.py ×N     (결정론, cache; 진행연도 분기/반기 포함)             │
│ 4 분기실적 계산 규칙          (결정론)                                            │
│ 4-1 fetch_extra_disclosures / fetch_stock_price  (결정론, 선택)                  │
│ 5 build_workbook.py --period annual  → 재무제표+지표+투자분석 시트  (결정론)      │
└──────────────────────────────────────────────────────────────────────────────┘
        │  (전체 범위 선택 시 자동 연결)
        ▼
┌─────────────────────────── investment-thesis-writer ───────────────────────────┐
│ a fetch_business_description.py            (결정론)                              │
│ b 웹 서치 (산업/경쟁사/점유율/리스크)       ★비결정론★                            │
│ c 투자분석 시트 수치 반영 → content.json 작성 ★비결정론(가정·텍스트)★             │
│ d build_thesis_sheet.py / build_valuation_sheet.py  → 시트 추가  (결정론)        │
└──────────────────────────────────────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────── valuation-verdict (별도 플러그인) ──────────────────────┐
│ e 가정(rf·β·ERP·목표 PER 등) 리서치 → assumptions.json   ★비결정론★             │
│ f valuation_verdict.py --source dart --xlsx <워크북>  → "가치평가_결론" 시트 (결정론)│
│   기본값을 쓴 가정은 결과에 열거하고 "이 정보들은 정확하지 않습니다" 표기          │
└──────────────────────────────────────────────────────────────────────────────┘
        │
        ▼
 recalc.py 로 수식 오류 0 확인 (결정론)
        │
        ▼
[전달] 엑셀(7개 시트: 가치평가_결론·연간_재무제표·지표_연간·투자분석·원본데이터·
       투자판단 종합·버핏멍거_가치평가) + 가정/기본값 목록
```

## 5. 검수에 대한 현재 입장

- 자동 검증은 **결정론 영역**에 한정된다: `recalc.py`(수식 오류), `missing_indicators`(계정 매칭 실패).
- 단위·가정·텍스트 경계에서 났던 과거 버그(시가총액 단위, 기대성장률 %p vs 분수, 등급과 정성 텍스트의 모순, DCF 발산 조건 등)는
  각 SKILL.md의 "구현 불변 규칙"으로 승격해 두었고, 자동 게이트로 강제하지는 않는다.
- 따라서 웹 서치·가정에 기반한 부분은 산출물에 출처·가정·기본값 사용 여부를 남기고, 최종 판단은 사람 검토로 넘긴다.

## 6. 디렉터리 구조

```
plugins/dart-kospi-financials/
├── .claude-plugin/plugin.json            # v0.19.0
├── README.md
├── HARNESS.md                            # ← 이 문서
└── skills/
    ├── dart-financial-extractor/         # [수집·계산·엑셀]  결정론 오케스트레이터
    │   ├── SKILL.md
    │   ├── CHANGELOG.md
    │   ├── references/dart_api_reference.md
    │   └── scripts/ (corp_code_lookup, fetch_financials, fetch_extra_disclosures,
    │                 fetch_stock_price, build_workbook, build_comparison_workbook)
    └── investment-thesis-writer/         # [정성·가정·웹서치]  비결정론 작성기
        ├── SKILL.md
        └── scripts/ (fetch_business_description, build_thesis_sheet, build_valuation_sheet)

plugins/valuation-verdict/                # [결론]  별도 플러그인, 파이프라인 마지막 단계
└── skills/valuation-verdict/ (SKILL.md, scripts/valuation_verdict.py, scripts/adapters.py,
                               examples/, tests/)
```
