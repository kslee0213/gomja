# dart-kospi-financials — Claude 하네스(Harness) 구성 가이드

이 문서는 플러그인을 "Claude 하네스" 형태로 볼 때, 어떤 **스킬(판단·오케스트레이션)**과
어떤 **스크립트(결정론적 도구)**가 어떤 순서/책임으로 물리는지 정리한다.

## 1. 핵심 설계 원칙: 결정론 ↔ 비결정론 분리

| 성격 | 담당 | 재현성 | 검증 방법 |
|---|---|---|---|
| **결정론(deterministic)** | `scripts/*.py` (API 호출, 분기계산, 엑셀 수식, 차트) | 같은 입력=같은 출력 | recalc.py + auditor 그룹 A/B |
| **비결정론(non-deterministic)** | Claude가 하는 일: 웹 서치, 성장률·할인율 가정, 정성 텍스트 | 매번 달라짐 | **thesis-auditor** 그룹 C/D/E + 사람 검토 |

> 핵심: "매 요청마다 웹 서치 결과가 달라진다"는 문제는 **없앨 수 없고 붙잡아야 한다.**
> 그래서 값을 고치지 않고 **표시(flag)만** 하는 독립 검수 게이트(thesis-auditor)를 둔다.

## 2. 스킬 3개 (역할 = 하네스의 '두뇌'/오케스트레이터)

| 스킬 | 책임 | 결정론? |
|---|---|---|
| `dart-financial-extractor` | DART 수집 → 분기계산 → 엑셀(재무제표/지표/투자분석) 작성 오케스트레이션 | 대부분 결정론(스크립트 위임). 계정명 매칭만 후보탐색 |
| `investment-thesis-writer` | 사업내용 원문 + **웹 서치** + 투자분석 결과 종합 → 정성 텍스트/가정 작성 → 시트 추가 | **비결정론 핵심** |
| `thesis-auditor` (신규) | 위 두 스킬 산출물을 **전달 직전** 자동 검증, PASS/WARN/FAIL 리포트 | 결정론(웹 서치 안 함) |

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

### thesis-auditor/scripts (신규)
| 스크립트 | 입력 | 출력 |
|---|---|---|
| `audit_workbook.py` | xlsx (+thesis/valuation content.json) | 검수 JSON + 리포트 md, 종료코드(FAIL=2/WARN=1/PASS=0) |

## 4. 하네스 파이프라인 (전체 요청 처리 순서)

```
[사용자 요청] "OO 년 분석 + 투자판단까지"
        │
        ▼
┌─────────────────────────── dart-financial-extractor ───────────────────────────┐
│ 1 corp_code_lookup.py        (결정론)                                            │
│ 2 대상 연도/보고서 산정       (결정론)                                            │
│ 3 fetch_financials.py ×N     (결정론, cache)                                     │
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
        │  ★★★ 전달 직전 게이트 ★★★
        ▼
┌─────────────────────────────── thesis-auditor ─────────────────────────────────┐
│ audit_workbook.py <xlsx> --thesis-content ... --valuation-content ...           │
│   A 결정론 무결성(오류셀·회계항등식·현금흐름)                                     │
│   B 단위 일관성(PER/시총/주식수 자릿수 — 과거 실제 버그 지점)                     │
│   C 정량 vs 정성 모순(등급 D인데 "안정적" 등)   ← 비결정론 텍스트 검증            │
│   D 가정 경계값(DCF 스프레드·영구성장률·성장률 리스트 길이/근거)                  │
│   E 출처/재현성(sources 존재·URL·회사명 일치)   ← 웹 서치 검증                    │
│                                                                                 │
│   → PASS/WARN/FAIL 리포트                                                        │
│      FAIL → 원인 고쳐 앞 단계 재실행 (또는 사용자 명시 승인 후 전달)              │
│      WARN → "사람 확인 권장" 목록과 함께 전달 가능                                │
└──────────────────────────────────────────────────────────────────────────────┘
        │
        ▼
[전달] 엑셀 + 검수 리포트(md)를 함께 제공
```

## 5. 왜 검수 스킬이 "필요"한가 (결론)

- 기존 안전장치는 **결정론 영역**만 자동 검증한다: `recalc.py`(수식 오류), `missing_indicators`(계정 매칭 실패).
- 그러나 v0.9~v1.4 릴리스 노트가 스스로 기록하듯, **실제 버그는 대부분 단위/가정/텍스트 경계**에서 났다:
  - 시가총액 1억배 단위 꼬임 (v1.4.0에서 사후 발견)
  - 기대성장률 %포인트 vs 분수 혼동 (v1.4.0)
  - "재무건전성 D인데 안정적이라고 쓰면 안 된다"는 규칙은 SKILL.md 지침으로만 존재 → 자동 강제 없음
  - DCF 발산(할인율−영구성장률<3%p)은 안내만 있고 게이트 없음
- 게다가 **웹 서치 결과는 매 실행마다 달라져** 사람이 매번 전량 재검토하기 비현실적이다.
- → 이 규칙들을 **전달 직전 자동 게이트**로 코드화한 것이 `thesis-auditor`다.
  값을 고치지 않고 **잡아서 표시**만 하므로 검수의 독립성이 보장되고,
  "자동으로 못 잡는 것(사실관계 진위 등)"은 정직하게 사람 검토로 넘긴다.

## 6. 디렉터리 구조 (하네스 정리 후)

```
plugins/dart-kospi-financials/
├── .claude-plugin/plugin.json            # v0.10.0 (auditor 반영)
├── README.md
├── HARNESS.md                            # ← 이 문서
└── skills/
    ├── dart-financial-extractor/         # [수집·계산·엑셀]  결정론 오케스트레이터
    │   ├── SKILL.md
    │   ├── references/dart_api_reference.md
    │   └── scripts/ (corp_code_lookup, fetch_financials, fetch_extra_disclosures,
    │                 fetch_stock_price, build_workbook, build_comparison_workbook)
    ├── investment-thesis-writer/         # [정성·가정·웹서치]  비결정론 작성기
    │   ├── SKILL.md
    │   ├── examples/ (thesis_content, valuation_content)
    │   └── scripts/ (fetch_business_description, build_thesis_sheet, build_valuation_sheet)
    └── thesis-auditor/                   # [검수 게이트]  결정론 검증기 (신규)
        ├── SKILL.md
        └── scripts/audit_workbook.py
```
