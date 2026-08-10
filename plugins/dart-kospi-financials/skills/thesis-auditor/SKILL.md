---
name: thesis-auditor
description: >
  Use this skill to REVIEW and VALIDATE the outputs of dart-financial-extractor
  and investment-thesis-writer BEFORE the workbook is delivered to the user.
  Trigger it whenever a "투자분석", "투자판단 종합", or "버핏멍거_가치평가" sheet has just
  been generated (i.e. any run that involved web research or Claude-provided
  assumptions), or when the user explicitly asks "검수해줘", "검증해줘",
  "가정 근거 확인해줘", "수치 모순 없는지 봐줘", "audit". It runs deterministic
  automated checks (unit sanity, formula recalc, quant-vs-qualitative consistency,
  assumption-bounds, source coverage) and produces an audit report, flagging
  anything the run cannot self-guarantee because it depended on non-deterministic
  web-search content or free-form assumptions.
metadata:
  version: "1.0.0"
---

# 산출물 검수 (Thesis Auditor)

`dart-financial-extractor`와 `investment-thesis-writer`는 **결정론적 계산**(DART 원자료·수식)과
**비결정론적 판단**(웹 서치 결과·성장률/할인율 가정·정성 텍스트)을 한 파일에 섞어서 만든다.
전자는 스크립트 안의 recalc·매칭실패목록으로 어느 정도 자체 검증되지만, 후자는
"매 실행마다 내용이 달라지는데도" 이를 붙잡아 확인하는 독립 단계가 없다.

이 스킬은 **전달(delivery) 직전의 게이트**다. 워크북과 (있다면) content.json을 입력받아
아래 검사를 자동 실행하고, 통과/경고/실패를 담은 검수 리포트를 만든다. 사람이 모든 걸
다시 읽지 않아도 "무엇을 신뢰할 수 있고, 무엇을 사람이 눈으로 확인해야 하는지"를 분리해 준다.

> 원칙: 이 스킬은 값을 **고치지 않는다**. 발견·표시(flag)만 한다. 수정은 사람 또는
> 앞 스킬을 재실행해 처리한다. (검수자가 데이터를 조용히 바꾸면 검수의 의미가 사라진다.)

## 언제 쓰는가

- `dart-financial-extractor`의 5단계(엑셀 작성) 또는 5-3단계(투자판단 종합)가 끝난 직후, **사용자에게 파일을 전달하기 전에** 자동으로 이어서 실행하는 것을 기본으로 한다.
- `investment-thesis-writer`가 "투자판단 종합"/"버핏멍거_가치평가" 시트를 추가한 직후.
- 사용자가 명시적으로 "검수/검증/audit"을 요청할 때.

## 0. 입력

| 항목 | 필수 | 설명 |
|---|---|---|
| `<xlsx 경로>` | 필수 | 검수 대상 워크북 (재무제표+투자분석[+투자판단종합][+가치평가]) |
| `<content.json 경로들>` | 선택 | thesis/valuation content.json (있으면 가정·출처 검사가 훨씬 강해짐). 없으면 시트에서 읽을 수 있는 것만 검사한다. |

## 1. 검수 스크립트 실행

```
python skills/thesis-auditor/scripts/audit_workbook.py <xlsx 경로> \
    [--thesis-content <thesis_content.json>] \
    [--valuation-content <valuation_content.json>] \
    [--report-md /mnt/user-data/outputs/<회사명>_검수리포트_<YYYYMMDD>.md] \
    [--fail-on error]
```

스크립트는 **결정론적 규칙만** 검사한다(웹 서치를 다시 하지 않는다 — 검수 자체가 비결정론적이면 안 되므로).
검사 항목은 각각 `PASS` / `WARN` / `FAIL` 중 하나로 판정되고, 리포트(md)와 stdout JSON으로 나온다.

### 검사 항목 (스크립트가 자동 수행)

**그룹 A — 결정론적 무결성 (재무 데이터)**
- A1. 수식 재계산: `#REF!`, `#DIV/0!`, `#VALUE!` 등 오류 셀이 남아 있으면 FAIL. (xlsx 스킬의 recalc.py가 이미 돌았다는 전제이지만, 여기서 오류 셀 존재만 재확인.)
- A2. 회계 항등식: 자산총계 ≈ 부채총계 + 자본총계 (허용오차 0.5%). 어긋나면 WARN(계정 매칭 누락 신호).
- A3. 현금흐름 3단 합계 ≈ 현금및현금성자산의증가 (허용오차). 어긋나면 WARN.
- A4. 매칭 실패 지표 목록(빨간 글씨/`missing_indicators`)이 있으면 그대로 리포트에 나열(WARN, 몇 개 이상이면 안내).

**그룹 B — 단위 일관성 (과거 실제 버그가 났던 지점)**
- B1. 주가(종가)는 억원 변환에서 제외돼야 한다 — PER = 주가/EPS가 상식 범위(음수 아님, 통상 0<PER<300)인지 검사. 벗어나면 WARN(단위 꼬임 의심).
- B2. 상장주식수 역산·시가총액 단위(억원) 정합성 — 시가총액/주가 ≈ 상장주식수가 자릿수(order of magnitude) 내에서 맞는지. 1억배 어긋나면 FAIL.
- B3. 시장 내재 기대성장률=(PER-8.5)/2 는 %포인트 단위인데 실제 CAGR(분수)과 직접 빼면 안 됨 — 두 값의 자릿수 차가 비정상(예: 100배)이면 WARN.

**그룹 C — 정량 vs 정성 모순 (비결정론 텍스트 검증의 핵심)**
- C1. "투자분석" N섹션 등급과 "투자판단 종합" 텍스트의 방향이 반대인지 키워드로 검사.
  예: 재무건전성 등급이 D/E인데 최종결론/리스크 텍스트에 "안정적/탄탄/우량" 류 긍정어만 있으면 WARN.
  수익성 등급 D/E인데 "고수익/높은 마진" 서술이면 WARN.
- C2. 위험신호(C섹션)에서 ⚠ 위험이 여러 개인데 리스크 텍스트가 비었거나 지나치게 짧으면(예: 40자 미만) WARN.
- C3. `expected_annual_return_pct`(기대수익률)가 과거 5개년 순이익 CAGR + 배당수익률에 비해 과도하게 높으면(예: 과거 실적 대비 +2배 이상이거나 절대값 30% 초과) WARN + 근거 확인 요청.

**그룹 D — 가정 경계값 (Claude가 채운 숫자)**
- D1. DCF: `discount_rate_pct - terminal_growth_pct < 3`이면 WARN(TV 발산 위험). `<= 0`이면 FAIL.
- D2. `terminal_growth_pct`가 장기 GDP 상단(통상 4%)을 초과하면 WARN.
- D3. `wacc_pct`, `discount_rate_pct`가 상식 범위(예: 3%~20%) 밖이면 WARN.
- D4. `revenue_growth_assumptions` 등 성장률 리스트에 근거 텍스트(growth_prediction_text 등)가 비어 있으면 WARN("근거 없이 지어낸 가정" 방지). 리스트 길이가 projection_years와 안 맞으면 FAIL.
- D5. 오너어닝/성장률 가정이 과거 5개년 실적 범위를 크게 벗어나면(예: 과거 최대 성장률의 2배 초과) WARN + 근거 확인.

**그룹 E — 출처/재현성 (웹 서치 결과)**
- E1. content.json에 `sources` 배열이 비어 있거나(웹 리서치를 했다면서 출처가 없음) URL 형식이 아닌 항목이 있으면 WARN.
- E2. 예측/가정을 담은 텍스트 필드가 있는데 sources가 하나도 없으면 WARN(리서치 미수행 의심).
- E3. content.json의 `company_name`/`stock_code`가 워크북 파일명/시트의 회사와 일치하는지 — 불일치면 FAIL(엉뚱한 회사 content를 붙였을 위험).

## 2. 판정 종합 및 사용자 안내

스크립트 stdout의 JSON(`{"summary": {...}, "checks": [...]}`)을 읽고 다음을 사용자에게 보고한다:

1. **한 줄 결론**: `FAIL 0 / WARN n / PASS m` — FAIL이 하나라도 있으면 "전달 전 수정 필요", WARN만 있으면 "전달 가능하나 사람 확인 권장".
2. **FAIL 항목**: 무엇이·어느 셀/필드에서 틀렸는지와 어떻게 고치는지(어느 스크립트 재실행/어느 가정 수정).
3. **WARN 항목**: "이건 자동으로 판단 못 하니 사람이 확인하라"는 목록. 특히 **비결정론 항목**(C/D/E 그룹)은 매 실행마다 달라질 수 있으므로 항상 사람 눈 검토 대상임을 명시한다.
4. **검수 리포트 md 파일**을 산출물 폴더에 함께 저장(감사 추적용).

> ⚠️ FAIL이 있으면 파일을 그대로 "완료"로 전달하지 않는다. 원인을 고쳐 앞 스킬을 재실행하거나,
> 사용자에게 명시적으로 "그래도 이대로 받으시겠습니까"를 확인한 뒤 전달한다.

## 3. 검수는 검수만 한다 (경계)

- 이 스킬은 웹 서치를 하지 않고, 값을 고치지 않고, 새 시트를 만들지 않는다.
- 여기서 못 잡는 것(예: 웹 서치 사실관계 자체의 진위, 산업 전망의 타당성)은 **정직하게 "자동 검수 불가"로 표시**하고 사람 검토로 넘긴다. 검수기가 못 잡은 걸 잡은 척하지 않는다.
