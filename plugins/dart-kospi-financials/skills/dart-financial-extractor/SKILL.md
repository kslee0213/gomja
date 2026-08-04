---
name: dart-financial-extractor
description: >
  This skill should be used when the user asks to fetch, extract, or automate
  Korean corporate financial statements from DART (전자공시), such as
  "OO 재무제표 뽑아줘", "다트에서 OO 재무제표 가져와줘", "OO DART 재무제표 자동화 실행",
  "OO 최근 12분기 재무제표", "OO 최근 5개년 사업보고서 재무제표 정리해줘", or any request
  to build a quarterly/annual financial statement Excel file for a KOSPI-listed company
  using the DART Open API.
metadata:
  version: "0.1.0"
---

# DART 재무제표 추출 및 엑셀 생성

기업명 하나를 입력받아 DART Open API로 재무제표를 수집하고, 최근 12개 분기와 최근 5개 사업연도 재무제표를 2개 시트 엑셀 파일로 만든다. 아래 순서를 그대로 따른다.

## 0. 설정 영역 (Configuration)

아래 값들은 이 스킬의 조정 가능한 설정이다. 사용자가 별도로 요청하지 않는 한 기본값을 사용한다.

| 설정 항목 | 기본값 | 설명 |
|---|---|---|
| `DART_API_KEY` | (환경 변수) | `os.environ["DART_API_KEY"]`로 읽는다. 없으면 실행 전에 사용자에게 요청하고, 응답에도 파일에도 값 자체를 노출하지 않는다. |
| `FS_DIV` | `CFS` (연결재무제표) | 응답의 `list`가 비어 있으면 `OFS`(별도재무제표)로 자동 재시도한다. 어떤 구분을 썼는지 최종 산출물에 표기한다. |
| `분기 수` | 12 | 최근 N개 분기. |
| `연도 수` | 5 | 최근 N개 사업연도. |
| 분기 시트명 | `분기_재무제표` | 사용자가 명시한 이름 그대로 고정. |
| 연간 시트명 | `연간_재무제표` | 사용자가 명시한 이름 그대로 고정. |
| 파일명 규칙 | `{기업명}_{YYYYMMDD}.xlsx` | YYYYMMDD는 실행 당일(로컬 날짜). |
| 시장 구분 확인 | KOSPI(코스피, corp_cls='Y')만 허용 | company.json API로 확인, 코스닥/코넥스면 사용자에게 확인 후 진행 여부를 묻는다. |

## 1. 기업 고유번호(corp_code) 조회

DART API는 기업명이 아니라 8자리 `corp_code`로 조회한다.

1. `scripts/corp_code_lookup.py`를 실행해 `corpCode.xml`(캐시가 7일 이상 오래됐으면 `https://opendart.fss.or.kr/api/corpCode.xml`에서 재다운로드, ZIP)에서 기업명을 매칭한다.
2. 동일한 이름이 여러 개면(자회사, 유사 상호 등) 후보 목록을 사용자에게 보여주고 선택받는다.
3. `company.json` API(`https://opendart.fss.or.kr/api/company.json`)로 `corp_cls`(법인구분: Y=코스피, K=코스닥, N=코넥스, E=기타)를 확인한다. Y가 아니면 사용자에게 알리고 계속할지 확인한다.

## 2. 대상 연도·보고서 목록 산정

- 연간 시트: 오늘 기준 최근 5개 사업연도의 사업보고서(`reprt_code=11011`).
- 분기 시트: 최근 12개 분기를 채우기 위해 필요한 보고서 목록을 만든다. 한 사업연도에는 4종 보고서가 필요하다.
  - `11013` 1분기보고서 (1~3월)
  - `11012` 반기보고서 (1~6월 누적)
  - `11014` 3분기보고서 (7~9월, 동시에 1~9월 누적 필드도 포함)
  - `11011` 사업보고서 (1~12월 누적)
- 아직 공시되지 않은 최신 분기(예: 반기보고서 법정 제출기한 전)는 건너뛰고, 그 이전 분기부터 역순으로 12개를 채운다. API 응답의 `status`가 `013`(조회된 데이터 없음)이면 해당 보고서가 아직 없는 것으로 간주한다.

## 3. 재무제표 원자료 조회

`scripts/fetch_financials.py`로 각 (연도, reprt_code) 조합마다 `fnlttSinglAcntAll.json`을 호출한다.

```
GET https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json
    ?crtfc_key={DART_API_KEY}&corp_code={corp_code}
    &bsns_year={연도}&reprt_code={보고서코드}&fs_div={FS_DIV}
```

- 응답 `list`의 각 원소를 `sj_div` 값으로 분류한다: `BS`=재무상태표, `IS`=손익계산서, `CIS`=포괄손익계산서, `CF`=현금흐름표. (`SCE` 자본변동표는 이번 산출물 범위 밖이므로 제외.)
- 계정을 매칭할 때는 `account_nm`(계정명)이 아니라 `account_id`(XBRL 표준계정ID)로 매칭한다. 같은 계정이라도 보고서마다 라벨 표기가 미세하게 다를 수 있다.
- 원자료는 재계산 방지를 위해 `cache/{corp_code}_{연도}_{reprt_code}_{fs_div}.json`에 저장해두고 재사용한다.
- 자세한 필드 설명과 reprt_code/sj_div 코드표는 `references/dart_api_reference.md`를 참고한다.

## 4. 분기 실적 계산 — 반드시 아래 구분을 지킨다

**재무상태표(BS)는 시점(저량) 데이터다.** 특정 시점의 잔액이므로 각 분기 말 보고서의 `thstrm_amount`를 그대로 쓴다. 뺄셈을 하지 않는다.

**손익계산서·포괄손익계산서·현금흐름표(IS/CIS/CF)는 흐름(유량) 데이터다.** DART 응답에는 두 금액 필드가 함께 온다.

- `thstrm_amount` (당기금액): 1분기·3분기보고서에서는 **해당 분기 단독(3개월)** 금액, 반기보고서에서는 **반기 누적(6개월)** 금액, 사업보고서에서는 **연간 누적(12개월)** 금액.
- `thstrm_add_amount` (당기누적금액): 1분기·3분기보고서에만 존재하며, **연초부터 해당 시점까지의 누적** 금액(1분기는 3개월 누적=단독과 동일, 3분기는 9개월 누적).

이를 바탕으로 분기별 실적은 다음과 같이 계산한다.

- **1분기** = 1분기보고서 `thstrm_amount` (그대로 사용, 계산 불필요)
- **2분기** = 반기보고서 `thstrm_amount`(6개월 누적) − 1분기보고서 `thstrm_amount`(3개월)
- **3분기** = 3분기보고서 `thstrm_amount` (이미 3개월 단독 값이므로 그대로 사용, 계산 불필요)
- **4분기** = 사업보고서 `thstrm_amount`(12개월 누적) − 3분기보고서 **`thstrm_add_amount`**(9개월 누적)

> ⚠️ 4분기 계산의 흔한 실수: 3분기보고서의 `thstrm_amount`(3개월 단독)를 사업보고서 값에서 빼면 안 된다. 그러면 "12개월 − 3개월"이 되어 4분기가 아니라 "1·2·4분기 합"이 산출된다. 반드시 3분기보고서의 **9개월 누적 필드(`thstrm_add_amount`)**를 사업보고서 값에서 빼야 4분기 단독 실적이 나온다.
>
> 실행 전 삼성전자 등 데이터가 풍부한 종목 1개로 `thstrm_add_amount` 필드가 실제로 채워져 있는지 먼저 확인하고(비어 있는 경우가 드물게 있음), 비어 있으면 대체 로직(1~2분기 합을 반기 누적값과 대조해 검증한 뒤, 3분기 누적 = 반기 누적 + 3분기 단독으로 직접 계산)을 사용한다.

## 5. 엑셀 작성

`scripts/build_workbook.py`가 2개 시트를 만든다. 반드시 `/mnt/skills/public/xlsx/SKILL.md`의 규칙(전문적인 글꼴, 수식 사용, 색상 규칙, `recalc.py` 무결성 검증)을 함께 따른다.

- **분기_재무제표**: 행 = 계정과목(재무상태표 → 손익계산서 → 현금흐름표 순서, 각 섹션은 `sj_nm`으로 구분), 열 = 최근 12개 분기(오래된 순 → 최근 순). 원본 누적 값(반기 누적, 9개월 누적, 연간)은 별도의 "원본데이터" 보조 열/시트에 입력값으로 두고, 2·4분기 열은 `=반기누적셀-1분기셀` 같은 실제 엑셀 수식으로 계산해 감사(audit) 가능하게 만든다. 하드코딩된 숫자로 2·4분기를 채우지 않는다.
- **연간_재무제표**: 행 = 계정과목, 열 = 최근 5개 사업연도. 사업보고서의 `thstrm_amount`를 그대로 입력값으로 사용한다.
- 각 시트 상단에 `fs_div`(연결/별도), 데이터 출처(DART corp_code, 접수번호 `rcept_no`), 생성일을 메모로 남긴다.
- 작성 후 `python scripts/office/recalc.py output.xlsx` (xlsx 스킬 경로 기준)를 실행해 수식 오류가 없는지 검증한다.

## 6. 저장 및 전달

- 파일명: `{기업명}_{YYYYMMDD}.xlsx` (예: `삼성전자_20260804.xlsx`)
- Cowork의 출력 폴더에 저장하고 사용자에게 전달한다.
- 완료 후 요약: 어떤 fs_div(연결/별도)를 사용했는지, 12분기 중 실제로 채워진 분기 수(공시 지연으로 못 채운 경우 몇 분기가 비었는지), 사용한 corp_code를 간단히 알려준다.

## 오류 처리

- `status != '000'`인 API 응답은 `message` 필드를 그대로 사용자에게 보고하고 중단한다(예: `013` 데이터 없음, `020` 사용 한도 초과).
- 사용 한도(요청 제한) 초과 시 재시도 간격을 두고, 계속 실패하면 사용자에게 상태를 알린다.
- 동일 기업명이 여러 corp_code에 매칭되면 반드시 사용자 확인을 받은 후 진행한다.
