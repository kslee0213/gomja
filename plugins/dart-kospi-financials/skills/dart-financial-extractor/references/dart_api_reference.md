# DART Open API 참조

## 사용 엔드포인트

| 목적 | 엔드포인트 | 비고 |
|---|---|---|
| 고유번호 전체 다운로드 | `GET /api/corpCode.xml` | ZIP 응답. 상장·비상장 전체 기업 목록(`corp_code`, `corp_name`, `stock_code`, `modify_date`). `stock_code`가 공백이 아니면 상장기업. |
| 기업개황 | `GET /api/company.json` | `corp_cls`(법인구분: Y=코스피, K=코스닥, N=코넥스, E=기타)로 시장 구분 확인. |
| 단일회사 전체 재무제표 | `GET /api/fnlttSinglAcntAll.json` | 이 플러그인의 핵심 API. 정기보고서 1건의 전 계정과목을 반환. |

베이스 URL: `https://opendart.fss.or.kr`

## reprt_code (보고서 코드)

| 코드 | 보고서 | 대상 기간 |
|---|---|---|
| `11013` | 1분기보고서 | 1~3월 |
| `11012` | 반기보고서 | 1~6월(누적) |
| `11014` | 3분기보고서 | 7~9월(단독) / 1~9월(누적) |
| `11011` | 사업보고서 | 1~12월(누적) |

## fs_div (재무제표 구분)

| 코드 | 의미 |
|---|---|
| `CFS` | 연결재무제표 (Consolidated) |
| `OFS` | 별도(개별)재무제표 (Separate) |

지주사가 없거나 자회사가 없는 중소형 상장사는 `CFS` 응답이 비어 있을 수 있다. 이 경우 `OFS`로 재조회한다.

## fnlttSinglAcntAll.json 응답 `list` 필드

| 필드 | 의미 |
|---|---|
| `rcept_no` | 접수번호 |
| `bsns_year` | 사업연도 |
| `corp_code` | 고유번호 |
| `sj_div` | 재무제표구분 (BS/IS/CIS/CF/SCE) |
| `sj_nm` | 재무제표명 |
| `account_id` | 계정ID (XBRL 표준계정, 매칭 기준으로 사용) |
| `account_nm` | 계정명 |
| `account_detail` | 계정상세 |
| `thstrm_nm` | 당기명 (예: "제 54 기 3분기") |
| `thstrm_amount` | 당기금액 — BS는 시점값, IS/CIS/CF는 보고서 유형별로 의미가 다름(본문 SKILL.md 4절 참고) |
| `thstrm_add_amount` | 당기누적금액 — 1·3분기보고서의 IS/CIS/CF 항목에만 존재, 연초~해당 시점 누적 |
| `frmtrm_nm` / `frmtrm_amount` | 전기명 / 전기금액 |
| `frmtrm_q_nm` / `frmtrm_q_amount` | 전기명(분/반기) / 전기금액(분/반기) |
| `frmtrm_add_amount` | 전기누적금액 |
| `bfefrmtrm_nm` / `bfefrmtrm_amount` | 전전기명 / 전전기금액 |
| `ord` | 계정과목 정렬순서 |
| `currency` | 통화 단위 |

## sj_div (재무제표구분)

| 코드 | 재무제표명 |
|---|---|
| `BS` | 재무상태표 |
| `IS` | 손익계산서 |
| `CIS` | 포괄손익계산서 |
| `CF` | 현금흐름표 |
| `SCE` | 자본변동표 (이 플러그인에서는 사용하지 않음) |

## status 코드 (오류)

| 코드 | 의미 |
|---|---|
| `000` | 정상 |
| `013` | 조회된 데이터가 없음 (해당 보고서 미제출 등) |
| `020` | 사용한도 초과 |
| `100` | 필드 누락/형식 오류 |
| `800` | 시스템 점검 |

## 참고 문서

- DART Open API 개발가이드: https://opendart.fss.or.kr/guide/main.do
