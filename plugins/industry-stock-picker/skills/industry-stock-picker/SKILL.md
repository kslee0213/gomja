---
name: industry-stock-picker
description: >
  Use this skill when the user asks for stock RECOMMENDATIONS within an industry,
  the reverse direction of the per-company skills — e.g. "투자할 만한 반도체 기업
  추천해줘", "한국의 우주 산업 기업 추천해줘", "미국 조선주 뭐 살만해?", "OO산업
  저평가 종목 골라줘". It screens 6~8 listed candidates in that industry with the
  same Graham/Buffett-Munger style metrics and the valuation-verdict engine used by
  dart-kospi-financials / us-stocks-financials, ranks them with an auditable score,
  and recommends 3 companies with reasons. The user states the market (한국/미국)
  in the request; if not stated, ask. Requires the same API keys/network as the
  underlying extractor plugins.
metadata:
  version: "1.0.0"
---

# Industry Stock Picker (산업 → 종목 추천)

"기업명 → 분석"의 역방향: **산업명 → 후보 스크리닝 → 3곳 추천**. 계산은 기존 플러그인(dart/us 추출기 + valuation-verdict 엔진)을 그대로 재사용한다. 버전 이력은 `CHANGELOG.md`.

> 절대 규칙: 추천은 **공시 재무데이터 + 명시된 공통 가정에 종속된 조건부 순위**이며 투자 권유가 아니다. 시장(한국/미국)이 요청에 없으면 반드시 물어본다. 산업 대표성 있는 후보를 지어내지 말 것 — 후보 선정 근거(출처)를 남긴다.

## 1. 후보 선정 (Claude가 리서치로)

1. 웹 검색으로 해당 산업의 **상장사 6~8개**를 고른다(시가총액 상위 + 사업 순도 높은 기업 혼합). 각 후보가 왜 그 산업 대표인지 한 줄 근거와 출처를 기록한다.
2. 식별자를 확정한다 — 이름만으로 스크립트가 회사를 추측하게 하지 않는다:
   - KR: `python plugins/dart-kospi-financials/.../scripts/corp_code_lookup.py --api-key <DART키> <기업명>` → corp_code·stock_code (코스피/코스닥 구분 확인)
   - US: `python plugins/us-stocks-financials/.../scripts/ticker_lookup.py <티커> --contact <이메일>` → cik
3. `candidates.json` 작성 (`examples/candidates.example.json` 스키마).

## 2. 공통 가정 준비 (권장)

후보 **전원에 같은 가정**을 적용한다(상대 비교의 공정성). 국고채/미국채 10년 rf, 시장 ERP, 산업 평균 β로 `assumptions_kr.json`/`assumptions_us.json`을 만들고 근거를 남긴다. 생략하면 시장 기본값(r 9% 등)이 쓰이고 결과에 "이 정보들은 정확하지 않습니다"로 표시된다.
**한계(결과 보고에 반드시 명시)**: 종목별 β·목표 PER를 따로 주지 않으므로 절대 적정주가는 거칠고, 순위는 상대 비교용이다.

## 3. 스크리닝 실행

```
python scripts/screen_candidates.py candidates.json --fetch \
    --dart-key <키> --krx-key <키>          # KR 후보가 있을 때
    --contact <이메일>                        # US 후보가 있을 때
    [--assumptions-kr a_kr.json] [--assumptions-us a_us.json]
python scripts/build_screening_workbook.py screening.json --outdir /mnt/user-data/outputs
```

- `--fetch`: 후보별로 연간 5개년 + 진행연도 분기 + 주가를 기존 추출 스크립트로 수집(캐시 재사용). 후보 하나가 실패해도 나머지는 계속되고 `errors`에 남는다 — **실패 후보를 조용히 빼고 "전부 분석했다"고 말하지 않는다.**
- 점수(0~100): 가치(상승여력, 40) + 수익성 ROE(15) + 건전성 부채비율(10) + 이익의 질 OCF/NI(8) + FCF(7) + 성장(10) + 흑자연속(10), **결측 항목은 0점이 아니라 분모 제외**로 정규화, 신뢰도 낮음 −8/보통 −4. 산식은 워크북 "점수_산식" 시트에 수식으로 들어간다.
- 워크북 저장 후 `python /mnt/skills/public/xlsx/scripts/recalc.py <xlsx>`로 수식 오류 0 확인.

## 4. 최종 추천 3곳 (Claude가 판단)

정량 순위를 그대로 베끼지 말고 다음을 얹어 3곳을 고른다:

1. **해자(정성)**: 상위 4~5곳에 대해 웹 리서치로 경쟁우위·점유율·산업 사이클 위치를 확인(체크리스트는 investment-thesis-writer의 해자 항목 준용). 워크북 "참고_주의" 시트의 해자 칸에 근거와 출처를 채운다.
2. 신뢰도·경고 반영: 신뢰도 "낮음" 후보를 추천할 땐 그 사유를 명시. 정량 1위를 제치고 다른 후보를 올리면 **그 이유를 반드시 쓴다.**
3. 보고 형식: ① 추천 3곳 각 한 줄(현재가 vs 적정주가, 상승여력, 판정, 종합점수) ② 각 사 추천 근거 2~3개(정량+해자) ③ 각 사 핵심 리스크 ④ 탈락한 상위 후보와 그 이유 ⑤ 공통 가정·기본값 목록 + 해당 부분 "이 정보들은 정확하지 않습니다" ⑥ "조건부 추정이며 투자 권유가 아님".

심화 분석을 원하면: 추천된 기업에 대해 "OO 년 분석해줘"로 기존 스킬(dart/us 전체 파이프라인 → 9시트 워크북)을 이어서 실행한다.

## 구현 불변 규칙 (스크립트 수정 시 유지)

1. **식별자 필수**: KR corp_code·stock_code / US ticker·cik는 candidates.json에 확정 기입 — 스크립트가 기업명으로 추측 매칭하지 않는다.
2. **결측=분모 제외**: 점수 항목에 데이터가 없으면 0점 처리 금지(적자·데이터 부족 기업이 부당하게 깎이거나, 결측이 많은 기업이 부풀려지지 않게 정규화). 결측 목록은 요약 시트에 표시.
3. **PER/PBR은 순이익·자본 ≤0이면 계산하지 않는다**(음수 배수 오판 방지 — us/dart N섹션과 동일 규칙). 배수는 (E) 추정연도가 아니라 마지막 실적 연도 기준.
4. 한 후보의 실패가 전체 실행을 중단시키지 않는다. errors는 결과 JSON과 사용자 보고에 남긴다.
5. API 키는 인자로만 받고 출력·JSON·워크북에 남기지 않는다.
6. 요약 시트의 종합점수는 값 복사가 아니라 "점수_산식" 시트 셀 참조, 합계·정규화·그레이엄 판정은 시트 내 수식(감사 가능성 원칙).

## 테스트

```
cd plugins/industry-stock-picker/skills/industry-stock-picker && python -m unittest discover -s tests -v
```
합성 후보 3사(json 소스)로 점수·정규화·순위·워크북 생성을 검증한다.
