valuation-verdict
`dart-kospi-financials`(한국)와 `us-stocks-financials`(미국)가 만든 재무 캐시·워크북을 받아 **"지금 주가가 기업가치 대비 싼가 비싼가, 적정주가는 얼마고 어디까지 갈 수 있나"**를 한 장의 시트와 한 줄의 판정으로 끝낸다.
무엇이 나오나
항목	설명
적정주가(기준)	자산가치·수익가치·fcf dcf·잔여이익(rim)·배당(ddm)의 가중 평균
괴리율 / 상승여력	현재가 대비
판정	저평가(강) / 저평가 / 적정 / 고평가 / 고평가(강)
목표주가 범위	비관·기준·낙관 시나리오(할인율·성장률·목표 per 동시 변동)
하방 참고	청산가치/주, 비관 시나리오 값
기대 연평균수익률	n년 내 적정가 수렴 가정의 산술값(예측 아님)
신뢰도	방법 간 분산·데이터 완결성·추정연도·이익의 질·주가 경과일로 감점
불확실성 고지	기본값을 쓴 가정 목록("이 정보들은 정확하지 않습니다")
산출물: 워크북 맨 앞의 `가치평가_결론` 시트(가정 셀은 파란색 입력값, 나머지는 전부 수식 → 가정을 바꾸면 기준 시나리오가 재계산), `verdict_{회사}.json`, 표준출력 마크다운 요약.
사용
```bash
s=plugins/valuation-verdict/skills/valuation-verdict/x-scripts
# kr
python $s/valuation_verdict.py --source dart --cache-dir plugins/dart-kospi-financials/skills/dart-financial-extractor/cache \
    --corp-code 00126380 --xlsx "삼성전자_연간_20260831.xlsx" --assumptions my_assumptions.json
# us
python $s/valuation_verdict.py --source sec --cache-dir plugins/us-stocks-financials/skills/us-stock-financial-extractor/cache \
    --ticker aapl --xlsx "apple inc._20260831.xlsx"
# 워크북만 있을 때 (수식 값이 없으면 libreoffice로 자동 재계산)
python $s/valuation_verdict.py --source xlsx --xlsx "삼성전자_연간_20260831.xlsx"
```
가정 파일 예시: `skills/valuation-verdict/examples/assumptions.example.json`.
테스트
```bash
cd plugins/valuation-verdict/skills/valuation-verdict && python -m unittest discover -s tests -v
```
외부 api 호출 없이 합성 픽스처로 dart/sec/xlsx/json 소스 전부와 시트 수식(libreoffice 재계산) 대조를 검증한다.
설계 메모
지배주주 귀속 순이익·자본을 우선 사용(비지배지분이 큰 지주사·조선사에서 per/pbr 왜곡 방지).
적자·자본잠식·주가 없음 등은 해당 방법을 제외하거나 "판정 불가"로 내보낸다. 0으로 채우지 않는다.
dart 진행연도(e) 추정에서 전년 동기가 0 이하거나 부호가 바뀌는 계정은 추정하지 않는다(턴어라운드 기업 오류 방지).
한계: 업종 상대가치 미포함(목표 per로만 반영), ev→equity 브릿지 없음(fcf를 주주 현금흐름으로 근사), 자기주식 미차감.
