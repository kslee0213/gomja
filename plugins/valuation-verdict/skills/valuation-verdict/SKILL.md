---
name: valuation-verdict
description: >
  Use this skill when the user wants a *conclusion* about a stock's valuation —
  whether the current price is over- or under-valued versus intrinsic value, what
  the fair price is, and how far the price could go (target range). Triggers on
  requests like "OO 적정주가 얼마야", "OO 지금 사도 돼? 고평가야 저평가야",
  "OO 목표주가/상승여력 계산해줘", "OO 가치평가 결론 내줘", "OO 얼마까지 오를 수 있어",
  "AAPL fair value", or after the dart-kospi-financials / us-stocks-financials
  workbooks are built and the user asks "그래서 결론이 뭐야". Works for both Korean
  (DART/KRX) and US (SEC/yfinance) companies from the caches or workbooks those
  plugins already produced; adds a "가치평가_결론" sheet plus a JSON/markdown verdict.
metadata:
  version: "1.0.0"
---

# 가치평가 결론 (valuation-verdict)

기존 두 플러그인은 재무제표·지표·안전마진·DCF를 **나열**만 하고, "그래서 지금 주가가 싼가 비싼가, 적정주가는 얼마고 어디까지 갈 수 있나"를 **한 줄로 말해주지 않는다.** 이 스킬은 그 마지막 단계다. 여러 방법을 같은 잣대(주당 가치)로 계산하고 가중 평균해 적정주가·괴리율·목표주가 범위·판정·신뢰도를 낸다.

> 절대 규칙: 결론은 **가정에 종속된 조건부 추정**이다. 기본값을 쓴 가정은 결과물에 그대로 열거하고, 그 부분에 대해 "이 정보들은 정확하지 않습니다"라고 명시한다. 주가 "예측"이라는 표현을 쓰지 않는다.

## 0. 입력 확인

| 소스 | 필요한 것 | 비고 |
|---|---|---|
| `--source dart` | `dart-financial-extractor`의 `cache/` 폴더 + `corp_code` | 사업보고서 5개년(+진행연도 분기/반기), `price_*.json`(KRX), `extra_*.json`(배당) |
| `--source sec` | `us-stock-financial-extractor`의 `cache/` 폴더 + 티커 | `secfacts_{티커}.json`, `price_{티커}.json` |
| `--source xlsx` | 두 플러그인이 만든 연간 워크북 | 수식 값이 비어 있으면 LibreOffice로 자동 재계산해 읽음 |
| `--source json` | 정규화 JSON(`examples/normalized_input.example.json`) | 다른 파이프라인 연동용 |

**주가가 없으면 판정을 낼 수 없다.** KR은 `fetch_stock_price.py`, US는 `fetch_extra_info.py`로 현재가를 먼저 받는다. KR에서 진행연도(E)가 없어 결산일(12/31) 주가만 있는 경우, 가장 최근 영업일 주가를 한 번 더 조회한다(오래된 주가는 신뢰도 감점).

## 1. 실행

```bash
# KR — 기존 연간 워크북에 시트 추가
python scripts/valuation_verdict.py --source dart --cache-dir <dart cache> --corp-code 00126380 \
    --xlsx "/mnt/user-data/outputs/삼성전자_연간_20260831.xlsx" [--assumptions a.json]

# US
python scripts/valuation_verdict.py --source sec --cache-dir <sec cache> --ticker AAPL \
    --xlsx "/mnt/user-data/outputs/Apple Inc._20260831.xlsx"

# 워크북만 있을 때
python scripts/valuation_verdict.py --source xlsx --xlsx <워크북>
```

- `--xlsx`를 주면 그 파일에 **"가치평가_결론" 시트를 맨 앞에 추가**한다(있으면 교체). 없으면 `--out-xlsx`로 새 파일.
- 표준출력의 마크다운 요약을 그대로 사용자 답변의 뼈대로 쓴다. `verdict_{회사}.json`에 모든 수치·가정·경고가 들어 있다.
- 시트를 쓴 뒤에는 `/mnt/skills/public/xlsx/scripts/recalc.py <파일>`을 한 번 실행해 수식 오류가 0인지 확인한다(openpyxl 저장 시 캐시된 계산값이 사라지므로).

## 2. 가정(assumptions) — Claude가 반드시 판단해서 채울 것

스크립트는 가정 없이도 돌아가지만, **그 경우 시장 기본값**(r 9%, g 2~2.5%, 목표 PER KR 10배/US 18배)을 쓰고 이를 "정확하지 않음"으로 표시한다. 사용자에게 의미 있는 결론을 주려면 아래를 리서치해 `--assumptions` JSON으로 넘긴다(`examples/assumptions.example.json`):

| 키 | 의미 | 어떻게 정하나 |
|---|---|---|
| `rf_pct`, `beta`, `erp_pct` | r = rf + β×ERP | 국고채/미국채 10년 금리 + 시장 ERP(KR 5~7%, US 4~6%) + 종목 베타(웹 검색). 셋 다 있으면 `cost_of_equity_pct`보다 우선 |
| `cost_of_equity_pct` | 직접 지정 | 베타를 못 찾으면 업종 관행(방어주 7~8%, 경기민감/성장주 10~12%) |
| `terminal_growth_pct` | 영구성장률 | 장기 명목 GDP 이하. KR 1.5~2.5, US 2~3 |
| `stage1_growth_pct` | 향후 5년 FCF 성장률 | 투자판단 종합의 매출 성장 가정, 애널리스트 컨센서스와 정합되게 |
| `target_per` | 목표 PER | 동종업계 평균/회사 과거 밴드/성장률 감안. 비워두면 과거 5년 PER 중앙값(5~40배 범위일 때) |
| `weights` | 방법별 가중치 | 자산주(지주·은행·건설)는 asset/rim ↑, 성장주는 dcf/earnings ↑, 배당주는 ddm에 소폭 |

**가정을 채운 근거를 답변에 반드시 쓴다** — "r=8.5%: 국고채 3.2% + β1.0×ERP 5.3%(출처 …)". 근거 없는 숫자는 기본값보다 나쁘다.

## 3. 결과 해석 → 사용자에게 말하는 법

1. **한 줄 결론 먼저**: "현재가 65,000원 vs 적정주가 13,600원 → 상승여력 −79%, **고평가(강)**, 신뢰도 낮음."
2. **왜 그런지 방법별로**: 자산가치/수익가치/DCF/RIM 각각의 값과 가중치. 값이 크게 갈리면 그 이유(예: "DCF는 FCF 성장 가정에 민감, 영구가치 비중 80%").
3. **목표주가 범위**: 비관/기준/낙관 3개 값과 각각의 가정 차이. "얼마까지 오르나"는 이 낙관 값이 상단이며 **시점은 말할 수 없다**고 분명히 한다. 기대 연평균수익률은 "N년 안에 적정가로 수렴한다는 가정"의 산술값임을 덧붙인다.
4. **하방**: 청산가치/주, 비관 시나리오 값.
5. **신뢰도와 감점 사유** (방법 간 분산, 추정연도(E), 이익의 질, 오래된 주가, 적자 등).
6. **기본값을 쓴 가정 목록** + "이 정보들은 정확하지 않습니다".
7. 기존 시트(투자분석 N섹션 등급, 버핏멍거 안전마진)와 **모순되면 그 사실을 말한다** — 예: "버핏멍거 시트의 자산기반 내재가치(BPS×10)는 산식 자체가 과대하므로 이 결론에서는 쓰지 않았다".

## 4. 판정 기준(고정)

| 상승여력 = 적정주가/현재가 − 1 | 판정 |
|---|---|
| ≥ +30% | 저평가(강) |
| +10 ~ +30% | 저평가 |
| -10 ~ +10% | 적정 |
| -30 ~ -10% | 고평가 |
| ≤ -30% | 고평가(강) |

신뢰도: 높음(3) → 방법 간 분산 ≥40% −1, ≥80% −2 / 적용 방법 ≤2개 −1 / 추정연도(E) −1 / OCF÷순이익 3년 평균 <0.8 −1 / 주가 60일 이상 경과 −1 / 최근 적자 −1 → 보통(2), 낮음(≤1).

## 5. 계산 방법 요약 (시트 3번 섹션 = 전부 수식)

- **A 자산가치** = 지배주주 BPS × 정당 PBR, 정당 PBR = (ROE_norm − g)/(r − g)를 [0.3, 4.0]으로 클램프. ROE_norm은 최근 3년 평균.
- **B 수익가치** = 정상화 EPS × 목표 PER. 정상화 EPS = 최근 3년 중 양수 평균(최근 연도가 흑자일 때만).
- **C 현금흐름가치** = FCF(OCF − CAPEX) 3년 평균을 5년 g1 성장 → 5년간 g로 선형 페이드 → 영구가치. 자기자본비용 r로 할인(FCF는 이자 지급 후 값이므로 주주 현금흐름 근사). 영구가치 비중 >80%면 DCF 가중치 40% 축소.
- **D 잔여이익(RIM)** = BPS + Σ(ROE_t − r)·B_{t−1}/(1+r)^t, ROE가 10년에 걸쳐 r로 수렴(이후 초과이익 0).
- **E 배당(DDM)** = DPS(1+g)/(r−g), 참고용(기본 가중치 0).
- 적자 기업: B·C 제외, A 50%·C(양수면) 30%·D 20%. PBR<0.7 & ROE<6%: 자산가치 비중 확대.
- 시나리오: 비관 r+1.5%p·g1×0.5·PER×0.8 / 낙관 r−1.0%p·g1×1.3(상한 20%)·PER×1.2.

## 6. 이 스킬이 하지 않는 것

- 업종 상대가치(동종업계 PER/EV/EBITDA 비교): 목표 PER를 사용자가 넣는 방식으로만 반영.
- 이자부부채 기반 EV→Equity 브릿지: FCF를 주주 현금흐름으로 근사(순현금 가산 없음, 보수적). 순부채가 큰 회사는 `weights`에서 dcf를 낮추거나 결론 텍스트에 한계를 적는다.
- 자기주식 차감 주식수: 발행주식수 기준(자기주식 비중이 크면 주당 값이 소폭 과소).
- 시점 예측: "언제" 오르는지는 계산 대상이 아니다.

## 7. 테스트

```bash
cd plugins/valuation-verdict/skills/valuation-verdict && python -m unittest discover -s tests -v
```
합성 픽스처로 dart/sec/xlsx/json 네 소스와 LibreOffice 수식 재계산 대조, 적자·주가없음·사용자 가정 재지정 케이스를 검증한다.
<img width="451" height="674" alt="image" src="https://github.com/user-attachments/assets/173feb23-e694-41df-a8ea-b3cdac566f27" />
