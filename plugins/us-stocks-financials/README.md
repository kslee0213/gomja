# us-stocks-financials

미국 나스닥/뉴욕증권거래소 상장 기업의 재무제표를 **SEC EDGAR 공식 API**로 수집해 엑셀 분석 파일을 생성합니다. `dart-kospi-financials`(한국 DART 버전)와 같은 설계 원칙(원본데이터 시트 + 수식 기반 감사 가능성)을 따릅니다.

> v3.0.0에서 데이터 소스를 yfinance(비공식 스크레이핑)에서 SEC EDGAR(공식 정부 API)로 전환했습니다. 자세한 내용은 `SKILL.md` 참고.

## 특징

- **SEC EDGAR 기반**: 재무제표는 미국 SEC가 직접 제공하는 무료 공식 API. API 키 불필요(연락처 이메일만 필요, SEC 정책). DART Open API와 같은 XBRL 표준계정 원리라 계정 매칭이 견고합니다.
- **주가만 최소한 yfinance 사용**: SEC엔 시장 주가가 없어서, 현재가·시가총액만 yfinance로 가져옵니다. 이 부분이 실패해도 재무제표(핵심 산출물)는 영향받지 않습니다.
- **원본데이터 시트 + 수식 참조**: 모든 셀이 감사 가능한 수식으로 구성됩니다. PER/PBR/PSR도 완제품 값을 그대로 쓰지 않고 우리가 직접 계산합니다.
- **캐싱**: 재실행 시 캐시를 실제로 읽어서 재사용합니다.

## 필수 설정

```bash
pip install requests openpyxl yfinance --break-system-packages
```

- SEC EDGAR 호출 시 User-Agent에 연락처 이메일 필요(인증키 아님).
- `data.sec.gov`가 Cowork 네트워크 허용 목록에 있는지 확인 필요.

## 사용 예시

```
애플 재무제표 뽑아줘
테슬라 년 분석해줘
마이크로소프트, 구글 비교해줘
```

## 구성

| 스크립트 | 역할 |
|---|---|
| `ticker_lookup.py` | 기업명 → Ticker → SEC CIK 변환(SEC 공식 매핑 활용) |
| `fetch_financials.py` | SEC EDGAR CompanyFacts API로 재무제표 수집·캐시 |
| `fetch_extra_info.py` | 현재 주가·시가총액만 yfinance로 최소 수집·캐시 |
| `build_workbook.py` | 단일 기업 엑셀 생성 |
| `build_comparison_workbook.py` | 여러 기업 비교 엑셀 생성 |

## 생성되는 엑셀 파일

- **분기_재무제표 / 연간_재무제표**: SEC 공시 이력에 있는 만큼(회사가 XBRL 공시를 시작한 이후 전체 이력 접근 가능 — yfinance의 4~5개 제약보다 넉넉함)
- **지표_분기 / 지표_연간**: 재무비율 11개 + 라인차트 4개
- **투자분석**: 재무비율 요약, 위험신호, 주가 연동 지표(PER/PBR/PSR 직접 계산), 간이 투자판단
- **원본데이터**: raw 값(숨김), 다른 모든 시트가 참조하는 소스

## 문서

- `SKILL.md`: 스킬 정의, 실행 절차, 버전 변경 이력
- `QUICKSTART.md`: 빠른 시작 가이드
- `HARNESS.md`: 파이프라인 아키텍처
- `skills/us-stock-financial-extractor/references/sec_edgar_xbrl_reference.md`: XBRL 태그 참고

## 라이선스

MIT License
