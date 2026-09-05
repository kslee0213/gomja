#!/usr/bin/env python3
"""fetch_financials.py — SEC EDGAR의 공식 XBRL CompanyFacts API로 재무제표를 수집한다.

v3.0.0: yfinance(비공식 스크레이핑, Yahoo의 403 차단 위험)를 버리고 SEC가
직접 제공하는 공식·무료·키불필요 API로 전환했다. DART Open API와 원리가
같다(XBRL 표준계정 태그 기반) — 다만 SEC는 태그가 us-gaap:Revenues처럼
영문 표준 태그이고, DART는 한국 XBRL 계정ID를 쓴다는 차이만 있다.

엔드포인트: https://data.sec.gov/api/xbrl/companyfacts/CIK{10자리}.json
- API 키 불필요. 다만 User-Agent에 "앱이름 연락처이메일" 형식이 반드시 있어야 한다
  (SEC 정책. 없으면 403이 날 수 있다).
- 레이트리밋: 초당 10회. 이 스크립트는 기업 하나당 요청 1회면 끝나므로 문제없다.
- 응답에는 그 회사가 XBRL로 공시한 "모든" 계정의 "모든" 과거 값이 들어있다
  (10-K/10-Q 전체 이력). 우리가 원하는 계정·기간만 골라 쓰면 된다.

⚠️ duration(손익계산서/현금흐름표) 계정은 같은 개념에 대해 "3개월(분기단독)",
"6개월(반기누적)", "9개월(3분기누적)", "12개월(연간)" 값이 전부 한 리스트에
섞여 들어있다. (end-start) 일수로 걸러서 원하는 기간 길이만 골라야 한다 —
이건 DART의 "1분기/반기/3분기/사업보고서를 각각 별도 API로 조회"하던 방식과
다른 부분이다.
"""
import argparse
import json
import sys
import time
from datetime import date, datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"
COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"

# 표준 us-gaap 태그. 회사마다 쓰는 태그가 조금씩 다를 수 있어(관행 차이) 후보를
# 여러 개 둔다 — 앞에서부터 순서대로 시도해 처음 값이 있는 것을 쓴다.
ACCOUNTS = {
    # key: (구분, [태그 후보들], 라벨)
    "매출액": ("income_statement", ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"], "매출액"),
    "매출원가": ("income_statement", ["CostOfRevenue", "CostOfGoodsAndServicesSold"], "매출원가"),
    "매출총이익": ("income_statement", ["GrossProfit"], "매출총이익"),
    "영업이익": ("income_statement", ["OperatingIncomeLoss"], "영업이익"),
    "세전이익": ("income_statement", [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
    ], "세전이익"),
    "법인세비용": ("income_statement", ["IncomeTaxExpenseBenefit"], "법인세비용"),
    "당기순이익": ("income_statement", ["NetIncomeLoss", "ProfitLoss"], "당기순이익"),
    "이자비용": ("income_statement", ["InterestExpense"], "이자비용"),
    "EPS(희석)": ("income_statement", ["EarningsPerShareDiluted"], "EPS(희석)"),

    "유동자산": ("balance_sheet", ["AssetsCurrent"], "유동자산"),
    "자산총계": ("balance_sheet", ["Assets"], "자산총계"),
    "유동부채": ("balance_sheet", ["LiabilitiesCurrent"], "유동부채"),
    "부채총계": ("balance_sheet", ["Liabilities"], "부채총계"),
    "자본총계": ("balance_sheet", ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"], "자본총계"),
    "현금및현금성자산": ("balance_sheet", ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"], "현금및현금성자산"),
    "재고자산": ("balance_sheet", ["InventoryNet"], "재고자산"),
    "매출채권": ("balance_sheet", ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"], "매출채권"),
    "유형자산": ("balance_sheet", ["PropertyPlantAndEquipmentNet"], "유형자산(PP&E)"),

    "매입채무": ("balance_sheet", ["AccountsPayableCurrent", "AccountsPayableAndAccruedLiabilitiesCurrent", "AccountsPayableTradeCurrent"], "매입채무"),
    "단기투자": ("balance_sheet", ["ShortTermInvestments", "MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent"], "단기투자(유가증권)"),
    "장기투자": ("balance_sheet", ["LongTermInvestments", "MarketableSecuritiesNoncurrent", "AvailableForSaleSecuritiesDebtSecuritiesNoncurrent"], "장기투자자산"),
    "무형자산": ("balance_sheet", ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet"], "무형자산"),
    "영업권": ("balance_sheet", ["Goodwill"], "영업권"),
    "기타유동자산": ("balance_sheet", ["OtherAssetsCurrent"], "기타유동자산"),
    "이익잉여금": ("balance_sheet", ["RetainedEarningsAccumulatedDeficit"], "이익잉여금"),

    "영업활동현금흐름": ("cash_flow", ["NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"], "영업활동현금흐름"),
    "설비투자": ("cash_flow", ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"], "유형자산의취득"),
    "투자활동현금흐름": ("cash_flow", ["NetCashProvidedByUsedInInvestingActivities", "NetCashProvidedByUsedInInvestingActivitiesContinuingOperations"], "투자활동현금흐름"),
    "재무활동현금흐름": ("cash_flow", ["NetCashProvidedByUsedInFinancingActivities", "NetCashProvidedByUsedInFinancingActivitiesContinuingOperations"], "재무활동현금흐름"),
    "감가상각비": ("cash_flow", ["DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet", "DepreciationAndAmortization", "Depreciation"], "감가상각비(D&A)"),
    "배당금지급": ("cash_flow", ["PaymentsOfDividends", "PaymentsOfDividendsCommonStock"], "배당금지급"),
    "현금및현금성자산의증가": ("cash_flow", [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
        "CashAndCashEquivalentsPeriodIncreaseDecrease",
        "CashAndCashEquivalentsPeriodIncreaseDecreaseExcludingExchangeRateEffect",
    ], "현금및현금성자산의증가"),
    "외화환산손익": ("income_statement", ["ForeignCurrencyTransactionGainLossBeforeTax"], "외화환산손익"),
}

# 주당(per-share) 계정: 백만달러 환산 금지 + 분기 파생(Q4=연간−3개 분기 합) 대상 제외
PER_SHARE_KEYS = {"EPS(희석)"}
# 회사에 따라 아예 없는 게 정상인 계정 — 매칭 실패 경고에서 제외
OPTIONAL_KEYS = {"외화환산손익", "단기투자", "장기투자", "무형자산", "영업권", "기타유동자산", "매입채무", "이자비용", "현금및현금성자산의증가"}

INSTANT_KEYS = {k for k, (sj, *_r) in ACCOUNTS.items() if sj == "balance_sheet"}


def _explain_error(e: Exception) -> str:
    msg = str(e)
    if "403" in msg:
        return (
            f"{msg} — SEC가 요청을 차단했습니다. (1) User-Agent에 연락처 이메일이 올바르게 "
            "들어갔는지, (2) data.sec.gov가 Cowork 네트워크 허용 목록에 있는지 확인하세요."
        )
    if "429" in msg:
        return f"{msg} — SEC 레이트리밋(초당 10회). 요청 간격을 늘려 재시도하세요."
    return msg


def _headers(contact_email: str) -> dict:
    if not contact_email or "@" not in contact_email:
        raise ValueError("SEC EDGAR는 User-Agent에 유효한 연락처 이메일이 필요합니다(--contact).")
    return {"User-Agent": f"us-stocks-financials-plugin {contact_email}"}


def fetch_companyfacts(cik: str, contact_email: str) -> dict | None:
    """SEC CompanyFacts 원본 JSON을 그대로 받아온다(가공하지 않음 — 가공은
    load 시점에 한다. 원본을 그대로 캐시해야 나중에 다른 계정을 추가로
    파싱하고 싶을 때 재조회할 필요가 없다)."""
    import requests

    url = COMPANYFACTS_URL.format(cik=cik)
    try:
        resp = requests.get(url, headers=_headers(contact_email), timeout=30)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"FAIL SEC CompanyFacts 조회 실패(CIK {cik}): {_explain_error(e)}", file=sys.stderr)
        return None


def save_companyfacts_cache(cik: str, ticker: str, data: dict, cache_dir: Path = CACHE_DIR) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"secfacts_{ticker.upper()}.json"
    payload = {"ticker": ticker.upper(), "cik": cik, "fetched_at": datetime.now().isoformat(), "raw": data}
    cache_file.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(f"SAVE 캐시 저장: {cache_file}")
    return cache_file


def load_companyfacts_cache(ticker: str, cache_dir: Path = CACHE_DIR) -> dict | None:
    cache_file = cache_dir / f"secfacts_{ticker.upper()}.json"
    if not cache_file.exists():
        return None
    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except Exception:
        return None


def _duration_days(fact: dict) -> int | None:
    if "start" not in fact or "end" not in fact:
        return None
    try:
        d1 = date.fromisoformat(fact["start"])
        d2 = date.fromisoformat(fact["end"])
        return (d2 - d1).days
    except Exception:
        return None


def _unit_facts(concept: dict) -> list[dict]:
    """USD 단위를 우선 쓰되, EPS처럼 USD/shares(또는 shares) 단위만 있는 계정도 지원한다."""
    units = concept.get("units", {})
    for u in ("USD", "USD/shares", "shares"):
        if units.get(u):
            return units[u]
    return []


def _best_points(facts: list[dict], want_instant: bool) -> tuple[dict[str, float], dict[str, float]]:
    """(quarterly {end: val}, annual {end: val}) — 같은 end에 여러 공시(원공시·비교표·정정)가
    있으면 filed가 가장 최신인 값을 쓴다(정정공시 반영)."""
    q: dict[str, tuple[str, float]] = {}
    a: dict[str, tuple[str, float]] = {}
    for fact in facts:
        val, end, filed = fact.get("val"), fact.get("end"), fact.get("filed", "")
        if val is None or end is None:
            continue
        if want_instant:
            if "start" in fact:
                continue
            # instant는 분기/연간 구분이 없으므로 임시로 양쪽에 후보를 둔다(호출부에서 날짜 매칭)
            if end not in a or filed >= a[end][0]:
                a[end] = (filed, float(val))
            continue
        dur = _duration_days(fact)
        if dur is None:
            continue
        bucket = q if 80 <= dur <= 100 else a if 350 <= dur <= 380 else None
        if bucket is None:
            continue
        if end not in bucket or filed >= bucket[end][0]:
            bucket[end] = (filed, float(val))
    return {k: v for k, (_, v) in q.items()}, {k: v for k, (_, v) in a.items()}


def _collect_duration_dates(raw_facts: dict) -> tuple[set[str], set[str]]:
    """IS/CF(duration) 계정들에서 실제 존재하는 "연간(FY, 350~380일)"과
    "분기(80~100일)" 날짜(end) 집합을 구한다. BS(instant) 항목은 이 날짜
    집합에 매칭되는 것만 annual/quarterly로 배정한다 — 그렇지 않으면 아직
    FY 10-K가 없는 진행 중인 분기말 BS 스냅샷이 "연간" 쪽에 잘못 섞여
    들어가는 문제가 생긴다(실제로 재현·확인된 버그)."""
    annual_dates: set[str] = set()
    quarterly_dates: set[str] = set()
    us_gaap = raw_facts.get("facts", {}).get("us-gaap", {})
    duration_keys = [k for k in ACCOUNTS if k not in INSTANT_KEYS]
    for key in duration_keys:
        _sj, candidates, _label = ACCOUNTS[key]
        for tag in candidates:
            concept = us_gaap.get(tag)
            if not concept:
                continue
            for fact in _unit_facts(concept):
                end = fact.get("end")
                dur = _duration_days(fact)
                if end is None or dur is None:
                    continue
                if 80 <= dur <= 100:
                    quarterly_dates.add(end)
                elif 350 <= dur <= 380:
                    annual_dates.add(end)
    return annual_dates, quarterly_dates


def extract_periods(raw_facts: dict, key: str, annual_dates: set[str], quarterly_dates: set[str]) -> dict[str, float]:
    """ACCOUNTS[key]의 태그 후보들을 raw_facts(companyfacts의 facts.us-gaap)에서
    찾아 {end_date: value} 형태로 반환한다. duration 계정은 (end-start) 일수로
    분기/연간을 스스로 구분하고, instant(BS) 계정은 duration 계정들에서 이미
    확인된 annual_dates/quarterly_dates에 매칭되는 날짜만 채택한다(위 설명 참고)."""
    sj, candidates, _label = ACCOUNTS[key]
    is_instant = key in INSTANT_KEYS
    us_gaap = raw_facts.get("facts", {}).get("us-gaap", {})

    out_q: dict[str, float] = {}
    out_a: dict[str, float] = {}
    # 태그 후보를 순서대로 훑되, 앞 태그가 채우지 못한 날짜만 뒤 태그로 보충한다.
    # (예: Revenues가 구연도에만 있고 최근 연도는 RevenueFromContract…에만 있는 회사 —
    # 첫 매칭 태그에서 끝내면 최근 연도 매출이 통째로 비는 실제 문제의 방지책)
    for tag in candidates:
        concept = us_gaap.get(tag)
        if not concept:
            continue
        facts = _unit_facts(concept)
        if not facts:
            continue
        t_q, t_a = _best_points(facts, is_instant)
        if is_instant:
            # instant는 duration 계정들의 실제 연간/분기 날짜에 매칭되는 것만 채택
            for end, v in t_a.items():
                if end in annual_dates:
                    out_a.setdefault(end, v)
                if end in quarterly_dates:
                    out_q.setdefault(end, v)
        else:
            for end, v in t_q.items():
                out_q.setdefault(end, v)
            for end, v in t_a.items():
                out_a.setdefault(end, v)
    return {"quarterly": out_q, "annual": out_a}


def _derive_q4(q: dict[str, float], a: dict[str, float]) -> dict[str, float]:
    """10-K에는 4분기 단독(3개월) 값이 보통 없다. 같은 회계연도의 분기 3개(Q1~Q3)가
    있으면 Q4 = 연간 − (Q1+Q2+Q3)로 파생한다. 3개가 안 모이면 지어내지 않고 건너뛴다."""
    derived = {}
    for end_a, val_a in a.items():
        if end_a in q:
            continue
        try:
            d_a = date.fromisoformat(end_a)
        except ValueError:
            continue
        in_fy = [v for e, v in q.items()
                 if 0 < (d_a - date.fromisoformat(e)).days <= 340]
        if len(in_fy) == 3:
            derived[end_a] = val_a - sum(in_fy)
    return derived


def build_frequency_payload(raw_facts: dict, frequency: str, n_periods: int) -> dict:
    """extract_periods 결과를 build_workbook.py가 기대하는
    {sj: {key: {period: val}}} 형태로 재구성하고, 최근 n_periods개만 남긴다.
    frequency="quarterly"면 Q4를 연간−(Q1~Q3)로 파생하고(주당 지표 제외),
    그 회계연도말 BS 스냅샷도 분기 축에 포함시킨다."""
    annual_dates, quarterly_dates = _collect_duration_dates(raw_facts)
    if frequency == "quarterly":
        # Q4 파생이 가능한 회계연도말 날짜를 분기 날짜 집합에 추가(BS 매칭용)
        quarterly_dates = set(quarterly_dates) | set(annual_dates)
    result = {"income_statement": {}, "balance_sheet": {}, "cash_flow": {}}
    all_dates: set[str] = set()
    per_key: dict[str, dict[str, float]] = {}
    for key, (sj, _cand, _label) in ACCOUNTS.items():
        periods = extract_periods(raw_facts, key, annual_dates, quarterly_dates)
        series = dict(periods[frequency])
        if frequency == "quarterly" and key not in INSTANT_KEYS and key not in PER_SHARE_KEYS:
            series.update(_derive_q4(periods["quarterly"], periods["annual"]))
        per_key[key] = series
        if key not in INSTANT_KEYS:
            all_dates.update(series.keys())
    # 기간 축은 유량(duration) 계정이 실제로 존재하는 날짜만 사용(BS 스냅샷 단독 날짜 방지)
    sorted_dates = sorted(all_dates)[-n_periods:] if n_periods else sorted(all_dates)
    for key, (sj, _cand, _label) in ACCOUNTS.items():
        series = per_key.get(key, {})
        row = {d: series[d] for d in sorted_dates if d in series}
        if row:
            result[sj][key] = row
    return result, sorted_dates


def main() -> None:
    ap = argparse.ArgumentParser(description="SEC EDGAR 재무제표 수집")
    ap.add_argument("ticker")
    ap.add_argument("cik", help="10자리 SEC CIK(ticker_lookup.py로 확인)")
    ap.add_argument("--contact", required=True, help="SEC User-Agent에 넣을 연락처 이메일(필수)")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    cached = None if args.force else load_companyfacts_cache(args.ticker)
    if cached:
        print(f"CACHE SEC CompanyFacts 캐시 재사용: {args.ticker}")
    else:
        data = fetch_companyfacts(args.cik, args.contact)
        if not data:
            print(json.dumps({"saved": False, "ticker": args.ticker}, ensure_ascii=False))
            sys.exit(1)
        save_companyfacts_cache(args.cik, args.ticker, data)
        time.sleep(0.15)  # SEC 레이트리밋(초당 10회) 여유

    print(json.dumps({"saved": True, "ticker": args.ticker, "cik": args.cik}, ensure_ascii=False))


if __name__ == "__main__":
    main()
