"""합성(synthetic) 픽스처 생성기.

실제 dart/sec api를 호출하지 않고, 두 추출 플러그인의 캐시 파일 형식과 동일한
json을 만들어 `valuation_verdict.py`의 어댑터·계산·시트 생성을 검증한다.

    python tests/make_fixtures.py --dart-cache <dart cache dir> --sec-cache <sec cache dir>

가상 회사:
  - kr "가상전자"(corp_code 00000001, 종목코드 000001): 2021~2025 사업보고서 + 2025/2026 반기보고서
  - us "fake"(cik 0000000001): fy2021~fy2025 10-k + 분기 10-q
숫자는 전부 임의값이며 실제 기업과 무관하다.
"""
import argparse
import datetime as dt
import json
from pathlib import path


def won(v: float) -> str:
    return f"{int(round(v)):,}"


# ---------------------------------------------------------------- dart ----
def dart_items(year: int, scale: float, half: bool = false) -> dict:
    """연도별 합성 재무제표. half=true면 반기 누적(약 48%)으로 만든다."""
    rev = 1_000_000_000_000 * scale          # 1조 × scale
    f = 0.48 if half else 1.0
    cogs = rev * 0.70
    gp = rev - cogs
    op = rev * 0.10
    fin_inc, fin_exp, oth_inc, oth_exp = rev * 0.005, rev * 0.008, rev * 0.002, rev * 0.003
    pbt = op + fin_inc - fin_exp + oth_inc - oth_exp
    tax = pbt * 0.22
    ni = pbt - tax
    ni_parent = ni * 0.92
    ni_nci = ni - ni_parent
    da = rev * 0.04
    ocf = ni + da + rev * 0.01
    capex = rev * 0.05
    div = rev * 0.012

    assets = rev * 1.2
    cash = rev * 0.10
    stfin = rev * 0.05
    ar = rev * 0.12
    inv = rev * 0.08
    oca = rev * 0.02
    ca = cash + stfin + ar + inv + oca
    ppe = rev * 0.45
    intang = rev * 0.05
    ltinv = rev * 0.06
    nca = assets - ca
    liab = assets * 0.45
    ap = rev * 0.09
    stdebt = rev * 0.08
    cl = ap + stdebt + rev * 0.05
    ltdebt = rev * 0.12
    ncl = liab - cl
    equity = assets - liab
    eq_parent = equity * 0.93
    nci = equity - eq_parent
    re_ = equity * 0.6
    shares = 100_000_000

    def row(sj, aid, nm, amt, add=none, ord_=0):
        d = {"sj_div": sj, "account_id": aid, "account_nm": nm, "ord": str(ord_),
             "thstrm_amount": won(amt)}
        if add is not none:
            d["thstrm_add_amount"] = won(add)
        return d

    bs = [
        row("bs", "ifrs-full_currentassets", "유동자산", ca, ord_=1),
        row("bs", "ifrs-full_cashandcashequivalents", "현금및현금성자산", cash, ord_=2),
        row("bs", "dart_shorttermdepositsnotclassifiedascashequivalents", "단기금융상품", stfin, ord_=3),
        row("bs", "ifrs-full_tradeandothercurrentreceivables", "매출채권및기타채권", ar, ord_=4),
        row("bs", "ifrs-full_inventories", "재고자산", inv, ord_=5),
        row("bs", "dart_othercurrentassets", "기타유동자산", oca, ord_=6),
        row("bs", "ifrs-full_noncurrentassets", "비유동자산", nca, ord_=7),
        row("bs", "ifrs-full_x-propertyplantandequipment", "유형자산", ppe, ord_=8),
        row("bs", "ifrs-full_intangibleassetsotherthangoodwill", "무형자산", intang, ord_=9),
        row("bs", "dart_longterminvestmentassets", "장기투자자산", ltinv, ord_=10),
        row("bs", "ifrs-full_assets", "자산총계", assets, ord_=11),
        row("bs", "ifrs-full_currentliabilities", "유동부채", cl, ord_=12),
        row("bs", "ifrs-full_tradeandothercurrentpayables", "매입채무및기타채무", ap, ord_=13),
        row("bs", "ifrs-full_shorttermborrowings", "단기차입금", stdebt, ord_=14),
        row("bs", "ifrs-full_noncurrentliabilities", "비유동부채", ncl, ord_=15),
        row("bs", "ifrs-full_longtermborrowings", "장기차입금", ltdebt, ord_=16),
        row("bs", "ifrs-full_liabilities", "부채총계", liab, ord_=17),
        row("bs", "ifrs-full_equityattributabletoownersofparent", "지배기업의 소유주에게 귀속되는 자본", eq_parent, ord_=18),
        row("bs", "ifrs-full_noncontrollinginterests", "비지배지분", nci, ord_=19),
        row("bs", "ifrs-full_equity", "자본총계", equity, ord_=20),
        row("bs", "ifrs-full_retainedearnings", "이익잉여금", re_, ord_=21),
    ]
    is_ = [
        row("is", "ifrs-full_revenue", "매출액", rev * f, ord_=1),
        row("is", "ifrs-full_costofsales", "매출원가", cogs * f, ord_=2),
        row("is", "ifrs-full_grossprofit", "매출총이익", gp * f, ord_=3),
        row("is", "dart_operatingincomeloss", "영업이익", op * f, ord_=4),
        row("is", "ifrs-full_financeincome", "금융수익", fin_inc * f, ord_=5),
        row("is", "ifrs-full_financecosts", "금융비용", fin_exp * f, ord_=6),
        row("is", "dart_othergains", "기타수익", oth_inc * f, ord_=7),
        row("is", "dart_otherlosses", "기타비용", oth_exp * f, ord_=8),
        row("is", "ifrs-full_interestexpense", "이자비용", fin_exp * 0.8 * f, ord_=9),
        row("is", "ifrs-full_profitlossbeforetax", "법인세비용차감전순이익", pbt * f, ord_=10),
        row("is", "ifrs-full_incometaxexpensecontinuingoperations", "법인세비용", tax * f, ord_=11),
        row("is", "ifrs-full_profitloss", "당기순이익" if not half else "반기순이익", ni * f, ord_=12),
        row("is", "ifrs-full_profitlossattributabletoownersofparent",
            "지배기업의 소유주에게 귀속되는 당기순이익" if not half else "지배기업의 소유주에게 귀속되는 반기순이익", ni_parent * f, ord_=13),
        row("is", "ifrs-full_profitlossattributabletononcontrollinginterests", "비지배지분에 귀속되는 당기순이익", ni_nci * f, ord_=14),
    ]
    cf = [
        row("cf", "ifrs-full_cashflowsfromusedinoperatingactivities", "영업활동현금흐름", ocf * f, ord_=1),
        row("cf", "dart_depreciationexpense", "감가상각비", da * f, ord_=2),
        row("cf", "ifrs-full_cashflowsfromusedininvestingactivities", "투자활동현금흐름", -(capex + rev * 0.01) * f, ord_=3),
        row("cf", "ifrs-full_purchaseofx-propertyplantandequipmentclassifiedasinvestingactivities", "유형자산의취득", -capex * f, ord_=4),
        row("cf", "ifrs-full_cashflowsfromusedinfinancingactivities", "재무활동현금흐름", -(div + rev * 0.005) * f, ord_=5),
        row("cf", "ifrs-full_dividendspaidclassifiedasfinancingactivities", "배당금지급", -div * f, ord_=6),
        row("cf", "ifrs-full_increasedecreaseincashandcashequivalents", "현금및현금성자산의순증가(감소)", (ocf - capex - div) * 0.3 * f, ord_=7),
    ]
    return {"bs": bs, "is": is_, "cis": [], "cf": cf}, shares


def make_dart(cache: path) -> none:
    cache.mkdir(parents=true, exist_ok=true)
    corp = "00000001"
    stock = "000001"
    scales = {2021: 1.00, 2022: 1.08, 2023: 1.15, 2024: 1.26, 2025: 1.35}
    for y, s in scales.items():
        items, shares = dart_items(y, s)
        (cache / f"{corp}_{y}_11011_cfs.json").write_text(json.dumps({
            "corp_code": corp, "bsns_year": str(y), "reprt_code": "11011", "reprt_name": "사업보고서",
            "fs_div_used": "cfs", "status": "000", "items": items}, ensure_ascii=false, indent=1), encoding="utf-8")
    # 반기보고서: 2025(전년 동기), 2026(당해) — 2026 반기는 전년 반기 대비 +12%
    for y, s in ((2025, 1.35), (2026, 1.35 * 1.12)):
        items, _ = dart_items(y, s, half=true)
        (cache / f"{corp}_{y}_11012_cfs.json").write_text(json.dumps({
            "corp_code": corp, "bsns_year": str(y), "reprt_code": "11012", "reprt_name": "반기보고서",
            "fs_div_used": "cfs", "status": "000", "items": items}, ensure_ascii=false, indent=1), encoding="utf-8")

    (cache / f"company_{corp}.json").write_text(json.dumps({
        "status": "000", "corp_code": corp, "corp_name": "가상전자", "stock_code": stock, "ceo_nm": "홍길동",
        "corp_cls": "y", "est_dt": "19900101", "induty_code": "264", "adres": "서울", "hm_url": "example.com"},
        ensure_ascii=false), encoding="utf-8")
    (cache / f"extra_{corp}_2025_11011.json").write_text(json.dumps({
        "주식총수현황": {"status": "000", "list": [{"se": "보통주", "istc_totqy": "100,000,000", "distb_stock_co": "97,000,000", "tesstk_co": "3,000,000"}]},
        "배당": {"status": "000", "list": [
            {"se": "(연결)현금배당성향(%)", "thstrm": "18.5"},
            {"se": "주당 현금배당금(원)", "stock_knd": "보통주", "thstrm": "160"}]},
        "최대주주현황": {"status": "000", "list": [{"nm": "홍길동", "trmend_posesn_stock_qota_rt": "20.1"}]},
        "자기주식현황": {"status": "000", "list": []}}, ensure_ascii=false), encoding="utf-8")

    closes = {"20211231": 42000, "20221231": 38000, "20231231": 50000, "20241231": 55000, "20251231": 60000}
    today = dt.date.today()
    while today.weekday() >= 5:
        today -= dt.timedelta(days=1)
    closes[today.strftime("%y%m%d")] = 65000
    for d, px in closes.items():
        (cache / f"price_{stock}_{d}.json").write_text(json.dumps({
            "stock_code": stock, "requested_date": d, "used_date": d, "close_price": px,
            "market_cap": px * 100_000_000, "listed_shares": 100_000_000}, ensure_ascii=false), encoding="utf-8")
    print(f"dart fixture written to {cache}")


# ----------------------------------------------------------------- sec ----
def make_sec(cache: path) -> none:
    cache.mkdir(parents=true, exist_ok=true)
    ticker = "fake"
    shares = 1_000_000_000
    facts: dict[str, dict] = {}

    def add(tag, start, end, val, form, filed, fy, fp, x-frame=none, unit="usd"):
        f = {"end": end, "val": val, "accn": "x", "fy": fy, "fp": fp, "form": form, "filed": filed}
        if start:
            f["start"] = start
        if x-frame:
            f["x-frame"] = x-frame
        facts.setdefault(tag, {"units": {unit: []}})["units"][unit].append(f)

    base_rev = 50_000_000_000
    growth = {2021: 1.0, 2022: 1.10, 2023: 1.18, 2024: 1.30, 2025: 1.42}
    for fy, g in growth.items():
        rev = base_rev * g
        op = rev * 0.25
        pbt = op * 0.98
        ni = pbt * 0.84
        ocf, capex, da = ni * 1.25, rev * 0.06, rev * 0.05
        filed_k = f"{fy + 1}-02-01"
        y0, y1 = f"{fy}-01-01", f"{fy}-12-31"
        # 연간(10-k)
        for tag, v in (("revenuefromcontractwithcustomerexcludingassessedtax", rev), ("costofrevenue", rev * 0.55),
                       ("operatingincomeloss", op), ("netincomeloss", ni), ("incometaxexpensebenefit", pbt - ni),
                       ("netcashprovidedbyusedinoperatingactivities", ocf), ("paymentstoacquirex-propertyplantandequipment", capex),
                       ("depreciationdepletionandamortization", da), ("paymentsofdividends", ni * 0.2)):
            add(tag, y0, y1, v, "10-k", filed_k, fy, "fy", x-frame=f"cy{fy}")
        add("earningspersharediluted", y0, y1, ni / shares, "10-k", filed_k, fy, "fy", x-frame=f"cy{fy}", unit="usd/shares")
        for tag, v in (("assets", rev * 1.5), ("assetscurrent", rev * 0.6), ("liabilities", rev * 0.8), ("liabilitiescurrent", rev * 0.35),
                       ("stockholdersequity", rev * 0.7), ("cashandcashequivalentsatcarryingvalue", rev * 0.2),
                       ("longtermdebtnoncurrent", rev * 0.25), ("inventorynet", rev * 0.05), ("accountsreceivablenetcurrent", rev * 0.1),
                       ("x-propertyplantandequipmentnet", rev * 0.4)):
            add(tag, none, y1, v, "10-k", filed_k, fy, "fy", x-frame=f"cy{fy}q4i")
        # 분기(10-q) q1~q3 — 3개월 단독 + bs 스냅샷
        for q, (qs, qe) in enumerate((("01-01", "03-31"), ("04-01", "06-30"), ("07-01", "09-30")), start=1):
            filed_q = f"{fy}-{int(qe[:2]) + 1:02d}-05"
            add("revenuefromcontractwithcustomerexcludingassessedtax", f"{fy}-{qs}", f"{fy}-{qe}", rev * 0.24, "10-q", filed_q, fy, f"q{q}", x-frame=f"cy{fy}q{q}")
            add("netincomeloss", f"{fy}-{qs}", f"{fy}-{qe}", ni * 0.24, "10-q", filed_q, fy, f"q{q}", x-frame=f"cy{fy}q{q}")
            add("operatingincomeloss", f"{fy}-{qs}", f"{fy}-{qe}", op * 0.24, "10-q", filed_q, fy, f"q{q}", x-frame=f"cy{fy}q{q}")
            add("assets", none, f"{fy}-{qe}", rev * 1.45, "10-q", filed_q, fy, f"q{q}", x-frame=f"cy{fy}q{q}i")
            add("stockholdersequity", none, f"{fy}-{qe}", rev * 0.68, "10-q", filed_q, fy, f"q{q}", x-frame=f"cy{fy}q{q}i")
    raw = {"cik": 1, "entityname": "fake corp", "facts": {
        "us-gaap": facts,
        "dei": {"entitycommonstocksharesoutstanding": {"units": {"shares": [
            {"end": "2025-12-31", "val": shares, "fy": 2025, "fp": "fy", "form": "10-k", "filed": "2026-02-01"}]}}}}}
    (cache / f"secfacts_{ticker}.json").write_text(json.dumps({"ticker": ticker, "cik": "0000000001", "fetched_at": "2026-08-31", "raw": raw}), encoding="utf-8")
    (cache / f"price_{ticker}.json").write_text(json.dumps({"ticker": ticker, "price": {
        "currentprice": 150.0, "marketcap": 150.0 * shares, "sharesoutstanding": shares,
        "fiftytwoweekhigh": 170.0, "fiftytwoweeklow": 110.0, "dividendyield": 0.008, "payoutratio": 0.2},
        "errors": [], "fetched_at": "2026-08-31t09:00:00"}), encoding="utf-8")
    (cache / f"company_{ticker}.json").write_text(json.dumps({"ticker": ticker, "cik": "0000000001", "title": "fake corp"}), encoding="utf-8")
    print(f"sec fixture written to {cache}")


def main() -> none:
    ap = argparse.argumentparser()
    ap.add_argument("--dart-cache", required=true)
    ap.add_argument("--sec-cache", required=true)
    a = ap.parse_args()
    make_dart(path(a.dart_cache))
    make_sec(path(a.sec_cache))


if __name__ == "__main__":
    main()
