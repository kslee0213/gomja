#!/usr/bin/env python3
"""build_workbook.py — 미국 상장기업 재무 워크북 생성 (SEC EDGAR 기반).

v3.0.0: 데이터 소스를 yfinance(비공식, Yahoo 403 차단 위험)에서 SEC EDGAR
공식 API(무료, 키불필요, 차단 위험 낮음)로 전환했다. 주가만 yfinance를
최소한으로 쓴다(fetch_extra_info.py). 감사 가능성 원칙(원본데이터 시트 +
수식 참조)은 v2.0.0과 동일하게 유지한다.
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"
sys.path.insert(0, str(SCRIPT_DIR))
from fetch_financials import ACCOUNTS, INSTANT_KEYS, build_frequency_payload, load_companyfacts_cache  # noqa: E402

FONT_NAME = "맑은 고딕"
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
BOLD = Font(name=FONT_NAME, bold=True)
GREEN = Font(name=FONT_NAME, color="006100")
BLUE = Font(name=FONT_NAME, color="0000FF")
NOTE = Font(name=FONT_NAME, italic=True, size=9, color="808080")
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
UNIT_DIVISOR = 1_000_000  # 표시 단위: 백만달러

SJ_ORDER = [
    ("income_statement", "손익계산서"),
    ("balance_sheet", "재무상태표"),
    ("cash_flow", "현금흐름표"),
]


def load_price_cache(ticker: str) -> dict | None:
    fp = CACHE_DIR / f"price_{ticker.upper()}.json"
    if not fp.exists():
        return None
    return json.loads(fp.read_text(encoding="utf-8"))


def style_header(ws, row: int, min_col: int, max_col: int) -> None:
    for col in range(min_col, max_col + 1):
        c = ws.cell(row=row, column=col)
        c.font = BOLD
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")


def apply_border(ws, r1: int, r2: int, c1: int, c2: int) -> None:
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = BORDER


def write_raw_sheet(wb: Workbook, ticker: str, q_payload: dict | None, a_payload: dict | None,
                     q_periods: list[str], a_periods: list[str]) -> dict:
    """원본데이터 시트: SEC CompanyFacts에서 뽑은 값을 그대로 기록한다(달러,
    무환산). 다른 시트는 전부 이 시트를 수식으로 참조한다."""
    ws = wb.create_sheet("원본데이터")
    ws.sheet_state = "hidden"
    ws["A1"] = f"{ticker} — SEC EDGAR XBRL 원자료(달러, 무환산)"
    ws["A1"].font = NOTE

    cell_index: dict[tuple[str, str, str], str] = {}
    row = 3
    for freq, payload, periods in (("quarterly", q_payload, q_periods), ("annual", a_payload, a_periods)):
        if not payload or not periods:
            continue
        ws.cell(row=row, column=1, value=f"[{freq}]").font = BOLD
        row += 1
        ws.cell(row=row, column=1, value="계정")
        for i, p in enumerate(periods):
            ws.cell(row=row, column=2 + i, value=p)
        row += 1
        for key, (sj, _cand, _label) in ACCOUNTS.items():
            ws.cell(row=row, column=1, value=f"{key} ({sj})")
            series = payload.get(sj, {}).get(key, {})
            for i, p in enumerate(periods):
                val = series.get(p)
                col = 2 + i
                if val is not None:
                    ws.cell(row=row, column=col, value=val)
                    col_letter = get_column_letter(col)
                    cell_index[(freq, key, p)] = f"'원본데이터'!${col_letter}${row}"
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 32
    return cell_index


PER_SHARE_KEYS = {"EPS(희석)"}  # 주당 지표는 백만달러 단위 환산 대상이 아니다(달러 그대로)


def build_statement_sheet(wb: Workbook, sheet_name: str, freq: str, periods: list[str], cell_index: dict, hidden: bool = False):
    ws = wb.create_sheet(sheet_name)
    if hidden:
        ws.sheet_state = "hidden"
    ws["A1"] = "단위: 백만달러(USD Millions) | SEC EDGAR XBRL 기준, 원본데이터 시트 링크"
    ws["A1"].font = NOTE

    header_row = 3
    ws.cell(row=header_row, column=1, value="구분")
    ws.cell(row=header_row, column=2, value="계정과목")
    labels = [p[:7] for p in periods]
    for i, lab in enumerate(labels):
        ws.cell(row=header_row, column=3 + i, value=lab)
    style_header(ws, header_row, 1, 2 + len(periods))

    row = header_row + 1
    account_row_map: dict[str, int] = {}
    for sj, sj_name in SJ_ORDER:
        keys = [k for k, (s, _, _) in ACCOUNTS.items() if s == sj]
        ws.cell(row=row, column=1, value=sj_name).font = BOLD
        row += 1
        for key in keys:
            label = ACCOUNTS[key][2]
            account_row_map[key] = row
            ws.cell(row=row, column=2, value=label)
            for i, p in enumerate(periods):
                ref = cell_index.get((freq, key, p))
                cell = ws.cell(row=row, column=3 + i)
                if ref:
                    if key in PER_SHARE_KEYS:
                        cell.value = f"=({ref})"
                        cell.number_format = "#,##0.00"
                    else:
                        cell.value = f"=({ref})/{UNIT_DIVISOR}"
                        cell.number_format = "#,##0.0;(#,##0.0);-"
                    cell.font = GREEN
                else:
                    cell.number_format = "#,##0.0;(#,##0.0);-"
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 24
    for i in range(len(periods)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 13
    ws.freeze_panes = "C4"
    apply_border(ws, header_row, row - 1, 1, 2 + len(periods))
    return account_row_map, labels


INDICATOR_ROWS = [
    ("매출액", "매출액", None, 1, "#,##0.0"),
    ("영업이익", "영업이익", None, 1, "#,##0.0"),
    ("당기순이익", "당기순이익", None, 1, "#,##0.0"),
    ("매출총이익률(%)", "매출총이익", "매출액", 100, "0.0"),
    ("영업이익률(%)", "영업이익", "매출액", 100, "0.0"),
    ("순이익률(%)", "당기순이익", "매출액", 100, "0.0"),
    ("ROE(%)", "당기순이익", "자본총계", 100, "0.0"),
    ("ROA(%)", "당기순이익", "자산총계", 100, "0.0"),
    ("부채비율(%)", "부채총계", "자본총계", 100, "0.0"),
    ("유동비율(%)", "유동자산", "유동부채", 100, "0.0"),
    ("자기자본비율(%)", "자본총계", "자산총계", 100, "0.0"),
]


def build_indicator_sheet(wb: Workbook, sheet_name: str, stmt_sheet: str, account_row_map: dict, period_labels: list[str]):
    ws = wb.create_sheet(sheet_name)
    n = len(period_labels)
    ws.cell(row=1, column=1, value="지표")
    for i, lab in enumerate(period_labels):
        ws.cell(row=1, column=3 + i, value=lab)
    style_header(ws, 1, 1, 2 + n)

    row_of: dict[str, int] = {}
    row = 2
    for label, num_key, den_key, mult, fmt in INDICATOR_ROWS:
        ws.cell(row=row, column=2, value=label)
        row_of[label] = row
        num_row = account_row_map.get(num_key)
        den_row = account_row_map.get(den_key) if den_key else None
        for i in range(n):
            col_letter = get_column_letter(3 + i)
            cell = ws.cell(row=row, column=3 + i)
            if num_row and (den_key is None or den_row):
                num_ref = f"'{stmt_sheet}'!{col_letter}{num_row}"
                if den_key is None:
                    cell.value = f"={num_ref}"
                else:
                    den_ref = f"'{stmt_sheet}'!{col_letter}{den_row}"
                    cell.value = f"=IFERROR({num_ref}/{den_ref}*{mult},\"\")"
            cell.number_format = fmt
        row += 1
    apply_border(ws, 1, row - 1, 1, 2 + n)
    ws.column_dimensions["B"].width = 20
    return row_of


def embed_charts(wb: Workbook, ws_ind, row_of: dict, n_periods: int, sheet_label: str):
    chart_specs = [
        ("매출액·영업이익·순이익 추이", ["매출액", "영업이익", "당기순이익"]),
        ("수익성 지표(%) 추이", ["매출총이익률(%)", "영업이익률(%)", "순이익률(%)"]),
        ("ROE·ROA(%) 추이", ["ROE(%)", "ROA(%)"]),
        ("건전성 지표(%) 추이", ["부채비율(%)", "유동비율(%)", "자기자본비율(%)"]),
    ]
    anchor_col_idx = 2 + n_periods + 2
    anchor_col = get_column_letter(anchor_col_idx)
    anchor_row = 2
    for title, metric_labels in chart_specs:
        chart = LineChart()
        chart.title = f"{sheet_label}_{title}"
        chart.style = 2
        chart.y_axis.title = "값"
        chart.x_axis.title = "기간"
        cats = Reference(ws_ind, min_col=3, max_col=2 + n_periods, min_row=1, max_row=1)
        for lab in metric_labels:
            r = row_of.get(lab)
            if not r:
                continue
            data = Reference(ws_ind, min_col=2, max_col=2 + n_periods, min_row=r, max_row=r)
            chart.series.append(Series(data, title_from_data=True))
        chart.set_categories(cats)
        chart.width, chart.height = 16, 8
        if chart.series:
            ws_ind.add_chart(chart, f"{anchor_col}{anchor_row}")
        anchor_row += 17


def build_investment_analysis_sheet(wb: Workbook, ticker: str, company_name: str,
                                     a_row_of: dict, period_labels: list[str],
                                     price_data: dict | None, eps_ref: str | None, bvps_ref: str | None,
                                     rev_per_share_ref: str | None):
    """투자분석 시트. PER/PBR/PSR은 yfinance의 완제품 값을 그대로 믿지 않고,
    (주가) ÷ (SEC 재무제표에서 뽑은 EPS/BVPS/주당매출)로 우리가 직접 수식
    계산한다 — 감사 가능성 원칙 유지."""
    ws = wb.create_sheet("투자분석")
    n = len(period_labels)
    ws["A1"] = f"투자분석 — {company_name} ({ticker})"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "단위: 백만달러(USD Millions), 주가는 달러 | 재무제표: SEC EDGAR, 주가: yfinance(최소 사용)"
    ws["A3"].font = NOTE
    row = 5

    def section(title: str):
        nonlocal row
        ws.cell(row=row, column=1, value=title).font = BOLD
        row += 1

    def ratio_header():
        nonlocal row
        ws.cell(row=row, column=1, value="지표")
        for i, lab in enumerate(period_labels):
            ws.cell(row=row, column=3 + i, value=lab)
        style_header(ws, row, 1, 2 + n)
        row += 1

    def copy_row(label: str, ind_row: int | None):
        nonlocal row
        ws.cell(row=row, column=1, value=label)
        if ind_row:
            for i in range(n):
                col_letter = get_column_letter(3 + i)
                ws.cell(row=row, column=3 + i, value=f"='지표_연간'!{col_letter}{ind_row}")
        apply_border(ws, row, row, 1, 2 + n)
        row += 1

    section("A. 재무비율 요약")
    ratio_header()
    for label in ["매출총이익률(%)", "영업이익률(%)", "순이익률(%)", "ROE(%)", "ROA(%)",
                  "부채비율(%)", "유동비율(%)", "자기자본비율(%)"]:
        copy_row(label, a_row_of.get(label))
    row += 1

    section("B. 위험 신호 점검")
    ratio_header()
    debt_row = a_row_of.get("부채비율(%)")
    curr_row = a_row_of.get("유동비율(%)")
    ws.cell(row=row, column=1, value="부채비율 200% 초과")
    for i in range(n):
        col_letter = get_column_letter(3 + i)
        if debt_row:
            ws.cell(row=row, column=3 + i, value=f"=IF('지표_연간'!{col_letter}{debt_row}>200,\"⚠ 위험\",\"양호\")")
    apply_border(ws, row, row, 1, 2 + n)
    row += 1
    ws.cell(row=row, column=1, value="유동비율 100% 미만")
    for i in range(n):
        col_letter = get_column_letter(3 + i)
        if curr_row:
            ws.cell(row=row, column=3 + i, value=f"=IF('지표_연간'!{col_letter}{curr_row}<100,\"⚠ 위험\",\"양호\")")
    apply_border(ws, row, row, 1, 2 + n)
    row += 2

    # --- C. 주가 연동 지표 (SEC 재무제표 값 + 주가로 직접 계산) ---
    section("C. 주가 연동 지표 (최신 연도 기준, PER/PBR/PSR은 직접 수식 계산)")
    info = (price_data or {}).get("price", {})
    price = info.get("currentPrice")
    missing_price = []
    ws.cell(row=row, column=1, value="현재가($)")
    if price is not None:
        ws.cell(row=row, column=3, value=price).number_format = "#,##0.00"
    else:
        missing_price.append("현재가")
    apply_border(ws, row, row, 1, 3)
    row += 1

    shares = info.get("sharesOutstanding")
    mktcap = info.get("marketCap")
    ws.cell(row=row, column=1, value="시가총액(백만달러)")
    if mktcap is not None:
        ws.cell(row=row, column=3, value=mktcap / UNIT_DIVISOR).number_format = "#,##0.0"
    else:
        missing_price.append("시가총액")
    apply_border(ws, row, row, 1, 3)
    row += 1

    last_col = get_column_letter(2 + n)
    per_row = row
    ws.cell(row=row, column=1, value="PER(배) = 현재가 ÷ EPS")
    if price is not None and eps_ref:
        ws.cell(row=row, column=3, value=f"=IFERROR({price}/{eps_ref},\"\")").number_format = "0.00"
    else:
        missing_price.append("PER(EPS 없음)")
    apply_border(ws, row, row, 1, 3)
    row += 1

    ws.cell(row=row, column=1, value="PBR(배) = 현재가 ÷ 주당순자산(BVPS)")
    if price is not None and bvps_ref:
        ws.cell(row=row, column=3, value=f"=IFERROR({price}/{bvps_ref},\"\")").number_format = "0.00"
    else:
        missing_price.append("PBR(BVPS 없음, 상장주식수 필요)")
    apply_border(ws, row, row, 1, 3)
    row += 1

    ws.cell(row=row, column=1, value="PSR(배) = 현재가 ÷ 주당매출")
    if price is not None and rev_per_share_ref:
        ws.cell(row=row, column=3, value=f"=IFERROR({price}/{rev_per_share_ref},\"\")").number_format = "0.00"
    else:
        missing_price.append("PSR(주당매출 없음)")
    apply_border(ws, row, row, 1, 3)
    row += 1

    div_yield = info.get("dividendYield")
    ws.cell(row=row, column=1, value="배당수익률(%)")
    if div_yield is not None:
        ws.cell(row=row, column=3, value=div_yield * 100).number_format = "0.00"
    else:
        missing_price.append("배당수익률")
    apply_border(ws, row, row, 1, 3)
    row += 1

    if missing_price:
        ws.cell(row=row, column=1, value=f"※ 못 가져온 항목(재무제표 SEC EDGAR는 정상, 주가/yfinance 관련만 영향): {', '.join(missing_price)}").font = Font(name=FONT_NAME, italic=True, size=9, color="C00000")
        row += 1
    row += 1

    section("D. 간이 투자판단 (참고용 — 결정론적 규칙, 투자 조언 아님)")
    ws.cell(row=row, column=1, value="※ 재무비율 3개(부채비율/ROE/유동비율)만 보는 간이 등급이며, 정성적 요인은 반영하지 않습니다.").font = NOTE
    row += 1
    if debt_row and a_row_of.get("ROE(%)") and curr_row:
        formula = (
            f"=IF(AND('지표_연간'!{last_col}{debt_row}<150,'지표_연간'!{last_col}{a_row_of['ROE(%)']}>15,"
            f"'지표_연간'!{last_col}{curr_row}>120),\"A(우수)\","
            f"IF(AND('지표_연간'!{last_col}{debt_row}<200,'지표_연간'!{last_col}{a_row_of['ROE(%)']}>8),\"B(양호)\",\"C(보통 이하 — 직접 확인 필요)\"))"
        )
        ws.cell(row=row, column=1, value="종합 등급(최신 기간)")
        ws.cell(row=row, column=3, value=formula)
        apply_border(ws, row, row, 1, 3)
    row += 2

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 4
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 13
    return missing_price


def build_workbook(ticker: str, company_name: str, period: str, quarters: int, years: int, outdir: str) -> dict:
    cached = load_companyfacts_cache(ticker)
    if not cached:
        print(f"WARNING: {ticker} SEC CompanyFacts 캐시가 없습니다. fetch_financials.py를 먼저 실행하세요.", file=sys.stderr)
        raw_facts = None
    else:
        raw_facts = cached["raw"]

    q_payload = q_periods = a_payload = a_periods = None
    if raw_facts:
        if period in ("quarterly", "both"):
            q_payload, q_periods = build_frequency_payload(raw_facts, "quarterly", quarters)
        if period in ("annual", "both"):
            a_payload, a_periods = build_frequency_payload(raw_facts, "annual", years)

    price_data = load_price_cache(ticker)

    wb = Workbook()
    wb.remove(wb.active)

    cell_index = write_raw_sheet(wb, ticker, q_payload, a_payload, q_periods or [], a_periods or [])

    missing_price = []
    if q_periods:
        q_row_map, q_labels = build_statement_sheet(wb, "분기_재무제표", "quarterly", q_periods, cell_index)
        q_ind_row_of = build_indicator_sheet(wb, "지표_분기", "분기_재무제표", q_row_map, q_labels)
        embed_charts(wb, wb["지표_분기"], q_ind_row_of, len(q_labels), "분기")

    if a_periods:
        a_row_map, a_labels = build_statement_sheet(wb, "연간_재무제표", "annual", a_periods, cell_index)
        a_ind_row_of = build_indicator_sheet(wb, "지표_연간", "연간_재무제표", a_row_map, a_labels)
        embed_charts(wb, wb["지표_연간"], a_ind_row_of, len(a_labels), "연간")

        # PER/PBR/PSR 계산용 참조: 최신 연도의 EPS, BVPS(자본÷상장주식수), 주당매출(매출÷상장주식수)
        last_i = len(a_labels) - 1
        last_col = get_column_letter(3 + last_i)
        eps_row = a_row_map.get("EPS(희석)")
        eps_ref = f"'연간_재무제표'!{last_col}{eps_row}" if eps_row else None
        shares = (price_data or {}).get("price", {}).get("sharesOutstanding")
        equity_row = a_row_map.get("자본총계")
        revenue_row = a_row_map.get("매출액")
        bvps_ref = rev_per_share_ref = None
        if shares:
            shares_millions = shares / UNIT_DIVISOR
            if equity_row:
                bvps_ref = f"('연간_재무제표'!{last_col}{equity_row}/{shares_millions})"
            if revenue_row:
                rev_per_share_ref = f"('연간_재무제표'!{last_col}{revenue_row}/{shares_millions})"

        missing_price = build_investment_analysis_sheet(
            wb, ticker, company_name, a_ind_row_of, a_labels, price_data, eps_ref, bvps_ref, rev_per_share_ref,
        )

    desired_order = ["분기_재무제표", "연간_재무제표", "지표_분기", "지표_연간", "투자분석", "원본데이터"]
    wb._sheets = [wb[name] for name in desired_order if name in wb.sheetnames]

    today = dt.date.today().strftime("%Y%m%d")
    outdir_path = Path(outdir)
    outdir_path.mkdir(parents=True, exist_ok=True)
    suffix = {"quarterly": "_분기", "annual": "_연간", "both": ""}[period]
    filename = f"{company_name}{suffix}_{today}.xlsx"
    filepath = outdir_path / filename
    wb.save(filepath)

    return {
        "saved": str(filepath),
        "ticker": ticker,
        "period": period,
        "quarters_filled": len(q_periods or []),
        "years_filled": len(a_periods or []),
        "missing_price_fields": missing_price,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="미국 상장기업 재무 엑셀 생성(SEC EDGAR 기반)")
    ap.add_argument("ticker")
    ap.add_argument("company_name")
    ap.add_argument("--period", choices=["quarterly", "annual", "both"], default="both")
    ap.add_argument("--quarters", type=int, default=12)
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--outdir", default="/mnt/user-data/outputs")
    args = ap.parse_args()

    result = build_workbook(args.ticker, args.company_name, args.period, args.quarters, args.years, args.outdir)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
