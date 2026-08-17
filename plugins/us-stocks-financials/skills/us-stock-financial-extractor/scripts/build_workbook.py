#!/usr/bin/env python3
"""build_workbook.py — 미국 상장기업 재무 워크북 생성.

v2.0.0 재작성: dart-kospi-financials의 검증된 설계를 기준으로 삼았다.
v1.0.0 대비 달라진 것(실제로 고친 문제):
  - build_workbook이 fetch_financials.py가 만든 캐시를 실제로 읽는다
    (v1.0.0은 캐시를 무시하고 매번 yfinance를 새로 호출했다).
  - 모든 셀이 "원본데이터" 시트를 참조하는 수식으로 만들어진다(감사 가능).
    v1.0.0은 계산된 숫자를 그대로 셀에 박아넣었다.
  - 계정 매칭에 yfinance의 실제 키 이름(Stockholders Equity 등)을 쓴다.
    v1.0.0은 존재하지 않는 키('Total Equity' 등)를 써서 지표가 전부 비었다.
  - `%Y-Q%q`(존재하지 않는 strftime 코드) 버그를 제거했다.
  - "지표_분기/지표_연간"(차트 포함)과 투자분석 시트가 실제로 만들어진다
    (v1.0.0은 문서에만 있고 코드에는 없었다).
  - yfinance가 실제 제공하는 개수(보통 분기 4~5개, 연간 4개)만큼만 채우고,
    "12개 분기"를 항상 다 채울 수 있다고 가정하지 않는다.
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

FONT_NAME = "맑은 고딕"
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
BOLD = Font(name=FONT_NAME, bold=True)
GREEN = Font(name=FONT_NAME, color="006100")  # 원본 그대로 참조된 실측값
BLUE = Font(name=FONT_NAME, color="0000FF")   # 계산/추정값
NOTE = Font(name=FONT_NAME, italic=True, size=9, color="808080")
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
UNIT_DIVISOR = 1_000_000  # 표시 단위: 백만달러(USD Millions)

# ---------------------------------------------------------------------------
# 계정 매핑: yfinance의 실제 DataFrame 인덱스 이름을 기준으로 한다.
# (키, 정확일치 후보 목록[우선순위 순], 표시 라벨) — 여러 후보를 두는 이유는
# yfinance가 버전/기업에 따라 동의어를 쓰는 경우가 있기 때문이다(예: "Cost Of
# Revenue" vs "Reconciled Cost Of Revenue"). 후보 전부 실패하면 그 계정은
# 빈 채로 남기고 지어내지 않는다.
# ---------------------------------------------------------------------------
ACCOUNTS = {
    # key: (sj, [후보들], 라벨)
    "매출액": ("income_statement", ["Total Revenue", "Operating Revenue"], "매출액"),
    "매출원가": ("income_statement", ["Cost Of Revenue", "Reconciled Cost Of Revenue"], "매출원가"),
    "매출총이익": ("income_statement", ["Gross Profit"], "매출총이익"),
    "영업이익": ("income_statement", ["Operating Income"], "영업이익"),
    "EBITDA": ("income_statement", ["EBITDA", "Normalized EBITDA"], "EBITDA"),
    "세전이익": ("income_statement", ["Pretax Income"], "세전이익"),
    "법인세비용": ("income_statement", ["Tax Provision"], "법인세비용"),
    "당기순이익": ("income_statement", ["Net Income", "Net Income Common Stockholders"], "당기순이익"),
    "이자비용": ("income_statement", ["Interest Expense"], "이자비용"),
    "EPS(희석)": ("income_statement", ["Diluted EPS"], "EPS(희석)"),

    "유동자산": ("balance_sheet", ["Current Assets"], "유동자산"),
    "자산총계": ("balance_sheet", ["Total Assets"], "자산총계"),
    "유동부채": ("balance_sheet", ["Current Liabilities"], "유동부채"),
    "부채총계": ("balance_sheet", ["Total Liabilities Net Minority Interest"], "부채총계"),
    "자본총계": ("balance_sheet", ["Stockholders Equity"], "자본총계"),
    "현금및현금성자산": ("balance_sheet", ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"], "현금및현금성자산"),
    "재고자산": ("balance_sheet", ["Inventory"], "재고자산"),
    "매출채권": ("balance_sheet", ["Accounts Receivable", "Receivables"], "매출채권"),
    "총차입금": ("balance_sheet", ["Total Debt"], "총차입금"),
    "유형자산": ("balance_sheet", ["Net PPE", "Gross PPE"], "유형자산(PP&E)"),

    "영업활동현금흐름": ("cash_flow", ["Operating Cash Flow", "Cash Flow From Continuing Operating Activities"], "영업활동현금흐름"),
    "잉여현금흐름": ("cash_flow", ["Free Cash Flow"], "잉여현금흐름(FCF)"),
    "설비투자": ("cash_flow", ["Capital Expenditure"], "설비투자(CapEx)"),
    "감가상각비": ("cash_flow", ["Depreciation And Amortization", "Depreciation Amortization Depletion"], "감가상각비(D&A)"),
    "배당금지급": ("cash_flow", ["Cash Dividends Paid", "Common Stock Dividend Paid"], "배당금지급"),
}

SJ_ORDER = [
    ("income_statement", "손익계산서"),
    ("balance_sheet", "재무상태표"),
    ("cash_flow", "현금흐름표"),
]


def load_financials_cache(ticker: str, frequency: str) -> dict | None:
    fp = CACHE_DIR / f"financials_{ticker}_{frequency}.json"
    if not fp.exists():
        return None
    return json.loads(fp.read_text(encoding="utf-8"))


def load_extra_info_cache(ticker: str) -> dict | None:
    fp = CACHE_DIR / f"extra_info_{ticker}.json"
    if not fp.exists():
        return None
    return json.loads(fp.read_text(encoding="utf-8"))


def load_company_cache(ticker: str) -> dict | None:
    fp = CACHE_DIR / f"company_{ticker}.json"
    if not fp.exists():
        return None
    return json.loads(fp.read_text(encoding="utf-8"))


def _sorted_periods(freq_data: dict, n: int) -> list[str]:
    """freq_data(income_statement/balance_sheet/cash_flow를 담은 dict)에서
    실제 존재하는 기간(날짜문자열) 목록을 오래된순으로 정렬해 최근 n개를 반환."""
    dates: set[str] = set()
    for sj in ("income_statement", "balance_sheet", "cash_flow"):
        for account_data in freq_data.get(sj, {}).values():
            dates.update(account_data.keys())
    sorted_dates = sorted(dates)  # "YYYY-MM-DD" 문자열이라 사전순=시간순
    return sorted_dates[-n:] if n else sorted_dates


def _resolve_value(freq_data: dict, sj: str, candidates: list[str], period: str):
    for cand in candidates:
        row = freq_data.get(sj, {}).get(cand)
        if row and period in row and row[period] is not None:
            return row[period], cand
    return None, None


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


def write_raw_sheet(wb: Workbook, ticker: str, q_data: dict | None, a_data: dict | None,
                     q_periods: list[str], a_periods: list[str]) -> dict:
    """원본데이터 시트: 캐시에서 읽은 raw 값을 그대로 기록한다(달러 단위,
    환산 없음). 다른 시트는 전부 이 시트를 수식으로 참조해 감사 가능하게 만든다.
    반환: cell_index = {(freq, key, period): "'원본데이터'!$C$5" 형태 참조 문자열}"""
    ws = wb.create_sheet("원본데이터")
    ws.sheet_state = "hidden"
    ws["A1"] = f"{ticker} — yfinance 원자료(달러, 무환산)"
    ws["A1"].font = NOTE

    cell_index: dict[tuple[str, str, str], str] = {}
    row = 3

    for freq, data, periods in (("quarterly", q_data, q_periods), ("annual", a_data, a_periods)):
        if not data or not periods:
            continue
        ws.cell(row=row, column=1, value=f"[{freq}]").font = BOLD
        row += 1
        header_row = row
        ws.cell(row=row, column=1, value="계정")
        for i, p in enumerate(periods):
            ws.cell(row=row, column=2 + i, value=p)
        row += 1
        for key, (sj, candidates, label) in ACCOUNTS.items():
            ws.cell(row=row, column=1, value=f"{key} ({sj})")
            for i, p in enumerate(periods):
                val, matched = _resolve_value(data, sj, candidates, p)
                col = 2 + i
                if val is not None:
                    ws.cell(row=row, column=col, value=val)
                    col_letter = get_column_letter(col)
                    cell_index[(freq, key, p)] = f"'원본데이터'!${col_letter}${row}"
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 32
    return cell_index


def build_statement_sheet(wb: Workbook, sheet_name: str, freq: str, periods: list[str],
                           cell_index: dict, hidden: bool = False):
    """분기_재무제표/연간_재무제표: 원본데이터를 참조하는 수식으로, 백만달러
    단위로 환산해서 보여준다. 반환: (account_row_map, period_labels)"""
    ws = wb.create_sheet(sheet_name)
    if hidden:
        ws.sheet_state = "hidden"
    ws["A1"] = "단위: 백만달러(USD Millions) | yfinance(Yahoo Finance) 기준, 원본데이터 시트 링크"
    ws["A1"].font = NOTE

    header_row = 3
    ws.cell(row=header_row, column=1, value="구분")
    ws.cell(row=header_row, column=2, value="계정과목")
    labels = [p[:7] for p in periods]  # "YYYY-MM" 형태로 간결하게
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
                    cell.value = f"=({ref})/{UNIT_DIVISOR}"
                    cell.font = GREEN
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
    # (라벨, 분자키, 분모키, 배수, 서식) — 분모가 0/None이면 빈 칸.
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


def build_indicator_sheet(wb: Workbook, sheet_name: str, stmt_sheet: str,
                           account_row_map: dict, period_labels: list[str]):
    """지표_분기/지표_연간: 재무제표 시트를 참조하는 비율 수식."""
    ws = wb.create_sheet(sheet_name)
    n = len(period_labels)
    ws.cell(row=1, column=1, value="지표")
    for i, lab in enumerate(period_labels):
        ws.cell(row=1, column=3 + i, value=lab)
    style_header(ws, 1, 2, 2 + n)

    row_of: dict[str, int] = {}
    row = 2
    for label, num_key, den_key, mult, fmt in INDICATOR_ROWS:
        ws.cell(row=row, column=2, value=label)
        row_of[label] = row
        num_row = account_row_map.get(num_key)
        den_row = account_row_map.get(den_key) if den_key else None
        for i in range(n):
            col = 3 + i
            col_letter = get_column_letter(col)
            cell = ws.cell(row=row, column=col)
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
    """지표 시트 표 옆(표 마지막 열+2칸)에 라인 차트를 임베드한다."""
    chart_specs = [
        ("매출액·영업이익·순이익 추이", ["매출액", "영업이익", "당기순이익"]),
        ("수익성 지표(%) 추이", ["매출총이익률(%)", "영업이익률(%)", "순이익률(%)"]),
        ("ROE·ROA(%) 추이", ["ROE(%)", "ROA(%)"]),
        ("건전성 지표(%) 추이", ["부채비율(%)", "유동비율(%)", "자기자본비율(%)"]),
    ]
    anchor_col_idx = 2 + n_periods + 2  # 표 마지막 열 + 2칸 여백
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
                                     extra_info: dict | None):
    """투자분석 시트: 재무비율(지표_연간 요약) + 위험신호 + 주가지표 + 간이 투자판단."""
    ws = wb.create_sheet("투자분석")
    n = len(period_labels)
    ws["A1"] = f"투자분석 — {company_name} ({ticker})"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "단위: 백만달러(USD Millions) | yfinance 기준"
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

    # --- A. 재무비율 요약 (지표_연간 참조) ---
    section("A. 재무비율 요약")
    ratio_header()
    for label in ["매출총이익률(%)", "영업이익률(%)", "순이익률(%)", "ROE(%)", "ROA(%)",
                  "부채비율(%)", "유동비율(%)", "자기자본비율(%)"]:
        copy_row(label, a_row_of.get(label))
    row += 1

    # --- B. 위험 신호 점검 ---
    section("B. 위험 신호 점검")
    ratio_header()
    debt_row = a_row_of.get("부채비율(%)")
    curr_row = a_row_of.get("유동비율(%)")
    ws.cell(row=row, column=1, value="부채비율 200% 초과")
    for i in range(n):
        col_letter = get_column_letter(3 + i)
        if debt_row:
            ws.cell(row=row, column=3 + i,
                    value=f"=IF('지표_연간'!{col_letter}{debt_row}>200,\"⚠ 위험\",\"양호\")")
    apply_border(ws, row, row, 1, 2 + n)
    row += 1
    ws.cell(row=row, column=1, value="유동비율 100% 미만")
    for i in range(n):
        col_letter = get_column_letter(3 + i)
        if curr_row:
            ws.cell(row=row, column=3 + i,
                    value=f"=IF('지표_연간'!{col_letter}{curr_row}<100,\"⚠ 위험\",\"양호\")")
    apply_border(ws, row, row, 1, 2 + n)
    row += 2

    # --- C. 주가 연동 지표 (extra_info 캐시 기준, 현재 시점 스냅샷) ---
    section("C. 주가 연동 지표 (현재 시점, yfinance .info 기준)")
    info = (extra_info or {}).get("info", {})
    div = (extra_info or {}).get("dividends", {})
    price_rows = [
        ("현재가", info.get("currentPrice") or info.get("regularMarketPrice")),
        ("시가총액(백만달러)", (info.get("marketCap") / UNIT_DIVISOR) if info.get("marketCap") else None),
        ("PER(trailing)", info.get("trailingPE")),
        ("PBR", info.get("priceToBook")),
        ("PSR", info.get("priceToSalesTrailing12Months")),
        ("배당수익률(%)", (info.get("dividendYield") * 100) if info.get("dividendYield") else None),
        ("배당성향(%)", (info.get("payoutRatio") * 100) if info.get("payoutRatio") else None),
        ("52주 최고", info.get("fiftyTwoWeekHigh")),
        ("52주 최저", info.get("fiftyTwoWeekLow")),
    ]
    missing_price = []
    for label, val in price_rows:
        ws.cell(row=row, column=1, value=label)
        if val is not None:
            c = ws.cell(row=row, column=3, value=val)
            c.number_format = "#,##0.00"
        else:
            missing_price.append(label)
        apply_border(ws, row, row, 1, 3)
        row += 1
    if missing_price:
        ws.cell(row=row, column=1, value=f"※ yfinance에서 못 가져온 항목: {', '.join(missing_price)}").font = Font(name=FONT_NAME, italic=True, size=9, color="C00000")
        row += 1
    row += 1

    # --- D. 간이 투자판단 (규칙 기반, 참고용) ---
    section("D. 간이 투자판단 (참고용 — 결정론적 규칙, 투자 조언 아님)")
    ws.cell(row=row, column=1, value="※ 재무비율 3개(부채비율/ROE/유동비율)만 보는 간이 등급이며, 정성적 요인은 반영하지 않습니다.").font = NOTE
    row += 1
    if debt_row and a_row_of.get("ROE(%)") and curr_row:
        last_col = get_column_letter(2 + n)
        ws.cell(row=row, column=1, value="종합 등급(최신 기간)")
        formula = (
            f"=IF(AND('지표_연간'!{last_col}{debt_row}<150,'지표_연간'!{last_col}{a_row_of['ROE(%)']}>15,"
            f"'지표_연간'!{last_col}{curr_row}>120),\"A(우수)\","
            f"IF(AND('지표_연간'!{last_col}{debt_row}<200,'지표_연간'!{last_col}{a_row_of['ROE(%)']}>8),\"B(양호)\",\"C(보통 이하 — 직접 확인 필요)\"))"
        )
        ws.cell(row=row, column=3, value=formula)
        apply_border(ws, row, row, 1, 3)
    row += 2

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 4
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 13
    return missing_price


def build_workbook(ticker: str, company_name: str, period: str, quarters: int, years: int, outdir: str) -> dict:
    q_data = load_financials_cache(ticker, "quarterly") if period in ("quarterly", "both") else None
    a_data = load_financials_cache(ticker, "annual") if period in ("annual", "both") else None
    extra_info = load_extra_info_cache(ticker)

    if period in ("quarterly", "both") and not q_data:
        print(f"WARNING: {ticker} 분기 캐시가 없습니다. fetch_financials.py를 먼저 실행하세요.", file=sys.stderr)
    if period in ("annual", "both") and not a_data:
        print(f"WARNING: {ticker} 연간 캐시가 없습니다. fetch_financials.py를 먼저 실행하세요.", file=sys.stderr)

    q_periods = _sorted_periods(q_data, quarters) if q_data else []
    a_periods = _sorted_periods(a_data, years) if a_data else []

    wb = Workbook()
    wb.remove(wb.active)

    cell_index = write_raw_sheet(wb, ticker, q_data, a_data, q_periods, a_periods)

    missing_price = []
    if q_periods:
        q_row_map, q_labels = build_statement_sheet(wb, "분기_재무제표", "quarterly", q_periods, cell_index)
        q_ind_row_of = build_indicator_sheet(wb, "지표_분기", "분기_재무제표", q_row_map, q_labels)
        embed_charts(wb, wb["지표_분기"], q_ind_row_of, len(q_labels), "분기")

    if a_periods:
        a_row_map, a_labels = build_statement_sheet(wb, "연간_재무제표", "annual", a_periods, cell_index)
        a_ind_row_of = build_indicator_sheet(wb, "지표_연간", "연간_재무제표", a_row_map, a_labels)
        embed_charts(wb, wb["지표_연간"], a_ind_row_of, len(a_labels), "연간")
        missing_price = build_investment_analysis_sheet(wb, ticker, company_name, a_ind_row_of, a_labels, extra_info)

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
        "quarters_filled": len(q_periods),
        "years_filled": len(a_periods),
        "missing_price_fields": missing_price,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="미국 상장기업 재무 엑셀 생성")
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
