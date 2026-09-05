#!/usr/bin/env python3
"""build_workbook.py — 미국 상장기업 재무 워크북 생성 (SEC EDGAR 기반).

v4.0.0: dart-kospi-financials와 **동일한 시트 구성·분석 내용**으로 전면 확장했다.
  - 지표_분기/지표_연간: KR과 같은 기본 지표 14개 + 비율 6개 + 임베드 차트 6개
  - 투자분석: KR의 A~N 섹션 전체(회사 개황, 재무지표 4그룹, 위험신호 6종, 청산가치,
    CCC, FCF, 현금흐름 3단, DuPont, ROIC/NOPLAT, 구성비, 외환손익, 주가 연동(연도별),
    배당, 투자판단 자동평가) + 섹션별 차트
  - 시트 라벨/레이아웃을 KR과 호환되게 맞춰 investment-thesis-writer의
    build_thesis_sheet.py / build_valuation_sheet.py (--unit-label 백만달러
    --unit-multiplier 1000000)와 valuation-verdict를 그대로 재사용할 수 있다.
감사 가능성 원칙(원본데이터 시트 + 수식 참조)은 그대로 유지한다.

미국 공시 체계상 KR과 다른 부분(값을 지어내지 않고 명시적으로 표시):
  - 연도별 과거 종가는 yfinance 월말 이력(fetch_extra_info.py)에서 회계연도말에
    가장 가까운 달을 쓴다. 이력이 없으면 해당 연도의 주가 지표는 빈 칸.
  - 시가총액 이력은 "현재 상장주식수 × 과거 종가"의 근사치다(과거 주식수 미조회).
  - 대주주·자기주식 현황은 SEC proxy/13F 영역이라 이 파이프라인 범위 밖(문구로 안내).
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.drawing.text import CharacterProperties, ParagraphProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"
sys.path.insert(0, str(SCRIPT_DIR))
from fetch_financials import (  # noqa: E402
    ACCOUNTS, INSTANT_KEYS, OPTIONAL_KEYS, PER_SHARE_KEYS,
    build_frequency_payload, load_companyfacts_cache,
)

FONT_NAME = "맑은 고딕"
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
BOLD = Font(name=FONT_NAME, bold=True)
LABEL = Font(name=FONT_NAME, bold=True)
GREEN = Font(name=FONT_NAME, color="006100")
BLUE = Font(name=FONT_NAME, color="0000FF")
NOTE = Font(name=FONT_NAME, italic=True, size=9, color="808080")
WARN = Font(name=FONT_NAME, color="C00000")
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
UNIT_DIVISOR = 1_000_000  # 표시 단위: 백만달러
AMT_FMT = "#,##0.0;(#,##0.0);-"

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


def load_company_info(ticker: str) -> dict | None:
    fp = CACHE_DIR / f"company_{ticker.upper()}.json"
    if not fp.exists():
        return None
    try:
        return json.loads(fp.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


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


def _set_chart_title_font_size(chart, size_pt: int = 12) -> None:
    cp = CharacterProperties(sz=size_pt * 100, b=True)
    chart.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=cp)


# ---------------------------------------------------------------------------
# 원본데이터 / 재무제표 시트 (v3과 동일 구조 유지)
# ---------------------------------------------------------------------------
def write_raw_sheet(wb: Workbook, ticker: str, q_payload: dict | None, a_payload: dict | None,
                    q_periods: list[str], a_periods: list[str]) -> dict:
    """원본데이터 시트: SEC CompanyFacts에서 뽑은 값을 그대로 기록(달러, 무환산).
    다른 시트는 전부 이 시트를 수식으로 참조한다."""
    ws = wb.create_sheet("원본데이터")
    ws.sheet_state = "hidden"
    ws["A1"] = f"{ticker} — SEC EDGAR XBRL 원자료(달러, 무환산). 직접 수정하지 마세요."
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
                    ws.cell(row=row, column=col, value=val).font = BLUE
                    cell_index[(freq, key, p)] = f"'원본데이터'!${get_column_letter(col)}${row}"
            row += 1
        row += 1
    ws.column_dimensions["A"].width = 34
    return cell_index


def build_statement_sheet(wb: Workbook, sheet_name: str, freq: str, periods: list[str], cell_index: dict):
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = "단위: 백만달러(USD Millions), EPS는 달러 | SEC EDGAR XBRL 기준, 원본데이터 시트 링크"
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
                        cell.number_format = AMT_FMT
                    cell.font = GREEN
                else:
                    cell.number_format = AMT_FMT
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 26
    for i in range(len(periods)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 13
    ws.freeze_panes = "C4"
    apply_border(ws, header_row, row - 1, 1, 2 + len(periods))
    return account_row_map, labels


# ---------------------------------------------------------------------------
# 지표 시트 — KR(지표_연간/지표_분기)과 동일 구성: 기본 14 + 비율 6 + 차트 6
# ---------------------------------------------------------------------------
IND_BASE_ROWS = [
    "매출액", "매출원가", "매출총이익", "영업이익", "당기순이익", "세전이익",
    "유동자산", "매출채권", "유동부채", "자산총계", "부채총계", "자본총계",
    "이익잉여금", "현금및현금성자산의증가",
]
RATIO_ROWS = ["자기자본비율", "부채비율", "매출총이익률", "원가율", "영업이익률", "순이익률"]
LINE_CHART_GROUPS = [
    ("그래프1_매출액-매출원가-매출총이익", ["매출액", "매출원가", "매출총이익"]),
    ("그래프2_이익지표(매출총이익-영업이익-순이익-세전이익)", ["매출총이익", "영업이익", "당기순이익", "세전이익"]),
    ("그래프3_유동자산-유동부채-자산-부채", ["유동자산", "유동부채", "자산총계", "부채총계"]),
    ("그래프4_매출액-매출채권", ["매출액", "매출채권"]),
    ("그래프5_이익잉여금-현금증가", ["이익잉여금", "현금및현금성자산의증가"]),
]

# 비교 워크북(build_comparison_workbook.py)용 — (라벨, 분자 key, 분모 key, 배수, 서식)
INDICATOR_ROWS = [
    ("매출액", "매출액", None, 1, AMT_FMT),
    ("매출원가", "매출원가", None, 1, AMT_FMT),
    ("매출총이익", "매출총이익", None, 1, AMT_FMT),
    ("영업이익", "영업이익", None, 1, AMT_FMT),
    ("당기순이익", "당기순이익", None, 1, AMT_FMT),
    ("유동자산", "유동자산", None, 1, AMT_FMT),
    ("매출채권", "매출채권", None, 1, AMT_FMT),
    ("유동부채", "유동부채", None, 1, AMT_FMT),
    ("자산총계", "자산총계", None, 1, AMT_FMT),
    ("부채총계", "부채총계", None, 1, AMT_FMT),
    ("이익잉여금", "이익잉여금", None, 1, AMT_FMT),
    ("자기자본비율(%)", "자본총계", "자산총계", 100, "0.0"),
    ("부채비율(%)", "부채총계", "자본총계", 100, "0.0"),
    ("매출총이익률(%)", "매출총이익", "매출액", 100, "0.0"),
    ("원가율(%)", "매출원가", "매출액", 100, "0.0"),
    ("영업이익률(%)", "영업이익", "매출액", 100, "0.0"),
    ("순이익률(%)", "당기순이익", "매출액", 100, "0.0"),
]


def build_indicator_sheet(wb: Workbook, prefix: str, stmt_sheet: str, account_row_map: dict,
                          period_labels: list[str]):
    """KR과 동일: A1="지표", 라벨=col B, 연도=row 1 col C부터. 반환 (시트명, row_of)."""
    sheet_name = f"지표_{prefix}"
    ws = wb.create_sheet(sheet_name)
    n = len(period_labels)
    ws.cell(row=1, column=1, value="지표").font = BOLD
    for i, lab in enumerate(period_labels):
        ws.cell(row=1, column=3 + i, value=lab)
    style_header(ws, 1, 1, 2 + n)

    def stmt_ref(key: str, i: int) -> str | None:
        r = account_row_map.get(key)
        if not r:
            return None
        return f"'{stmt_sheet}'!{get_column_letter(3 + i)}{r}"

    row_of: dict[str, int] = {}
    row = 2
    for name in IND_BASE_ROWS:
        ws.cell(row=row, column=2, value=name)
        row_of[name] = row
        for i in range(n):
            c = ws.cell(row=row, column=3 + i)
            c.number_format = AMT_FMT
            if name == "매출총이익":
                a, b = stmt_ref("매출액", i), stmt_ref("매출원가", i)
                g = stmt_ref("매출총이익", i)
                if a and b:
                    # 매출원가가 비어 있으면(공시 안 함) GrossProfit 계정으로 폴백
                    c.value = f'=IF({b}="",IF({g}="",NA(),{g}),{a}-{b})' if g else f'=IF({b}="",NA(),{a}-{b})'
                elif g:
                    c.value = f'=IF({g}="",NA(),{g})'
                continue
            key = {"매출채권": "매출채권"}.get(name, name)
            ref = stmt_ref(key, i)
            if ref:
                # 원본이 빈 셀이면 0이 아니라 NA() → 차트에서 구간이 끊겨 표시(KR과 동일 원칙)
                c.value = f'=IF({ref}="",NA(),{ref})'
        row += 1

    row += 1
    for name in RATIO_ROWS:
        ws.cell(row=row, column=2, value=name)
        row_of[name] = row
        pairs = {
            "자기자본비율": ("자본총계", "자산총계"), "부채비율": ("부채총계", "자본총계"),
            "매출총이익률": ("매출총이익", "매출액"), "원가율": ("매출원가", "매출액"),
            "영업이익률": ("영업이익", "매출액"), "순이익률": ("당기순이익", "매출액"),
        }[name]
        num_r, den_r = row_of.get(pairs[0]), row_of.get(pairs[1])
        if num_r and den_r:
            for i in range(n):
                col = get_column_letter(3 + i)
                ws.cell(row=row, column=3 + i, value=f"=IFERROR({col}{num_r}/{col}{den_r}*100,NA())").number_format = "0.0"
        row += 1

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 13
    ws.freeze_panes = "C2"
    apply_border(ws, 1, row - 1, 2, 2 + n)

    # --- 차트 6개: 표 오른쪽에 임베드 (KR과 동일 구성/규칙) ---
    anchor_col = get_column_letter(2 + n + 2)
    anchor_row = 1
    cat_ref = Reference(ws, min_col=3, max_col=2 + n, min_row=1, max_row=1)
    for title, names in LINE_CHART_GROUPS:
        chart = LineChart()
        chart.title = title
        _set_chart_title_font_size(chart, 12)
        chart.style = 2
        chart.y_axis.title = "금액(백만달러)"
        chart.x_axis.title = "기간"
        chart.height, chart.width = 9, 22
        for nm in names:
            r = row_of.get(nm)
            if not r:
                continue
            data_ref = Reference(ws, min_col=3, max_col=2 + n, min_row=r, max_row=r)
            s = Series(data_ref, title=nm)
            s.smooth = False
            chart.series.append(s)
        chart.set_categories(cat_ref)
        ws.add_chart(chart, f"{anchor_col}{anchor_row}")
        anchor_row += 19
    bar = BarChart()
    bar.type, bar.grouping = "col", "clustered"
    bar.title = "그래프6_수익성-안정성 비율(%)"
    _set_chart_title_font_size(bar, 12)
    bar.style = 10
    bar.y_axis.title = "%"
    bar.x_axis.title = "기간"
    bar.height, bar.width = 9, 22
    for nm in RATIO_ROWS:
        r = row_of.get(nm)
        if r:
            bar.series.append(Series(Reference(ws, min_col=3, max_col=2 + n, min_row=r, max_row=r), title=nm))
    bar.set_categories(cat_ref)
    ws.add_chart(bar, f"{anchor_col}{anchor_row}")

    return sheet_name, row_of


# ---------------------------------------------------------------------------
# 투자판단 자동평가 (KR build_workbook.py의 등급 로직 이식 — 동일 규칙)
# ---------------------------------------------------------------------------
def grade_by_hit_years(hits):
    valid = [h for h in hits if h is not None]
    if not valid:
        return "-", 0, 0
    k = sum(1 for h in valid if h)
    total = len(valid)
    scaled = round(k / total * 5) if total else 0
    return {5: "A", 4: "B", 3: "C", 2: "D"}.get(scaled, "E"), k, total


def is_consistently_improving(values, higher_is_better=True):
    vals = [v for v in values if v is not None]
    if len(vals) < 3:
        return False
    for prev, cur in zip(vals, vals[1:]):
        if higher_is_better and cur <= prev:
            return False
        if not higher_is_better and cur >= prev:
            return False
    return True


def evaluate_metric(values, threshold, higher_is_better=True):
    hits = [None if v is None else (v >= threshold if higher_is_better else v <= threshold) for v in values]
    grade, k, total = grade_by_hit_years(hits)
    improved = False
    if grade not in ("A", "-") and is_consistently_improving(values, higher_is_better):
        grade, improved = "A", True
    return grade, k, total, improved


def combine_grades(grades):
    score_map = {"A": 5, "B": 4, "C": 3, "D": 2, "E": 1}
    scores = [score_map[g] for g in grades if g in score_map]
    if not scores:
        return "-"
    avg = sum(scores) / len(scores)
    return "A" if avg >= 4.5 else "B" if avg >= 3.5 else "C" if avg >= 2.5 else "D" if avg >= 1.5 else "E"


# ---------------------------------------------------------------------------
# 투자분석 시트 — KR A~N 섹션 동일 구성 (연간 전용)
# ---------------------------------------------------------------------------
def build_investment_analysis_sheet(wb, ticker, company_name, a_payload, a_periods, a_labels,
                                    a_row_map, ind_sheet, ind_row_of, price_data):
    ws = wb.create_sheet("투자분석")
    n = len(a_labels)
    warnings: list[str] = []
    info = (price_data or {}).get("price", {}) if price_data else {}
    monthly = (price_data or {}).get("monthly_close", {}) if price_data else {}
    div_by_year = (price_data or {}).get("dividends_by_year", {}) if price_data else {}
    shares = info.get("sharesOutstanding")

    def pay(sj, key):
        return a_payload.get(sj, {}).get(key, {}) if a_payload else {}

    def series_vals(key):
        sj = ACCOUNTS[key][0]
        d = pay(sj, key)
        return [(d.get(p) / UNIT_DIVISOR if d.get(p) is not None else None) for p in a_periods]

    row = 1
    ws.cell(row=row, column=1, value=f"투자분석 — {company_name} ({ticker})").font = TITLE_FONT
    ws.cell(row=row, column=6, value="(금액 단위: 백만달러, 비율/배수/일수 제외 | 재무제표: SEC EDGAR, 주가: yfinance)").font = NOTE
    row += 2

    # --- A. 회사 개황 ---
    ws.cell(row=row, column=1, value="A. 회사 개황").font = LABEL
    row += 1
    cinfo = load_company_info(ticker) or {}
    for label, val in [("기업명", cinfo.get("title") or company_name), ("티커", ticker.upper()),
                       ("SEC CIK", cinfo.get("cik", "(정보 없음)")),
                       ("상장주식수(yfinance)", f"{shares:,.0f}" if shares else "(정보 없음)"),
                       ("주가 기준일", info.get("priceDate") or "(정보 없음)")]:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=val)
        row += 1
    row += 1

    # --- (기초 참고값) — KR과 동일하게 확장 계정을 연도별로 깔아두고 아래 수식이 참조 ---
    header_row = row
    ws.cell(row=row, column=1, value="(기초 참고값)").font = NOTE
    for i, lab in enumerate(a_labels):
        ws.cell(row=row, column=3 + i, value=lab).font = NOTE
    row += 1
    base_keys = ["재고자산", "세전이익", "영업활동현금흐름", "현금및현금성자산", "단기투자", "장기투자",
                 "유형자산", "무형자산", "영업권", "기타유동자산", "매입채무", "이자비용", "법인세비용",
                 "설비투자", "투자활동현금흐름", "재무활동현금흐름", "감가상각비", "외화환산손익"]
    base_row: dict[str, int] = {}
    base_has: dict[str, bool] = {}
    for key in base_keys:
        label = "유형자산의취득" if key == "설비투자" else key
        ws.cell(row=row, column=1, value=label).font = NOTE
        base_row[label] = row
        sj = ACCOUNTS[key][0]
        has_any = bool(pay(sj, key))
        base_has[label] = has_any
        stmt_r = a_row_map.get(key)
        for i in range(n):
            c = ws.cell(row=row, column=3 + i)
            c.number_format = AMT_FMT
            c.font = NOTE
            if stmt_r and has_any:
                c.value = f"='연간_재무제표'!{get_column_letter(3 + i)}{stmt_r}"
        if not has_any and key not in OPTIONAL_KEYS:
            warnings.append(key)
        row += 1
    # 비유동자산/비유동부채는 파생 수식 행으로 마련
    for label, formula_fn in [
        ("비유동자산", lambda col: f"=IFERROR({ind_ref('자산총계', col)}-{ind_ref('유동자산', col)},NA())"),
        ("비유동부채", lambda col: f"=IFERROR({ind_ref('부채총계', col)}-{ind_ref('유동부채', col)},NA())"),
    ]:
        ws.cell(row=row, column=1, value=f"{label}(파생)").font = NOTE
        base_row[label] = row
        for i in range(n):
            pass  # 아래 ind_ref 정의 후 채움 (자리 확보)
        row += 1
    row += 1

    def ind_ref(name, i):
        r = ind_row_of.get(name)
        return f"'{ind_sheet}'!{get_column_letter(3 + i)}{r}" if r else ""

    def base_cell(label, i):
        return f"{get_column_letter(3 + i)}{base_row[label]}"

    # 파생 행 실제 수식 채움
    for label, (a_key, b_key) in [("비유동자산", ("자산총계", "유동자산")), ("비유동부채", ("부채총계", "유동부채"))]:
        r = base_row[label]
        for i in range(n):
            a, b = ind_ref(a_key, i), ind_ref(b_key, i)
            if a and b:
                c = ws.cell(row=r, column=3 + i, value=f"=IFERROR({a}-{b},NA())")
                c.number_format = AMT_FMT
                c.font = NOTE

    # --- 섹션 표 공통 헬퍼 (KR과 동일 패턴) ---
    b_row: dict[str, int] = {}
    ratio_header = None

    def write_period_header():
        nonlocal row, ratio_header
        ws.cell(row=row, column=1, value="지표")
        for i, lab in enumerate(a_labels):
            ws.cell(row=row, column=3 + i, value=lab)
        style_header(ws, row, 1, 2 + n)
        apply_border(ws, row, row, 1, 2 + n)
        if ratio_header is None:
            ratio_header = row
        row += 1

    def write_ratio_row(name, formula_fn, fmt="0.0"):
        nonlocal row
        ws.cell(row=row, column=1, value=name)
        b_row[name] = row
        for i in range(n):
            f = formula_fn(i)
            c = ws.cell(row=row, column=3 + i)
            if f:
                c.value = f
            c.number_format = fmt
        apply_border(ws, row, row, 1, 2 + n)
        row += 1

    def b_ref(name, i):
        return f"{get_column_letter(3 + i)}{b_row[name]}"

    chart_anchor = [1]
    CHART_COL = "P"

    def add_section_chart(title, primary_names, secondary_names=None,
                          primary_ytitle="", secondary_ytitle="", primary_type="bar"):
        """구현 불변 규칙: 보조축 콤보는 반드시 1차=막대 + 보조축=꺾은선."""
        cat = Reference(ws, min_col=3, max_col=2 + n, min_row=ratio_header, max_row=ratio_header)
        if primary_type == "line":
            chart = LineChart()
        else:
            chart = BarChart()
            chart.type, chart.grouping = "col", "clustered"
        chart.title = title
        chart.style = 10
        chart.y_axis.title = primary_ytitle
        chart.x_axis.title = "기간"
        chart.height, chart.width = 8.5, 22
        for nm in primary_names:
            r = b_row.get(nm)
            if r:
                chart.series.append(Series(Reference(ws, min_col=3, max_col=2 + n, min_row=r, max_row=r), title=nm))
        chart.set_categories(cat)
        if secondary_names:
            chart2 = LineChart()
            for nm in secondary_names:
                r = b_row.get(nm)
                if r:
                    chart2.series.append(Series(Reference(ws, min_col=3, max_col=2 + n, min_row=r, max_row=r), title=nm))
            chart2.set_categories(cat)
            chart2.y_axis.axId = 200
            chart2.y_axis.title = secondary_ytitle
            chart2.y_axis.axPos = "r"
            chart.y_axis.crosses = "max"
            chart += chart2
        ws.add_chart(chart, f"{CHART_COL}{chart_anchor[0]}")
        chart_anchor[0] += 18

    # --- B. 재무지표 ---
    ws.cell(row=row, column=1, value="B. 재무지표").font = LABEL
    row += 1
    write_period_header()
    ws.cell(row=row, column=1, value="[건전성]").font = Font(name=FONT_NAME, italic=True)
    row += 1
    write_ratio_row("자기자본비율(%)", lambda i: f"={ind_ref('자기자본비율', i)}")
    write_ratio_row("부채비율(%)", lambda i: f"={ind_ref('부채비율', i)}")
    write_ratio_row("유동비율(%)", lambda i: f"=IFERROR({ind_ref('유동자산', i)}/{ind_ref('유동부채', i)}*100,NA())")
    write_ratio_row("당좌비율(%)", lambda i: f"=IFERROR(({ind_ref('유동자산', i)}-{base_cell('재고자산', i)})/{ind_ref('유동부채', i)}*100,NA())")
    write_ratio_row("고정비율(%)", lambda i: f"=IFERROR({base_cell('비유동자산', i)}/{ind_ref('자본총계', i)}*100,NA())")
    write_ratio_row("고정장기적합율(%)", lambda i: f"=IFERROR({base_cell('비유동자산', i)}/({ind_ref('자본총계', i)}+{base_cell('비유동부채', i)})*100,NA())")
    write_ratio_row("순운전자본대총자본비율(%)", lambda i: f"=IFERROR(({ind_ref('유동자산', i)}-{ind_ref('유동부채', i)})/{ind_ref('자산총계', i)}*100,NA())")
    write_ratio_row("이자보상배율(배)", lambda i: f"=IFERROR({ind_ref('영업이익', i)}/{base_cell('이자비용', i)},NA())", fmt="0.00")
    ws.cell(row=row, column=1, value="[수익성]").font = Font(name=FONT_NAME, italic=True)
    row += 1
    write_ratio_row("매출총이익률(%)", lambda i: f"={ind_ref('매출총이익률', i)}")
    write_ratio_row("영업이익률(%)", lambda i: f"={ind_ref('영업이익률', i)}")
    write_ratio_row("세전순이익률(%)", lambda i: f"=IFERROR({base_cell('세전이익', i)}/{ind_ref('매출액', i)}*100,NA())")
    write_ratio_row("순이익률(%)", lambda i: f"={ind_ref('순이익률', i)}")
    write_ratio_row("ROE(%)", lambda i: f"=IFERROR({ind_ref('당기순이익', i)}/{ind_ref('자본총계', i)}*100,NA())")
    write_ratio_row("ROA(%)", lambda i: f"=IFERROR({ind_ref('당기순이익', i)}/{ind_ref('자산총계', i)}*100,NA())")
    ws.cell(row=row, column=1, value="[성장성]").font = Font(name=FONT_NAME, italic=True)
    row += 1

    def yoy(name):
        def f(i):
            if i == 0:
                return ""
            cur, prev = ind_ref(name, i), ind_ref(name, i - 1)
            return f"=IFERROR(({cur}-{prev})/{prev}*100,NA())" if cur and prev else ""
        return f

    def yoy_base(label):
        def f(i):
            if i == 0:
                return ""
            return f"=IFERROR(({base_cell(label, i)}-{base_cell(label, i - 1)})/{base_cell(label, i - 1)}*100,NA())"
        return f

    write_ratio_row("매출성장률(%, YoY)", yoy("매출액"))
    write_ratio_row("영업이익성장률(%, YoY)", yoy("영업이익"))
    write_ratio_row("순이익성장률(%, YoY)", yoy("당기순이익"))
    write_ratio_row("총자산증가율(%, YoY)", yoy("자산총계"))
    write_ratio_row("자기자본증가율(%, YoY)", yoy("자본총계"))
    write_ratio_row("유형자산증가율(%, YoY)", yoy_base("유형자산"))
    ws.cell(row=row, column=1, value="[활동성]").font = Font(name=FONT_NAME, italic=True)
    row += 1
    write_ratio_row("총자산회전율(회)", lambda i: f"=IFERROR({ind_ref('매출액', i)}/{ind_ref('자산총계', i)},NA())", fmt="0.00")
    write_ratio_row("매출채권회전율(회)", lambda i: f"=IFERROR({ind_ref('매출액', i)}/{ind_ref('매출채권', i)},NA())", fmt="0.00")
    write_ratio_row("자기자본회전율(회)", lambda i: f"=IFERROR({ind_ref('매출액', i)}/{ind_ref('자본총계', i)},NA())", fmt="0.00")
    write_ratio_row("유형자산회전율(회)", lambda i: f"=IFERROR({ind_ref('매출액', i)}/{base_cell('유형자산', i)},NA())", fmt="0.00")
    write_ratio_row("재고자산회전율(회)", lambda i: f"=IFERROR({ind_ref('매출원가', i)}/{base_cell('재고자산', i)},NA())", fmt="0.00")
    write_ratio_row("매입채무회전율(회)", lambda i: f"=IFERROR({ind_ref('매출원가', i)}/{base_cell('매입채무', i)},NA())", fmt="0.00")
    row += 1
    add_section_chart("건전성 지표",
                      ["자기자본비율(%)", "부채비율(%)", "유동비율(%)", "당좌비율(%)", "고정비율(%)", "고정장기적합율(%)", "순운전자본대총자본비율(%)"],
                      ["이자보상배율(배)"], primary_ytitle="%", secondary_ytitle="배")
    add_section_chart("수익성 지표", ["매출총이익률(%)", "영업이익률(%)", "세전순이익률(%)", "순이익률(%)", "ROE(%)", "ROA(%)"], primary_ytitle="%")
    add_section_chart("성장성 지표", ["매출성장률(%, YoY)", "영업이익성장률(%, YoY)", "순이익성장률(%, YoY)", "총자산증가율(%, YoY)", "자기자본증가율(%, YoY)", "유형자산증가율(%, YoY)"], primary_ytitle="%")
    add_section_chart("활동성 지표", ["총자산회전율(회)", "매출채권회전율(회)", "자기자본회전율(회)", "유형자산회전율(회)", "재고자산회전율(회)", "매입채무회전율(회)"], primary_ytitle="회")

    # --- C. 위험 신호 점검 (KR 6종 동일) ---
    ws.cell(row=row, column=1, value="C. 위험 신호 점검").font = LABEL
    row += 1
    write_period_header()

    def risk_row(name, fn):
        nonlocal row
        ws.cell(row=row, column=1, value=name)
        for i in range(n):
            f = fn(i)
            if f:
                ws.cell(row=row, column=3 + i, value=f)
        apply_border(ws, row, row, 1, 2 + n)
        row += 1

    risk_row("유동부채 > 유동자산", lambda i: f"=IF({ind_ref('유동부채', i)}>{ind_ref('유동자산', i)},\"⚠ 위험\",\"양호\")")
    risk_row("차입금 과다 (자기자본비율<20%)", lambda i: f"=IF({ind_ref('자기자본비율', i)}<20,\"⚠ 위험\",\"양호\")")
    risk_row("순자산 마이너스 (채무초과)", lambda i: f"=IF({ind_ref('자본총계', i)}<0,\"⚠ 위험\",\"양호\")")

    def receivable_spike(i):
        if i == 0:
            return ""
        rev_c, rev_p = ind_ref("매출액", i), ind_ref("매출액", i - 1)
        rec_c, rec_p = ind_ref("매출채권", i), ind_ref("매출채권", i - 1)
        if not (rev_c and rev_p and rec_c and rec_p):
            return ""
        return (f"=IFERROR(IF((({rec_c}-{rec_p})/{rec_p}*100)-(({rev_c}-{rev_p})/{rev_p}*100)>20,"
                f"\"⚠ 위험(매출채권 급증)\",\"양호\"),\"\")")

    risk_row("매출채권 급증 (매출 증가율 대비 +20%p 이상)", receivable_spike)
    risk_row("영업활동현금흐름 마이너스", lambda i: f"=IF({base_cell('영업활동현금흐름', i)}<0,\"⚠ 위험\",\"양호\")")
    risk_row("이자보상배율 1 미만 (이자도 못 갚는 수준)", lambda i: f"=IFERROR(IF({ind_ref('영업이익', i)}/{base_cell('이자비용', i)}<1,\"⚠ 위험\",\"양호\"),\"\")")
    row += 1

    # --- D. 청산가치 (최신 연도, 적용비율은 셀 참조 — 구현 불변 규칙) ---
    ws.cell(row=row, column=1, value="D. 청산가치 (자산가치주 체크 · 최신 연도 기준)").font = LABEL
    row += 1
    last_i = n - 1
    ws.cell(row=row, column=1, value="항목")
    ws.cell(row=row, column=2, value="적용비율")
    ws.cell(row=row, column=3, value="조정가치")
    style_header(ws, row, 1, 3)
    d_header = row
    row += 1
    haircuts = [("현금및현금성자산", "현금및현금성자산", 1.00), ("단기투자(유가증권)", "단기투자", 1.00),
                ("매출채권", None, 0.85), ("재고자산", "재고자산", 0.50), ("장기투자자산", "장기투자", 0.50),
                ("유형자산", "유형자산", 0.50), ("무형자산", "무형자산", 0.00), ("영업권", "영업권", 0.00),
                ("기타유동자산", "기타유동자산", 0.00)]
    adj_rows = []
    for label, base_key, pct in haircuts:
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=pct).number_format = "0%"
        src = ind_ref("매출채권", last_i) if base_key is None else base_cell(base_key, last_i)
        if src:
            ws.cell(row=row, column=3, value=f"=IFERROR({src}*B{row},0)").number_format = AMT_FMT
        adj_rows.append(row)
        row += 1
    sum_row = row
    ws.cell(row=row, column=1, value="조정자산 합계").font = LABEL
    ws.cell(row=row, column=3, value=f"=SUM(C{adj_rows[0]}:C{adj_rows[-1]})").number_format = AMT_FMT
    row += 1
    ws.cell(row=row, column=1, value="총부채(부채총계)").font = LABEL
    ws.cell(row=row, column=3, value=f"={ind_ref('부채총계', last_i)}").number_format = AMT_FMT
    row += 1
    ws.cell(row=row, column=1, value="청산가치 (조정자산 − 총부채)").font = LABEL
    ws.cell(row=row, column=3, value=f"=C{sum_row}-C{row - 1}").number_format = AMT_FMT
    apply_border(ws, d_header, row, 1, 3)
    row += 2

    # --- E. CCC ---
    ws.cell(row=row, column=1, value="E. 현금전환주기 (CCC)").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("매출채권회수기간(일)", lambda i: f"=IFERROR(365/{b_ref('매출채권회전율(회)', i)},NA())")
    write_ratio_row("재고자산처리기간(일)", lambda i: f"=IFERROR(365/{b_ref('재고자산회전율(회)', i)},NA())")
    write_ratio_row("매입채무지불기간(일)", lambda i: f"=IFERROR(365/{b_ref('매입채무회전율(회)', i)},NA())")
    write_ratio_row("CCC = 회수+처리-지불(일)", lambda i: f"=IFERROR({b_ref('매출채권회수기간(일)', i)}+{b_ref('재고자산처리기간(일)', i)}-{b_ref('매입채무지불기간(일)', i)},NA())")
    ws.cell(row=row, column=1, value="※ CCC가 짧을수록(마이너스에 가까울수록) 운전자본 부담이 적은 우량한 구조입니다.").font = NOTE
    row += 2
    add_section_chart("현금전환주기 (CCC)", ["매출채권회수기간(일)", "재고자산처리기간(일)", "매입채무지불기간(일)", "CCC = 회수+처리-지불(일)"], primary_ytitle="일")

    # --- F. FCF (라벨은 build_valuation_sheet.py가 참조하므로 KR과 동일 유지) ---
    ws.cell(row=row, column=1, value="F. 잉여현금흐름 (FCF)").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("FCF = 영업활동현금흐름 - CAPEX", lambda i: f"=IFERROR({base_cell('영업활동현금흐름', i)}-ABS({base_cell('유형자산의취득', i)}),NA())", fmt=AMT_FMT)
    write_ratio_row("FCF마진(%, FCF/매출액)", lambda i: f"=IFERROR({b_ref('FCF = 영업활동현금흐름 - CAPEX', i)}/{ind_ref('매출액', i)}*100,NA())")
    row += 1
    add_section_chart("잉여현금흐름 (FCF)", ["FCF = 영업활동현금흐름 - CAPEX"], ["FCF마진(%, FCF/매출액)"], primary_ytitle="백만달러", secondary_ytitle="%")

    # --- G. 현금흐름 3단 구분 ---
    ws.cell(row=row, column=1, value="G. 현금흐름 3단 구분").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("영업활동현금흐름", lambda i: f"={base_cell('영업활동현금흐름', i)}", fmt=AMT_FMT)
    write_ratio_row("투자활동현금흐름", lambda i: f"={base_cell('투자활동현금흐름', i)}" if base_has.get("투자활동현금흐름") else "", fmt=AMT_FMT)
    write_ratio_row("재무활동현금흐름", lambda i: f"={base_cell('재무활동현금흐름', i)}" if base_has.get("재무활동현금흐름") else "", fmt=AMT_FMT)
    write_ratio_row("현금 순증감 (3단 합계, 검증용)", lambda i: f"=IFERROR({b_ref('영업활동현금흐름', i)}+{b_ref('투자활동현금흐름', i)}+{b_ref('재무활동현금흐름', i)},NA())", fmt=AMT_FMT)
    ws.cell(row=row, column=1, value="※ 검증용 합계는 지표 시트의 '현금및현금성자산의증가'와 대체로 비슷해야 합니다(환율 변동 등으로 소폭 차이 가능).").font = NOTE
    row += 2
    add_section_chart("현금흐름 3단 구분", ["영업활동현금흐름", "투자활동현금흐름", "재무활동현금흐름", "현금 순증감 (3단 합계, 검증용)"], primary_ytitle="백만달러")

    # --- H. DuPont ---
    ws.cell(row=row, column=1, value="H. DuPont 분해 (ROE = 순이익률 × 총자산회전율 × 레버리지)").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("레버리지 (자산/자기자본, 배)", lambda i: f"=IFERROR({ind_ref('자산총계', i)}/{ind_ref('자본총계', i)},NA())", fmt="0.00")
    write_ratio_row("ROE 검증 (순이익률×회전율×레버리지, %)", lambda i: f"=IFERROR({b_ref('순이익률(%)', i)}*{b_ref('총자산회전율(회)', i)}*{b_ref('레버리지 (자산/자기자본, 배)', i)},NA())")
    ws.cell(row=row, column=1, value="※ 위 B섹션의 ROE(%)와 거의 같아야 정상입니다.").font = NOTE
    row += 2
    add_section_chart("DuPont 분해 (ROE)", ["레버리지 (자산/자기자본, 배)"], ["ROE 검증 (순이익률×회전율×레버리지, %)"], primary_ytitle="배", secondary_ytitle="%")

    # --- I. ROIC / NOPLAT (간이) — 실효세율은 세전이익≤0이면 NA, 0~50% 클램프(구현 불변 규칙) ---
    ws.cell(row=row, column=1, value="I. ROIC / NOPLAT (간이 계산)").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("실효세율(%)", lambda i: (
        f"=IFERROR(IF({base_cell('세전이익', i)}<=0,NA(),MIN(MAX({base_cell('법인세비용', i)}/{base_cell('세전이익', i)}*100,0),50)),NA())"))
    write_ratio_row("NOPLAT = 영업이익×(1-실효세율)", lambda i: f"=IFERROR({ind_ref('영업이익', i)}*(1-{b_ref('실효세율(%)', i)}/100),NA())", fmt=AMT_FMT)
    write_ratio_row("투하자본(간이) = 자기자본+비유동부채", lambda i: f"=IFERROR({ind_ref('자본총계', i)}+{base_cell('비유동부채', i)},NA())", fmt=AMT_FMT)
    write_ratio_row("ROIC(간이, %)", lambda i: f"=IFERROR({b_ref('NOPLAT = 영업이익×(1-실효세율)', i)}/{b_ref('투하자본(간이) = 자기자본+비유동부채', i)}*100,NA())")
    ws.cell(row=row, column=1, value="※ 간이 버전입니다. 정교한 ROIC은 이자부부채만 골라 투하자본을 계산해야 합니다.").font = NOTE
    row += 2
    add_section_chart("ROIC / NOPLAT (간이)", ["NOPLAT = 영업이익×(1-실효세율)", "투하자본(간이) = 자기자본+비유동부채"], ["실효세율(%)", "ROIC(간이, %)"], primary_ytitle="백만달러", secondary_ytitle="%")

    # --- J. 자산·부채 구성비 ---
    ws.cell(row=row, column=1, value="J. 자산·부채 구성비 변화 (간이)").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("유동자산 비중(%)", lambda i: f"=IFERROR({ind_ref('유동자산', i)}/{ind_ref('자산총계', i)}*100,NA())")
    write_ratio_row("비유동자산 비중(%)", lambda i: f"=IFERROR({base_cell('비유동자산', i)}/{ind_ref('자산총계', i)}*100,NA())")
    write_ratio_row("유동부채 비중(%, 총부채 대비)", lambda i: f"=IFERROR({ind_ref('유동부채', i)}/{ind_ref('부채총계', i)}*100,NA())")
    write_ratio_row("비유동부채 비중(%, 총부채 대비)", lambda i: f"=IFERROR({base_cell('비유동부채', i)}/{ind_ref('부채총계', i)}*100,NA())")
    row += 1
    add_section_chart("자산·부채 구성비 변화", ["유동자산 비중(%)", "비유동자산 비중(%)", "유동부채 비중(%, 총부채 대비)", "비유동부채 비중(%, 총부채 대비)"], primary_ytitle="%")

    # --- K. 외환손익 ---
    ws.cell(row=row, column=1, value="K. 외환손익").font = LABEL
    row += 1
    write_period_header()
    write_ratio_row("외화환산손익", lambda i: f"={base_cell('외화환산손익', i)}" if base_has.get("외화환산손익") else "", fmt=AMT_FMT)
    if not base_has.get("외화환산손익"):
        ws.cell(row=row, column=1, value="※ 이 회사 공시에서 외화환산손익 태그를 찾지 못했습니다(달러 결제 위주 회사에는 없는 게 정상 — 매칭 실패 경고에는 포함하지 않았습니다).").font = NOTE
        row += 1
    row += 1

    # --- L. 주가 연동 지표 (연도별, KR과 동일 레이아웃) ---
    ws.cell(row=row, column=1, value="L. 주가 연동 지표 (회계연도말 종가 기준)").font = LABEL
    row += 1
    price_by_i: list[float | None] = []
    date_by_i: list[str] = []
    for p_end in a_periods:
        ym = p_end[:7]
        px = monthly.get(ym)
        if px is None:  # 회계연도말이 월중이면 직전 달로 폴백
            y, m = int(ym[:4]), int(ym[5:7])
            m2, y2 = (m - 1, y) if m > 1 else (12, y - 1)
            px = monthly.get(f"{y2:04d}-{m2:02d}")
        price_by_i.append(round(float(px), 2) if px is not None else None)
        date_by_i.append(ym if px is not None else "")
    price_available = any(v is not None for v in price_by_i)
    write_period_header()
    write_ratio_row("기준일(종가)", lambda i: date_by_i[i], fmt="General")
    write_ratio_row("종가", lambda i: price_by_i[i], fmt="#,##0.00")
    shares_m = (shares / UNIT_DIVISOR) if shares else None
    write_ratio_row("시가총액", lambda i: (f"={b_ref('종가', i)}*{shares_m}" if (shares_m and price_by_i[i] is not None) else ""), fmt="#,##0.0")
    write_ratio_row("PER(배)", lambda i: (f"=IFERROR({b_ref('시가총액', i)}/{ind_ref('당기순이익', i)},NA())" if (shares_m and price_by_i[i] is not None) else ""), fmt="0.0")
    write_ratio_row("PBR(배)", lambda i: (f"=IFERROR({b_ref('시가총액', i)}/{ind_ref('자본총계', i)},NA())" if (shares_m and price_by_i[i] is not None) else ""), fmt="0.00")
    write_ratio_row("PSR(배)", lambda i: (f"=IFERROR({b_ref('시가총액', i)}/{ind_ref('매출액', i)},NA())" if (shares_m and price_by_i[i] is not None) else ""), fmt="0.00")
    ws.cell(row=row, column=1, value="※ 과거 시가총액은 '현재 상장주식수 × 당시 종가'의 근사치입니다(과거 주식수 미반영 — 자사주 매입이 큰 회사는 과거 PER가 실제보다 높게 보일 수 있음).").font = NOTE
    row += 1
    if not price_available:
        ws.cell(row=row, column=1, value="※ 월별 종가 이력이 없어 연도별 주가 지표가 비었습니다. fetch_extra_info.py를 다시 실행해 보세요(yfinance 차단 시 현재가만 존재).").font = WARN
        row += 1
    # 현재가 블록 (최근 종가와 별도)
    cur_price = info.get("currentPrice")
    ws.cell(row=row, column=1, value=f"현재가($){' · 기준 ' + str(info.get('priceDate')) if info.get('priceDate') else ''}")
    if cur_price is not None:
        ws.cell(row=row, column=3, value=cur_price).number_format = "#,##0.00"
    apply_border(ws, row, row, 1, 3)
    cur_price_row = row
    row += 1
    ws.cell(row=row, column=1, value="시가총액(백만달러, 현재)")
    if info.get("marketCap") is not None:
        ws.cell(row=row, column=3, value=info["marketCap"] / UNIT_DIVISOR).number_format = "#,##0.0"
    apply_border(ws, row, row, 1, 3)
    cur_mc_row = row
    row += 1
    for lab, den_key, fmt in [("현재가 기준 PER(배, 최신 회계연도 이익)", "당기순이익", "0.0"),
                              ("현재가 기준 PBR(배)", "자본총계", "0.00"),
                              ("현재가 기준 PSR(배)", "매출액", "0.00")]:
        ws.cell(row=row, column=1, value=lab)
        if info.get("marketCap") is not None:
            ws.cell(row=row, column=3, value=f"=IFERROR(C{cur_mc_row}/{ind_ref(den_key, last_i)},NA())").number_format = fmt
        apply_border(ws, row, row, 1, 3)
        row += 1
    row += 1
    if price_available:
        add_section_chart("주가 연동 지표 (PER/PBR/PSR)", ["PER(배)", "PBR(배)", "PSR(배)"], primary_ytitle="배", primary_type="line")

    # --- M. 배당 ---
    ws.cell(row=row, column=1, value="M. 배당").font = LABEL
    row += 1
    dy, pr, dr = info.get("dividendYield"), info.get("payoutRatio"), info.get("dividendRate")
    ws.cell(row=row, column=1, value="배당수익률(%)")
    if dy is not None:
        ws.cell(row=row, column=2, value=(dy if dy > 1 else dy * 100)).number_format = "0.00"
    row += 1
    ws.cell(row=row, column=1, value="배당성향(%)")
    if pr is not None:
        ws.cell(row=row, column=2, value=pr * 100).number_format = "0.0"
    row += 1
    ws.cell(row=row, column=1, value="주당 현금배당금")
    last_div = None
    if div_by_year:
        last_y = sorted(div_by_year)[-1]
        last_div = div_by_year[last_y]
        ws.cell(row=row, column=2, value=last_div).number_format = "0.00"
        ws.cell(row=row, column=3, value=f"({last_y}년 합계, $)").font = NOTE
    elif dr is not None:
        ws.cell(row=row, column=2, value=dr).number_format = "0.00"
        ws.cell(row=row, column=3, value="(yfinance dividendRate, $)").font = NOTE
    row += 1
    ws.cell(row=row, column=1, value="※ 대주주·자기주식 현황은 SEC proxy(DEF 14A)/13F 영역이라 이 파이프라인 범위 밖입니다(DART의 M섹션과 다른 부분).").font = NOTE
    row += 2

    # --- N. 투자 판단 (자동 평가 · A~E) — KR과 동일 규칙 ---
    ws.cell(row=row, column=1, value="N. 투자 판단 (자동 평가 · A~E)").font = LABEL
    row += 1
    ws.cell(row=row, column=1, value="항목")
    ws.cell(row=row, column=2, value="평가(A~E)")
    ws.cell(row=row, column=3, value="근거 메모")
    style_header(ws, row, 1, 3)
    row += 1

    def sv(key):
        return series_vals(key)

    def ratio_l(nums, dens, mult=100.0):
        return [a / b * mult if (a is not None and b not in (None, 0)) else None for a, b in zip(nums, dens)]

    def yoy_l(vals):
        out = [None]
        for prev, cur in zip(vals, vals[1:]):
            out.append((cur - prev) / prev * 100 if (prev not in (None, 0) and cur is not None) else None)
        return out

    s_rev, s_op, s_ni = sv("매출액"), sv("영업이익"), sv("당기순이익")
    s_assets, s_eq, s_liab = sv("자산총계"), sv("자본총계"), sv("부채총계")
    s_ca, s_cl = sv("유동자산"), sv("유동부채")
    raw = {
        "자기자본비율": ratio_l(s_eq, s_assets), "부채비율": ratio_l(s_liab, s_eq), "유동비율": ratio_l(s_ca, s_cl),
        "영업이익률": ratio_l(s_op, s_rev), "ROE": ratio_l(s_ni, s_eq), "ROA": ratio_l(s_ni, s_assets),
        "매출성장률": yoy_l(s_rev), "영업이익성장률": yoy_l(s_op), "총자산회전율": ratio_l(s_rev, s_assets, 1.0),
    }
    per_list, pbr_list = [], []
    for i in range(n):
        px = price_by_i[i]
        mc = px * shares / UNIT_DIVISOR if (px is not None and shares) else None
        ni, eq = s_ni[i], s_eq[i]
        # 순이익·자본이 0 이하이면 음수 PER/PBR이 '저평가 충족'으로 오판되므로 판정 제외(KR과 동일)
        per_list.append(mc / ni if (mc is not None and ni is not None and ni > 0) else None)
        pbr_list.append(mc / eq if (mc is not None and eq is not None and eq > 0) else None)

    def fmt_years(k, total):
        return f"{total}년 중 {k}년 충족"

    results = []
    g1, k1, t1, i1 = evaluate_metric(raw["자기자본비율"], 40, True)
    g2, k2, t2, i2 = evaluate_metric(raw["부채비율"], 200, False)
    g3, k3, t3, i3 = evaluate_metric(raw["유동비율"], 100, True)
    results.append(("재무건전성", combine_grades([g1, g2, g3]),
                    f"자기자본비율 40%↑ {fmt_years(k1, t1)}({g1}{', 지속개선' if i1 else ''}) / 부채비율 200%↓ {fmt_years(k2, t2)}({g2}{', 지속개선' if i2 else ''}) / 유동비율 100%↑ {fmt_years(k3, t3)}({g3}{', 지속개선' if i3 else ''})"))
    g4, k4, t4, i4 = evaluate_metric(raw["영업이익률"], 5, True)
    g5, k5, t5, i5 = evaluate_metric(raw["ROE"], 10, True)
    g6, k6, t6, i6 = evaluate_metric(raw["ROA"], 5, True)
    results.append(("수익성", combine_grades([g4, g5, g6]),
                    f"영업이익률 5%↑ {fmt_years(k4, t4)}({g4}{', 지속개선' if i4 else ''}) / ROE 10%↑ {fmt_years(k5, t5)}({g5}{', 지속개선' if i5 else ''}) / ROA 5%↑ {fmt_years(k6, t6)}({g6}{', 지속개선' if i6 else ''})"))
    g7, k7, t7, i7 = evaluate_metric(raw["매출성장률"], 10, True)
    g8, k8, t8, i8 = evaluate_metric(raw["영업이익성장률"], 0, True)
    g9, k9, t9, i9 = evaluate_metric(raw["총자산회전율"], 1.0, True)
    results.append(("성장성", combine_grades([g7, g8, g9]),
                    f"매출성장률 10%↑ {fmt_years(k7, t7)}({g7}{', 지속개선' if i7 else ''}) / 영업이익성장률 + {fmt_years(k8, t8)}({g8}{', 지속개선' if i8 else ''}) / 총자산회전율 1.0↑ {fmt_years(k9, t9)}({g9}{', 지속개선' if i9 else ''})"))
    if any(v is not None for v in pbr_list):
        gA, kA, tA, iA = evaluate_metric(pbr_list, 1.0, False)
        results.append(("자산으로 본 저평가 정도", gA, f"PBR 1.0 미만 {fmt_years(kA, tA)}({gA}{', 지속개선' if iA else ''}). 청산가치 대비 시가총액은 D·L 섹션 참고."))
    else:
        results.append(("자산으로 본 저평가 정도", "-", "주가 이력 없음 — fetch_extra_info.py 재실행 후 재생성하면 자동 평가됩니다."))
    if any(v is not None for v in per_list):
        gB, kB, tB, iB = evaluate_metric(per_list, 15.0, False)
        results.append(("수익 창출 능력으로 본 저평가 정도", gB, f"PER 15배 미만 {fmt_years(kB, tB)}({gB}{', 지속개선' if iB else ''}). PSR은 L섹션 참고."))
    else:
        results.append(("수익 창출 능력으로 본 저평가 정도", "-", "주가 이력 없음 — fetch_extra_info.py 재실행 후 재생성하면 자동 평가됩니다."))
    results.append(("사업역량", "(직접 입력)", "사업 단순성·거래처 분산·경쟁력은 공시 수치로 판단할 수 없습니다. 10-K Item 1(Business)과 웹 리서치로 직접 확인하세요."))
    if pr is not None:
        grade_sh = "A" if pr * 100 >= 30 else "B" if pr * 100 >= 20 else "C" if pr > 0 else "D"
        results.append(("주주 중시 자세", grade_sh, f"배당성향 {pr * 100:.1f}% (yfinance). ※ 자사주 매입 규모는 재무활동현금흐름(G섹션)과 10-K를 함께 확인하세요(미국은 배당보다 바이백 비중이 큰 회사가 많음)."))
    else:
        results.append(("주주 중시 자세", "-", "배당 데이터 없음. 자사주 매입 중심 회사일 수 있으니 G섹션 재무활동현금흐름을 확인하세요."))
    for item, grade, memo in results:
        ws.cell(row=row, column=1, value=item)
        c = ws.cell(row=row, column=2, value=grade)
        if grade in ("A", "B"):
            c.font = Font(name=FONT_NAME, bold=True, color="1F7A1F")
        elif grade in ("D", "E"):
            c.font = Font(name=FONT_NAME, bold=True, color="C00000")
        ws.cell(row=row, column=3, value=memo)
        row += 1
    ws.cell(row=row, column=1, value="※ 등급 규칙: 최근 5개년 중 기준 충족 연수로 A(5년)~E(0~1년), 기준 미달이어도 5년 연속 개선이면 A로 승격(KR 파이프라인과 동일 규칙).").font = NOTE

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 16
    for i in range(max(n, 5)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14

    return warnings


def build_workbook(ticker: str, company_name: str, period: str, quarters: int, years: int, outdir: str) -> dict:
    cached = load_companyfacts_cache(ticker)
    if not cached:
        print(f"ERROR: {ticker} SEC CompanyFacts 캐시가 없습니다. fetch_financials.py를 먼저 실행하세요.", file=sys.stderr)
        sys.exit(1)
    raw_facts = cached["raw"]

    q_payload = q_periods = a_payload = a_periods = None
    if period in ("quarterly", "both"):
        q_payload, q_periods = build_frequency_payload(raw_facts, "quarterly", quarters)
    if period in ("annual", "both"):
        a_payload, a_periods = build_frequency_payload(raw_facts, "annual", years)
    price_data = load_price_cache(ticker)

    wb = Workbook()
    wb.remove(wb.active)
    cell_index = write_raw_sheet(wb, ticker, q_payload, a_payload, q_periods or [], a_periods or [])

    ia_warnings = []
    if q_periods:
        q_row_map, q_labels = build_statement_sheet(wb, "분기_재무제표", "quarterly", q_periods, cell_index)
        build_indicator_sheet(wb, "분기", "분기_재무제표", q_row_map, q_labels)
    if a_periods:
        a_row_map, a_labels = build_statement_sheet(wb, "연간_재무제표", "annual", a_periods, cell_index)
        ind_sheet, ind_row_of = build_indicator_sheet(wb, "연간", "연간_재무제표", a_row_map, a_labels)
        ia_warnings = build_investment_analysis_sheet(
            wb, ticker, company_name, a_payload, a_periods, a_labels, a_row_map, ind_sheet, ind_row_of, price_data)

    desired_order = ["분기_재무제표", "연간_재무제표", "지표_분기", "지표_연간", "투자분석", "원본데이터"]
    wb._sheets = [wb[nm] for nm in desired_order if nm in wb.sheetnames]

    today = dt.date.today().strftime("%Y%m%d")
    outdir_path = Path(outdir)
    outdir_path.mkdir(parents=True, exist_ok=True)
    suffix = {"quarterly": "_분기", "annual": "_연간", "both": ""}[period]
    filepath = outdir_path / f"{company_name}{suffix}_{today}.xlsx"
    wb.save(filepath)

    return {
        "saved": str(filepath), "ticker": ticker, "period": period,
        "quarters_filled": len(q_periods or []), "years_filled": len(a_periods or []),
        "missing_indicators": ia_warnings,
        "price_history_used": bool((price_data or {}).get("monthly_close")),
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="미국 상장기업 재무 엑셀 생성(SEC EDGAR 기반, KR 파이프라인과 동일 구성)")
    ap.add_argument("ticker")
    ap.add_argument("company_name")
    ap.add_argument("--period", choices=["quarterly", "annual", "both"], default="both")
    ap.add_argument("--quarters", type=int, default=12)
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--outdir", default="/mnt/user-data/outputs")
    args = ap.parse_args()
    print(json.dumps(build_workbook(args.ticker, args.company_name, args.period, args.quarters, args.years, args.outdir), ensure_ascii=False))


if __name__ == "__main__":
    main()
