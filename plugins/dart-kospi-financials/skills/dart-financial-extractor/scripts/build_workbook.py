"""
fetch_financials.py가 cache/ 에 쌓아 둔 원자료를 읽어
'분기_재무제표'(최근 12분기)와 '연간_재무제표'(최근 5개년) 2개 시트 엑셀을 만든다.

핵심 설계:
  - 재무상태표(BS)는 시점 데이터이므로 각 분기 말 보고서 값을 그대로 링크한다.
  - 손익계산서/포괄손익계산서/현금흐름표(IS/CIS/CF)는 흐름 데이터이므로
    2분기·4분기는 반드시 엑셀 수식으로 계산한다(하드코딩 금지):
        2분기 = 반기보고서 thstrm_amount - 1분기보고서 thstrm_amount
        4분기 = 사업보고서 thstrm_amount - 3분기보고서 thstrm_add_amount(9개월 누적)
    1분기·3분기는 원본 값을 그대로 링크한다.
  - 모든 원본 값은 '원본데이터' 시트에 먼저 적재한 뒤, 화면에 보이는 시트의 셀은
    그 원본데이터 시트를 참조하는 수식으로 채운다. → 감사(audit) 가능, 값 변경 시 자동 재계산.

사용법:
    python build_workbook.py <corp_code> <company_name> [--years 5] [--quarters 12] [--outdir /mnt/user-data/outputs]

전제:
    같은 corp_code에 대해 fetch_financials.py를 먼저 필요한 (연도, reprt_code) 조합만큼
    실행해서 cache/ 에 JSON이 쌓여 있어야 한다.
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

SJ_ORDER = [("BS", "재무상태표"), ("IS", "손익계산서"), ("CIS", "포괄손익계산서"), ("CF", "현금흐름표")]
FLOW_TYPES = {"IS", "CIS", "CF"}  # 누적/차감 로직이 필요한 유형
FONT_NAME = "Arial"

BLUE = Font(name=FONT_NAME, color="0000FF")  # 하드코딩 입력값
BLACK = Font(name=FONT_NAME, color="000000")  # 수식
GREEN = Font(name=FONT_NAME, color="008000")  # 다른 시트 링크
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
BOLD = Font(name=FONT_NAME, bold=True)


def load_cache(corp_code: str) -> dict:
    """cache/{corp_code}_{year}_{reprt}_{fs_div}.json 전부를 로드해
    reports[year][reprt_code] = data 형태로 반환한다."""
    reports: dict[str, dict[str, dict]] = {}
    for fp in CACHE_DIR.glob(f"{corp_code}_*.json"):
        data = json.loads(fp.read_text(encoding="utf-8"))
        if data.get("status") != "000":
            continue
        year = data["bsns_year"]
        reprt = data["reprt_code"]
        reports.setdefault(year, {})[reprt] = data
    return reports


def build_quarter_plan(reports: dict, n_quarters: int) -> list[dict]:
    """사용 가능한 데이터로부터 최근 n_quarters개 분기 계획을 오래된 순으로 만든다.
    각 원소: {year, q, needs: {1Q:[year,'11013'], H1:[year,'11012'], 3Q:[year,'11014'], FY:[year,'11011']}}
    """
    years_desc = sorted(reports.keys(), reverse=True)
    plan = []
    for year in years_desc:
        y_reports = reports[year]
        # 4분기 -> 1분기 역순으로 채운다
        if "11011" in y_reports and "11014" in y_reports:
            plan.append({"year": year, "q": 4, "fy": y_reports["11011"], "q3": y_reports["11014"]})
        if "11014" in y_reports:
            plan.append({"year": year, "q": 3, "q3": y_reports["11014"]})
        if "11012" in y_reports and "11013" in y_reports:
            plan.append({"year": year, "q": 2, "h1": y_reports["11012"], "q1": y_reports["11013"]})
        if "11013" in y_reports:
            plan.append({"year": year, "q": 1, "q1": y_reports["11013"]})
        if len(plan) >= n_quarters:
            break
    plan.sort(key=lambda p: (p["year"], p["q"]))
    return plan[-n_quarters:]


def collect_accounts(periods: list[dict], sj_div: str, period_key_fn) -> list[dict]:
    """여러 기간의 데이터에서 계정과목 유니온을 ord 순으로 만든다.
    반환: [{account_id, account_nm}] ord 오름차순, 최근 기간 기준 우선."""
    seen = {}
    order_hint = {}
    for p in reversed(periods):  # 최근 기간을 우선 기준으로
        data = period_key_fn(p)
        if not data:
            continue
        for item in data.get("items", {}).get(sj_div, []):
            aid = item["account_id"]
            if aid not in seen:
                seen[aid] = item["account_nm"]
                order_hint[aid] = int(item.get("ord") or 9999)
    ordered = sorted(seen.keys(), key=lambda a: order_hint[a])
    return [{"account_id": a, "account_nm": seen[a]} for a in ordered]


def amount_lookup(data: dict, sj_div: str, account_id: str, field: str = "thstrm_amount"):
    if not data:
        return None
    for item in data.get("items", {}).get(sj_div, []):
        if item["account_id"] == account_id:
            val = item.get(field)
            if val in (None, ""):
                return None
            try:
                return float(str(val).replace(",", ""))
            except ValueError:
                return None
    return None


def write_raw_sheet(wb: Workbook, quarter_plan: list[dict], year_list: list[str], reports: dict, sheet_name: str = "원본데이터"):
    """모든 원자료를 원본 시트에 적재하고, 셀 좌표 인덱스를 반환한다."""
    ws = wb.create_sheet(sheet_name)
    ws.sheet_state = "hidden"
    ws["A1"] = "이 시트는 DART 원본 응답값을 그대로 담은 참조용 데이터입니다. 직접 수정하지 마세요."
    ws["A1"].font = Font(name=FONT_NAME, italic=True, size=9)

    row = 3
    cell_index = {}  # (sj_div, account_id, period_label, field) -> "'{sheet_name}'!$X$Y"

    def dump_period(period_data: dict, label: str):
        nonlocal row
        if not period_data:
            return
        ws.cell(row=row, column=1, value=label).font = BOLD
        row += 1
        for sj_div, sj_name in SJ_ORDER:
            items = period_data.get("items", {}).get(sj_div, [])
            if not items:
                continue
            ws.cell(row=row, column=1, value=sj_name).font = Font(name=FONT_NAME, italic=True)
            row += 1
            for item in items:
                ws.cell(row=row, column=1, value=item["account_id"])
                ws.cell(row=row, column=2, value=item["account_nm"])
                amt = item.get("thstrm_amount")
                ws.cell(row=row, column=3, value=float(str(amt).replace(",", "")) if amt not in (None, "") else None).font = BLUE
                cell_index[(sj_div, item["account_id"], label, "thstrm_amount")] = f"'{sheet_name}'!${get_column_letter(3)}${row}"
                add_amt = item.get("thstrm_add_amount")
                if add_amt not in (None, ""):
                    ws.cell(row=row, column=4, value=float(str(add_amt).replace(",", ""))).font = BLUE
                    cell_index[(sj_div, item["account_id"], label, "thstrm_add_amount")] = f"'{sheet_name}'!${get_column_letter(4)}${row}"
                row += 1
            row += 1
        row += 1

    # 분기용 원자료
    for p in quarter_plan:
        for key, rep_field in (("q1", "11013"), ("h1", "11012"), ("q3", "11014"), ("fy", "11011")):
            if key in p:
                dump_period(p[key], f"{p['year']}_{key}")

    # 연간용 원자료 (사업보고서, 분기용과 중복될 수 있으나 라벨을 분리해 독립적으로 유지)
    for year in year_list:
        fy_data = reports.get(year, {}).get("11011")
        dump_period(fy_data, f"{year}_사업보고서")

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    return cell_index


def style_header(ws, row: int, ncols: int):
    for col in range(1, ncols + 1):
        c = ws.cell(row=row, column=col)
        c.font = BOLD
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")


def build_quarterly_sheet(wb: Workbook, quarter_plan: list[dict], cell_index: dict, sheet_name: str = "분기_재무제표"):
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = "단위: 원 | 음영 셀은 원본데이터 시트 링크 또는 수식으로 자동 계산됩니다."
    ws["A1"].font = Font(name=FONT_NAME, italic=True, size=9)

    header_row = 3
    ws.cell(row=header_row, column=1, value="구분")
    ws.cell(row=header_row, column=2, value="계정과목")
    for i, p in enumerate(quarter_plan):
        ws.cell(row=header_row, column=3 + i, value=f"{p['year']}Q{p['q']}")
    style_header(ws, header_row, 2 + len(quarter_plan))

    account_row_map: dict[tuple[str, str], int] = {}
    account_name_map: dict[tuple[str, str], str] = {}

    row = header_row + 1
    for sj_div, sj_name in SJ_ORDER:
        accounts = collect_accounts(
            quarter_plan, sj_div,
            lambda p: p.get("fy") or p.get("q3") or p.get("h1") or p.get("q1"),
        )
        if not accounts:
            continue
        ws.cell(row=row, column=1, value=sj_name).font = BOLD
        row += 1
        for acc in accounts:
            account_row_map[(sj_div, acc["account_id"])] = row
            account_name_map[(sj_div, acc["account_id"])] = acc["account_nm"]
            ws.cell(row=row, column=2, value=acc["account_nm"])
            for i, p in enumerate(quarter_plan):
                col = 3 + i
                cell = ws.cell(row=row, column=col)
                q = p["q"]
                label = f"{p['year']}_"
                if sj_div == "BS":
                    # 시점 데이터: 해당 분기말 보고서를 그대로 링크
                    src_key = "fy" if q == 4 else "q3" if q == 3 else "h1" if q == 2 else "q1"
                    ref = cell_index.get((sj_div, acc["account_id"], f"{label}{src_key}", "thstrm_amount"))
                    if ref:
                        cell.value = f"={ref}"
                        cell.font = GREEN
                elif q == 1:
                    ref = cell_index.get((sj_div, acc["account_id"], f"{label}q1", "thstrm_amount"))
                    if ref:
                        cell.value = f"={ref}"
                        cell.font = GREEN
                elif q == 3:
                    ref = cell_index.get((sj_div, acc["account_id"], f"{label}q3", "thstrm_amount"))
                    if ref:
                        cell.value = f"={ref}"
                        cell.font = GREEN
                elif q == 2:
                    h1_ref = cell_index.get((sj_div, acc["account_id"], f"{label}h1", "thstrm_amount"))
                    q1_ref = cell_index.get((sj_div, acc["account_id"], f"{label}q1", "thstrm_amount"))
                    if h1_ref and q1_ref:
                        cell.value = f"={h1_ref}-{q1_ref}"
                        cell.font = BLACK
                elif q == 4:
                    fy_ref = cell_index.get((sj_div, acc["account_id"], f"{label}fy", "thstrm_amount"))
                    # 4분기는 반드시 3분기보고서의 '누적' 필드를 사용한다 (thstrm_add_amount)
                    q3_add_ref = cell_index.get((sj_div, acc["account_id"], f"{label}q3", "thstrm_add_amount"))
                    if fy_ref and q3_add_ref:
                        cell.value = f"={fy_ref}-{q3_add_ref}"
                        cell.font = BLACK
                cell.number_format = "#,##0;(#,##0);-"
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 32
    for i in range(len(quarter_plan)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 15
    ws.freeze_panes = "C4"

    period_labels = [f"{p['year']}Q{p['q']}" for p in quarter_plan]
    return account_row_map, account_name_map, period_labels


def build_annual_sheet(wb: Workbook, year_list: list[str], reports: dict, cell_index: dict, sheet_name: str = "연간_재무제표"):
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = "단위: 원 | 사업보고서(연결/별도) 기준, 원본데이터 시트 링크"
    ws["A1"].font = Font(name=FONT_NAME, italic=True, size=9)

    header_row = 3
    ws.cell(row=header_row, column=1, value="구분")
    ws.cell(row=header_row, column=2, value="계정과목")
    for i, year in enumerate(year_list):
        ws.cell(row=header_row, column=3 + i, value=f"{year}(사업보고서)")
    style_header(ws, header_row, 2 + len(year_list))

    account_row_map: dict[tuple[str, str], int] = {}
    account_name_map: dict[tuple[str, str], str] = {}

    row = header_row + 1
    fy_periods = [{"fy": reports.get(y, {}).get("11011")} for y in year_list]
    for sj_div, sj_name in SJ_ORDER:
        accounts = collect_accounts(fy_periods, sj_div, lambda p: p.get("fy"))
        if not accounts:
            continue
        ws.cell(row=row, column=1, value=sj_name).font = BOLD
        row += 1
        for acc in accounts:
            account_row_map[(sj_div, acc["account_id"])] = row
            account_name_map[(sj_div, acc["account_id"])] = acc["account_nm"]
            ws.cell(row=row, column=2, value=acc["account_nm"])
            for i, year in enumerate(year_list):
                col = 3 + i
                cell = ws.cell(row=row, column=col)
                ref = cell_index.get((sj_div, acc["account_id"], f"{year}_사업보고서", "thstrm_amount"))
                if ref:
                    cell.value = f"={ref}"
                    cell.font = GREEN
                cell.number_format = "#,##0;(#,##0);-"
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 32
    for i in range(len(year_list)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 18
    ws.freeze_panes = "C4"

    period_labels = [f"{year}(사업보고서)" for year in year_list]
    return account_row_map, account_name_map, period_labels


# ---------------------------------------------------------------------------
# 지표(추세) 시트 & 차트 시트
#
# DART 계정명(account_nm)은 회사마다 표기가 조금씩 다를 수 있어(예: "매출액" vs
# "수익(매출액)", "기타수익" vs "기타영업외수익") 계정ID 하나로 고정 매칭하지 않고
# 우선순위가 있는 후보 이름 목록 + 부분일치 폴백으로 탐색한다. 못 찾으면 해당
# 지표는 빈 칸으로 두고 경고를 출력한다(값을 임의로 지어내지 않음).
# ---------------------------------------------------------------------------

# key: 내부 식별자, value: (sj_div, [완전일치 후보(우선순위순)], [부분일치 폴백 키워드], [제외 키워드])
METRIC_RULES: dict[str, tuple[str, list[str], list[str], list[str]]] = {
    "매출액": ("IS", ["매출액", "수익(매출액)", "영업수익"], ["매출액"], ["매출원가", "율"]),
    "매출원가": ("IS", ["매출원가"], ["매출원가"], ["율"]),
    "영업이익": ("IS", ["영업이익", "영업이익(손실)"], ["영업이익"], ["율", "률"]),
    "당기순이익": (
        "IS",
        ["당기순이익", "당기순이익(손실)", "분기순이익(손실)", "반기순이익(손실)", "분기순이익", "반기순이익"],
        ["순이익"],
        ["지배", "비지배", "주당"],
    ),
    "유동자산": ("BS", ["유동자산"], ["유동자산"], ["비유동"]),
    "유동부채": ("BS", ["유동부채"], ["유동부채"], ["비유동"]),
    "자산총계": ("BS", ["자산총계"], ["자산총계"], []),
    "부채총계": ("BS", ["부채총계"], ["부채총계"], []),
    "자본총계": ("BS", ["자본총계"], ["자본총계"], []),
    "매출채권및기타채권": (
        "BS",
        ["매출채권및기타채권", "매출채권및기타유동채권", "매출채권"],
        ["매출채권"],
        ["비유동"],
    ),
    "이익잉여금": ("BS", ["이익잉여금", "이익잉여금(결손금)"], ["이익잉여금"], []),
    "현금및현금성자산의증가": (
        "CF",
        ["현금및현금성자산의순증가(감소)", "현금및현금성자산의 증가(감소)", "현금및현금성자산의증가(감소)"],
        ["현금및현금성자산의"],
        ["기초", "기말", "환율"],
    ),
    "금융수익": ("IS", ["금융수익"], ["금융수익"], []),
    "금융비용": ("IS", ["금융비용"], ["금융비용"], []),
    "기타수익": ("IS", ["기타수익", "기타영업외수익"], ["기타수익", "기타영업외수익"], []),
    "기타비용": ("IS", ["기타비용", "기타영업외비용"], ["기타비용", "기타영업외비용"], []),
}


def resolve_metric(
    key: str, account_row_map: dict, account_name_map: dict
) -> tuple[str, str] | None:
    """METRIC_RULES에 따라 (sj_div, account_id)를 찾아 반환한다. 못 찾으면 None."""
    sj_div, exact_candidates, substr_keywords, excludes = METRIC_RULES[key]
    same_sj = [
        (sj, aid) for (sj, aid) in account_row_map if sj == sj_div
    ]

    def excluded(name: str) -> bool:
        return any(ex in name for ex in excludes)

    # 1차: 완전일치 (우선순위 순서대로)
    for cand in exact_candidates:
        for sj, aid in same_sj:
            name = account_name_map[(sj, aid)]
            if name == cand and not excluded(name):
                return sj, aid

    # 2차: 부분일치
    for kw in substr_keywords:
        for sj, aid in same_sj:
            name = account_name_map[(sj, aid)]
            if kw in name and not excluded(name):
                return sj, aid

    return None


# 지표 시트에 표시할 행 순서와 표시 이름 (라인차트 5개가 참조하는 "기본" 지표들)
INDICATOR_ROWS = [
    "매출액",
    "매출원가",
    "매출총이익",
    "영업이익",
    "당기순이익",
    "경상이익",
    "유동자산",
    "매출채권및기타채권",
    "유동부채",
    "자산총계",
    "부채총계",
    "자본총계",
    "이익잉여금",
    "현금및현금성자산의증가",
]

RATIO_ROWS = ["자기자본비율", "부채비율", "매출총이익률", "원가율", "영업이익률", "순이익률"]

# 그래프1~5 구성 (지표 시트의 행 이름 기준)
LINE_CHART_GROUPS = [
    ("그래프1_매출액-매출원가-매출총이익", ["매출액", "매출원가", "매출총이익"]),
    ("그래프2_이익지표(매출총이익-영업이익-순이익-경상이익)", ["매출총이익", "영업이익", "당기순이익", "경상이익"]),
    ("그래프3_유동자산-유동부채-자산-부채", ["유동자산", "유동부채", "자산총계", "부채총계"]),
    ("그래프4_매출액-매출채권", ["매출액", "매출채권및기타채권"]),
    ("그래프5_이익잉여금-현금증가", ["이익잉여금", "현금및현금성자산의증가"]),
]


def build_indicator_sheet(
    wb: Workbook,
    prefix: str,
    source_sheet_name: str,
    period_labels: list[str],
    account_row_map: dict,
    account_name_map: dict,
) -> tuple[str, dict[str, int], list[str]]:
    """소스 시트(분기_재무제표/연간_재무제표)를 참조하는 '지표_{prefix}' 시트를 만든다.
    반환: (시트이름, {행이름: 행번호}, 못찾은 지표 목록)"""
    sheet_name = f"지표_{prefix}"
    ws = wb.create_sheet(sheet_name)
    n = len(period_labels)

    ws.cell(row=1, column=1, value="지표").font = BOLD
    for i, label in enumerate(period_labels):
        ws.cell(row=1, column=3 + i, value=label)
    style_header(ws, 1, 2 + n)

    resolved: dict[str, tuple[str, str] | None] = {
        key: resolve_metric(key, account_row_map, account_name_map)
        for key in METRIC_RULES
    }
    missing = [key for key, v in resolved.items() if v is None]

    row_of: dict[str, int] = {}
    row = 2
    for name in INDICATOR_ROWS:
        ws.cell(row=row, column=2, value=name)
        row_of[name] = row

        if name == "매출총이익":
            if row_of.get("매출액") and row_of.get("매출원가"):
                a, b = row_of["매출액"], row_of["매출원가"]
                for i in range(n):
                    col = get_column_letter(3 + i)
                    ws.cell(row=row, column=3 + i, value=f"={col}{a}-{col}{b}")
        elif name == "경상이익":
            if row_of.get("영업이익"):
                op_row = row_of["영업이익"]
                fin_terms = []  # (부호, sj_div, account_id)
                for key, sign in (("금융수익", "+"), ("금융비용", "-"), ("기타수익", "+"), ("기타비용", "-")):
                    hit = resolved.get(key)
                    if hit:
                        fin_terms.append((sign, hit))
                for i in range(n):
                    col = get_column_letter(3 + i)
                    src_col = get_column_letter(3 + i)
                    parts = [f"{src_col}{op_row}"]
                    for sign, (sj, aid) in fin_terms:
                        r = account_row_map[(sj, aid)]
                        parts.append(f"{sign}'{source_sheet_name}'!{src_col}{r}")
                    ws.cell(row=row, column=3 + i, value="=" + "".join(parts))
        else:
            hit = resolved.get(name)
            if hit:
                sj, aid = hit
                src_row = account_row_map[(sj, aid)]
                for i in range(n):
                    col = get_column_letter(3 + i)
                    ref = f"'{source_sheet_name}'!{col}{src_row}"
                    # 해당 회사/기간에 실제 데이터가 없어 원본 셀이 비어 있으면 0이 아니라
                    # NA()를 반환한다 → 차트에서 0으로 떨어지지 않고 구간이 끊겨(gap) 표시됨.
                    ws.cell(row=row, column=3 + i, value=f'=IF({ref}="",NA(),{ref})')

        for i in range(n):
            ws.cell(row=row, column=3 + i).number_format = "#,##0;(#,##0);-"
        row += 1

    row += 1  # 구분 여백
    for name in RATIO_ROWS:
        ws.cell(row=row, column=2, value=name)
        row_of[name] = row
        if name == "자기자본비율" and row_of.get("자본총계") and row_of.get("자산총계"):
            num, den = row_of["자본총계"], row_of["자산총계"]
        elif name == "부채비율" and row_of.get("부채총계") and row_of.get("자본총계"):
            num, den = row_of["부채총계"], row_of["자본총계"]
        elif name == "매출총이익률" and row_of.get("매출총이익") and row_of.get("매출액"):
            num, den = row_of["매출총이익"], row_of["매출액"]
        elif name == "원가율" and row_of.get("매출원가") and row_of.get("매출액"):
            num, den = row_of["매출원가"], row_of["매출액"]
        elif name == "영업이익률" and row_of.get("영업이익") and row_of.get("매출액"):
            num, den = row_of["영업이익"], row_of["매출액"]
        elif name == "순이익률" and row_of.get("당기순이익") and row_of.get("매출액"):
            num, den = row_of["당기순이익"], row_of["매출액"]
        else:
            num, den = None, None
        if num and den:
            for i in range(n):
                col = get_column_letter(3 + i)
                ws.cell(row=row, column=3 + i, value=f"=IFERROR({col}{num}/{col}{den}*100,NA())")
                ws.cell(row=row, column=3 + i).number_format = "0.0"
        row += 1

    if missing:
        ws.cell(row=row + 1, column=1, value="※ 아래 지표는 이 회사 공시에서 계정명을 찾지 못해 비어 있습니다:").font = Font(
            name=FONT_NAME, italic=True, size=9, color="C00000"
        )
        ws.cell(row=row + 2, column=1, value=", ".join(missing)).font = Font(
            name=FONT_NAME, italic=True, size=9, color="C00000"
        )

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    for i in range(n):
        ws.column_dimensions[get_column_letter(3 + i)].width = 15
    ws.freeze_panes = "C2"

    return sheet_name, row_of, missing


def _add_line_series(chart: LineChart, ws, row_of: dict, names: list[str], n_periods: int):
    for name in names:
        r = row_of.get(name)
        if not r:
            continue
        data_ref = Reference(ws, min_col=3, max_col=2 + n_periods, min_row=r, max_row=r)
        series = Series(data_ref, title=name)
        series.smooth = False
        chart.series.append(series)


def build_chart_sheet(
    wb: Workbook, prefix: str, indicator_sheet_name: str, row_of: dict, n_periods: int
):
    ws = wb.create_sheet(f"차트_{prefix}")
    ind_ws = wb[indicator_sheet_name]
    cat_ref = Reference(ind_ws, min_col=3, max_col=2 + n_periods, min_row=1, max_row=1)

    anchor_row = 1
    for title, names in LINE_CHART_GROUPS:
        chart = LineChart()
        chart.title = title
        chart.style = 2
        chart.y_axis.title = "금액(원)"
        chart.x_axis.title = "기간"
        chart.height = 9
        chart.width = 22
        _add_line_series(chart, ind_ws, row_of, names, n_periods)
        chart.set_categories(cat_ref)
        ws.add_chart(chart, f"A{anchor_row}")
        anchor_row += 19

    # 비율 막대그래프 (계열 6개)
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.title = "그래프6_수익성-안정성 비율(%)"
    bar.style = 10
    bar.y_axis.title = "%"
    bar.x_axis.title = "기간"
    bar.height = 9
    bar.width = 22
    for name in RATIO_ROWS:
        r = row_of.get(name)
        if not r:
            continue
        data_ref = Reference(ind_ws, min_col=3, max_col=2 + n_periods, min_row=r, max_row=r)
        series = Series(data_ref, title=name)
        bar.series.append(series)
    bar.set_categories(cat_ref)
    ws.add_chart(bar, f"A{anchor_row}")

    ws.sheet_view.showGridLines = False


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("corp_code")
    ap.add_argument("company_name")
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--quarters", type=int, default=12)
    ap.add_argument("--outdir", default="/mnt/user-data/outputs")
    args = ap.parse_args()

    reports = load_cache(args.corp_code)
    if not reports:
        print(f"ERROR: {args.corp_code}에 대한 캐시 데이터가 없습니다. fetch_financials.py를 먼저 실행하세요.", file=sys.stderr)
        sys.exit(1)

    quarter_plan = build_quarter_plan(reports, args.quarters)
    year_list = sorted([y for y in reports if "11011" in reports[y]], reverse=True)[: args.years]
    year_list.sort()  # 오래된 -> 최근

    if len(quarter_plan) < args.quarters:
        print(
            f"WARNING: 요청한 {args.quarters}분기 중 {len(quarter_plan)}분기만 채웠습니다 "
            f"(공시 지연 또는 상장 이력 부족 가능).",
            file=sys.stderr,
        )
    if len(year_list) < args.years:
        print(
            f"WARNING: 요청한 {args.years}개년 중 {len(year_list)}개년만 채웠습니다.",
            file=sys.stderr,
        )

    wb = Workbook()
    wb.remove(wb.active)

    cell_index = write_raw_sheet(wb, quarter_plan, year_list, reports)
    q_row_map, q_name_map, q_labels = build_quarterly_sheet(wb, quarter_plan, cell_index)
    y_row_map, y_name_map, y_labels = build_annual_sheet(wb, year_list, reports, cell_index)

    all_missing: dict[str, list[str]] = {}
    if quarter_plan:
        q_ind_sheet, q_row_of, q_missing = build_indicator_sheet(
            wb, "분기", "분기_재무제표", q_labels, q_row_map, q_name_map
        )
        build_chart_sheet(wb, "분기", q_ind_sheet, q_row_of, len(q_labels))
        if q_missing:
            all_missing["분기"] = q_missing
    if year_list:
        y_ind_sheet, y_row_of, y_missing = build_indicator_sheet(
            wb, "연간", "연간_재무제표", y_labels, y_row_map, y_name_map
        )
        build_chart_sheet(wb, "연간", y_ind_sheet, y_row_of, len(y_labels))
        if y_missing:
            all_missing["연간"] = y_missing

    # 시트 순서 고정
    desired_order = [
        "분기_재무제표", "연간_재무제표",
        "지표_분기", "차트_분기",
        "지표_연간", "차트_연간",
        "원본데이터",
    ]
    wb._sheets = [wb[name] for name in desired_order if name in wb.sheetnames]

    today = dt.date.today().strftime("%Y%m%d")
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    outfile = outdir / f"{args.company_name}_{today}.xlsx"
    wb.save(outfile)
    print(json.dumps(
        {
            "saved": str(outfile),
            "quarters_filled": len(quarter_plan),
            "years_filled": len(year_list),
            "missing_indicators": all_missing,
        },
        ensure_ascii=False,
    ))


if __name__ == "__main__":
    main()
