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


def write_raw_sheet(wb: Workbook, quarter_plan: list[dict], year_list: list[str], reports: dict):
    """모든 원자료를 '원본데이터' 시트에 적재하고, 셀 좌표 인덱스를 반환한다."""
    ws = wb.create_sheet("원본데이터")
    ws.sheet_state = "hidden"
    ws["A1"] = "이 시트는 DART 원본 응답값을 그대로 담은 참조용 데이터입니다. 직접 수정하지 마세요."
    ws["A1"].font = Font(name=FONT_NAME, italic=True, size=9)

    row = 3
    cell_index = {}  # (sj_div, account_id, period_label, field) -> "'원본데이터'!$X$Y"

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
                cell_index[(sj_div, item["account_id"], label, "thstrm_amount")] = f"'원본데이터'!${get_column_letter(3)}${row}"
                add_amt = item.get("thstrm_add_amount")
                if add_amt not in (None, ""):
                    ws.cell(row=row, column=4, value=float(str(add_amt).replace(",", ""))).font = BLUE
                    cell_index[(sj_div, item["account_id"], label, "thstrm_add_amount")] = f"'원본데이터'!${get_column_letter(4)}${row}"
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


def build_quarterly_sheet(wb: Workbook, quarter_plan: list[dict], cell_index: dict):
    ws = wb.create_sheet("분기_재무제표")
    ws["A1"] = "단위: 원 | 음영 셀은 원본데이터 시트 링크 또는 수식으로 자동 계산됩니다."
    ws["A1"].font = Font(name=FONT_NAME, italic=True, size=9)

    header_row = 3
    ws.cell(row=header_row, column=1, value="구분")
    ws.cell(row=header_row, column=2, value="계정과목")
    for i, p in enumerate(quarter_plan):
        ws.cell(row=header_row, column=3 + i, value=f"{p['year']}Q{p['q']}")
    style_header(ws, header_row, 2 + len(quarter_plan))

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


def build_annual_sheet(wb: Workbook, year_list: list[str], reports: dict, cell_index: dict):
    ws = wb.create_sheet("연간_재무제표")
    ws["A1"] = "단위: 원 | 사업보고서(연결/별도) 기준, 원본데이터 시트 링크"
    ws["A1"].font = Font(name=FONT_NAME, italic=True, size=9)

    header_row = 3
    ws.cell(row=header_row, column=1, value="구분")
    ws.cell(row=header_row, column=2, value="계정과목")
    for i, year in enumerate(year_list):
        ws.cell(row=header_row, column=3 + i, value=f"{year}(사업보고서)")
    style_header(ws, header_row, 2 + len(year_list))

    row = header_row + 1
    fy_periods = [{"fy": reports.get(y, {}).get("11011")} for y in year_list]
    for sj_div, sj_name in SJ_ORDER:
        accounts = collect_accounts(fy_periods, sj_div, lambda p: p.get("fy"))
        if not accounts:
            continue
        ws.cell(row=row, column=1, value=sj_name).font = BOLD
        row += 1
        for acc in accounts:
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
    build_quarterly_sheet(wb, quarter_plan, cell_index)
    build_annual_sheet(wb, year_list, reports, cell_index)

    # 시트 순서: 분기_재무제표 -> 연간_재무제표 -> 원본데이터(숨김) 순으로 고정
    desired_order = ["분기_재무제표", "연간_재무제표", "원본데이터"]
    wb._sheets = [wb[name] for name in desired_order]

    today = dt.date.today().strftime("%Y%m%d")
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    outfile = outdir / f"{args.company_name}_{today}.xlsx"
    wb.save(outfile)
    print(json.dumps({"saved": str(outfile), "quarters_filled": len(quarter_plan), "years_filled": len(year_list)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
