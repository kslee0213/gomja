"""
"투자판단 종합" 시트를 만든다. 사용자가 첨부한 "성공투자노트" 양식의 셀 배치를
그대로 재현한다(회사명/종목코드, 산업분석·판매과정·경쟁우위·생산과정·제품분석·
리스크·예측가능기간·성장성예측·경쟁상황·수익성예측 텍스트 박스 + 장기(과거+예측)
재무추세 표 + 최종결론).

이 스크립트는 정성적 텍스트나 성장률 가정을 스스로 만들어내지 않는다 — 그건
Claude가 DART 사업의 내용 + 웹 리서치 + 기존 투자분석 결과를 종합해 작성한 뒤
JSON 파일로 넘겨줘야 한다. 이 스크립트는 그 JSON을 받아 정해진 레이아웃에
꽂아 넣고, 과거 실적은 이미 만들어진 "지표_연간" 시트를 참조하는 수식으로,
예측 구간은 Claude가 제시한 성장률 가정을 compounding하는 수식으로 채운다
(둘 다 감사 가능하도록 하드코딩 값이 아니라 수식으로 남긴다).

사용법:
    python build_thesis_sheet.py <기존 투자분석 xlsx 경로> <content.json 경로>

content.json 스키마는 이 파일 하단 CONTENT_SCHEMA_EXAMPLE 참고.
"""
import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT_NAME = "맑은 고딕"
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=11, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="2F5597")
BODY_FONT = Font(name=FONT_NAME, size=10)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="808080")

CONTENT_SCHEMA_EXAMPLE = {
    "company_name": "회사명",
    "stock_code": "005930",
    "industry_analysis": "산업 분석 텍스트",
    "sales_process": "판매 과정 텍스트",
    "competitive_advantage": "경쟁 우위 텍스트",
    "production_process": "생산 과정 텍스트",
    "production_process_extra": "생산/설비투자 관련 보충 텍스트(선택)",
    "products": [
        {"name": "제품1", "description": "제품1 설명"},
        {"name": "제품2", "description": "제품2 설명"},
    ],
    "risk": "리스크 텍스트",
    "predictable_period_text": "예측 가능 기간 텍스트",
    "growth_prediction_text": "성장성 예측 텍스트",
    "competitive_situation": "경쟁 상황(시장점유율 등) 텍스트",
    "profitability_prediction_text": "수익성 예측 텍스트",
    "final_conclusion_text": "최종 결론 텍스트",
    "sustainable_period_years": 10,
    "expected_annual_return_pct": 15.6,
    "projection_years": 10,
    "revenue_growth_assumptions": [0.10] * 10,
    "op_margin_assumptions": [0.11] * 10,
    "net_margin_assumptions": [0.10] * 10,
    "sources": ["웹 리서치에 사용한 출처 URL들(선택, 각주용)"],
}


def find_year_row(ws, label: str) -> int | None:
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=2, max_col=2):
        if row[0].value == label:
            return row[0].row
    return None


def build_thesis_sheet(src_path: str, content_path: str, outdir: str | None = None) -> str:
    wb = load_workbook(src_path)
    if "투자분석" not in wb.sheetnames:
        print("ERROR: 원본 파일에 '투자분석' 시트가 없습니다. 먼저 build_workbook.py로 만드세요.", file=sys.stderr)
        sys.exit(1)
    ind_sheet = wb["지표_연간"] if "지표_연간" in wb.sheetnames else None
    if ind_sheet is None:
        print("ERROR: 원본 파일에 '지표_연간' 시트가 없습니다(연간 데이터 필요).", file=sys.stderr)
        sys.exit(1)

    content = json.loads(Path(content_path).read_text(encoding="utf-8"))

    if "투자판단 종합" in wb.sheetnames:
        del wb["투자판단 종합"]
    ws = wb.create_sheet("투자판단 종합")
    ws.sheet_view.showGridLines = False

    # --- 상단: 회사명 / 종목코드 ---
    ws["B2"] = content.get("company_name", "")
    ws["B2"].font = TITLE_FONT
    ws["D2"] = content.get("stock_code", "")
    ws["D2"].font = Font(name=FONT_NAME, size=12)

    def section(header_cell: str, header_merge: str, header_text: str,
                body_cell: str, body_merge: str, body_text: str):
        ws.merge_cells(header_merge)
        hc = ws[header_cell]
        hc.value = header_text
        hc.font = SECTION_FONT
        hc.fill = SECTION_FILL
        hc.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(body_merge)
        bc = ws[body_cell]
        bc.value = body_text or "(내용 없음)"
        bc.font = BODY_FONT
        bc.alignment = WRAP

    section("B4", "B4:C4", "산업 분석", "B5", "B5:O10", content.get("industry_analysis", ""))
    section("Q4", "Q4:R4", "판매 과정", "Q5", "Q5:T37", content.get("sales_process", ""))
    section("V4", "V4:W4", "경쟁 우위", "V5", "V5:AC15", content.get("competitive_advantage", ""))
    section("B12", "B12:C12", "생산 과정", "B13", "B13:H21", content.get("production_process", ""))

    ws.merge_cells("I12:J12")
    ws["I12"] = "제품 분석"
    ws["I12"].font = SECTION_FONT
    ws["I12"].fill = SECTION_FILL
    ws["I12"].alignment = Alignment(horizontal="center", vertical="center")

    products = content.get("products", [])
    if len(products) > 0:
        ws.merge_cells("L12:O20")
        ws["L12"] = products[0].get("description", "")
        ws["L12"].font = BODY_FONT
        ws["L12"].alignment = WRAP
        ws.merge_cells("I14:K14")
        ws["I14"] = products[0].get("name", "")
        ws["I14"].font = Font(name=FONT_NAME, bold=True)
    if len(products) > 1:
        ws.merge_cells("L21:O29")
        ws["L21"] = products[1].get("description", "")
        ws["L21"].font = BODY_FONT
        ws["L21"].alignment = WRAP
        ws.merge_cells("I22:K22")
        ws["I22"] = products[1].get("name", "")
        ws["I22"].font = Font(name=FONT_NAME, bold=True)

    if content.get("production_process_extra"):
        ws.merge_cells("B22:H29")
        ws["B22"] = content["production_process_extra"]
        ws["B22"].font = BODY_FONT
        ws["B22"].alignment = WRAP

    section("V16", "V16:W16", "리스크", "V17", "V17:AC27", content.get("risk", ""))
    section("V28", "V28:X28", "예측 가능 기간", "V29", "V29:Y43", content.get("predictable_period_text", ""))
    section("Z28", "Z28:AB28", "성장성 예측", "Z29", "Z29:AC35", content.get("growth_prediction_text", ""))
    section("B31", "B31:C31", "경쟁 상황", "B32", "B32:O37", content.get("competitive_situation", ""))
    section("Z36", "Z36:AB36", "수익성 예측", "Z37", "Z37:AC43", content.get("profitability_prediction_text", ""))

    # --- 최종 결론 ---
    ws.merge_cells("V44:W44")
    ws["V44"] = "최종 결론"
    ws["V44"].font = SECTION_FONT
    ws["V44"].fill = SECTION_FILL
    ws["V44"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("V45:X45")
    ws["V45"] = "지속 가능 기간"
    ws["V45"].font = Font(name=FONT_NAME, bold=True)
    ws["Y45"] = content.get("sustainable_period_years")
    ws["Z45"] = "년"

    ws.merge_cells("V46:X46")
    ws["V46"] = "연평균 예상 수익률"
    ws["V46"].font = Font(name=FONT_NAME, bold=True)
    ret = content.get("expected_annual_return_pct")
    ws["Y46"] = (ret / 100) if ret is not None else None
    ws["Y46"].number_format = "0.0%"

    ws.merge_cells("V47:AC55")
    ws["V47"] = content.get("final_conclusion_text", "")
    ws["V47"].font = BODY_FONT
    ws["V47"].alignment = WRAP

    # --- 장기 재무추세 표: 과거(지표_연간 참조 수식) + 예측(가정 성장률 compounding 수식) ---
    # 지표_연간 시트의 헤더 행(연도)과 각 지표 행 위치를 찾는다.
    header_row = None
    for r in range(1, 5):
        if ind_sheet.cell(row=r, column=1).value == "지표":
            header_row = r
            break
    if header_row is None:
        print("ERROR: '지표_연간' 시트 구조를 인식하지 못했습니다.", file=sys.stderr)
        sys.exit(1)

    hist_years = []
    col = 3
    while ind_sheet.cell(row=header_row, column=col).value not in (None, ""):
        hist_years.append((col, ind_sheet.cell(row=header_row, column=col).value))
        col += 1
    n_hist = len(hist_years)

    proj_years_n = int(content.get("projection_years", 10))
    last_year = int(hist_years[-1][1]) if hist_years else None
    proj_years = [last_year + i for i in range(1, proj_years_n + 1)] if last_year else []

    def find_row_in_col_a(ws, label: str) -> int | None:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
            if row[0].value == label:
                return row[0].row
        return None

    ia_sheet = wb["투자분석"] if "투자분석" in wb.sheetnames else None
    ia_row_map = {}
    if ia_sheet is not None:
        ia_row_map = {
            "종가": find_row_in_col_a(ia_sheet, "종가"),
            "시가총액": find_row_in_col_a(ia_sheet, "시가총액"),
            "PER(배)": find_row_in_col_a(ia_sheet, "PER(배)"),
            "PBR(배)": find_row_in_col_a(ia_sheet, "PBR(배)"),
        }

    row_map = {
        "매출액": find_year_row(ind_sheet, "매출액"),
        "영업이익": find_year_row(ind_sheet, "영업이익"),
        "당기순이익": find_year_row(ind_sheet, "당기순이익"),
        "자산총계": find_year_row(ind_sheet, "자산총계"),
        "부채총계": find_year_row(ind_sheet, "부채총계"),
        "자본총계": find_year_row(ind_sheet, "자본총계"),
    }

    TABLE_START_COL = 4  # D열부터
    year_row = 39
    idx_row = 38
    ws.cell(row=idx_row, column=1, value="(억원)").font = NOTE_FONT
    ws.cell(row=idx_row, column=2, value="연차")
    for i, _ in enumerate(hist_years + [(None, y) for y in proj_years]):
        c = TABLE_START_COL + i
        ws.cell(row=idx_row, column=c, value=i + 1)
    ws.cell(row=year_row, column=2, value="연도")
    all_years = [y for _, y in hist_years] + proj_years
    for i, y in enumerate(all_years):
        ws.cell(row=year_row, column=TABLE_START_COL + i, value=int(y))

    labels = [
        ("시가총액", 40), ("주가", 41), ("PER", 42), ("PBR", 43),
        ("매출액", 44), ("Growth", 45), ("영업이익", 46), ("영업이익률", 47), ("Growth", 48),
        ("순이익", 49), ("순이익률", 50), ("Growth", 51), ("ROE", 52),
        ("자산", 53), ("부채", 54), ("자본", 55),
    ]
    for name, r in labels:
        ws.cell(row=r, column=2, value=name)

    # 시가총액·주가·PER·PBR은 이미 "투자분석" 시트 L섹션에서 계산해 둔 값을
    # 그대로 참조한다(과거 실적 연도만 — 미래 주가는 예측하지 않는다).
    if ia_sheet is not None and all(ia_row_map.values()):
        for i in range(n_hist):
            col_letter = get_column_letter(TABLE_START_COL + i)
            ia_col = get_column_letter(3 + i)  # 투자분석 L섹션도 동일한 연도 순서로 C열부터 시작
            ws[f"{col_letter}40"] = f"='투자분석'!{ia_col}{ia_row_map['시가총액']}"
            ws[f"{col_letter}41"] = f"='투자분석'!{ia_col}{ia_row_map['종가']}"
            ws[f"{col_letter}42"] = f"='투자분석'!{ia_col}{ia_row_map['PER(배)']}"
            ws[f"{col_letter}43"] = f"='투자분석'!{ia_col}{ia_row_map['PBR(배)']}"
            ws.cell(row=40, column=TABLE_START_COL + i).number_format = "#,##0.0"
            ws.cell(row=41, column=TABLE_START_COL + i).number_format = "#,##0"
            ws.cell(row=42, column=TABLE_START_COL + i).number_format = "0.0"
            ws.cell(row=43, column=TABLE_START_COL + i).number_format = "0.00"
    else:
        ws.cell(row=40, column=1, value=(
            "※ '투자분석' 시트에서 시가총액/종가/PER/PBR을 찾지 못해 비워둡니다"
            "(KRX_AUTH_KEY 없이 만든 파일일 수 있습니다)."
        )).font = NOTE_FONT

    rev_g = content.get("revenue_growth_assumptions", [0.1] * proj_years_n)
    op_m = content.get("op_margin_assumptions", [0.1] * proj_years_n)
    net_m = content.get("net_margin_assumptions", [0.1] * proj_years_n)

    for i in range(len(all_years)):
        col_letter = get_column_letter(TABLE_START_COL + i)
        is_hist = i < n_hist

        if is_hist:
            src_col = get_column_letter(hist_years[i][0])
            if row_map["매출액"]:
                ws[f"{col_letter}44"] = f"='지표_연간'!{src_col}{row_map['매출액']}"
            if row_map["영업이익"]:
                ws[f"{col_letter}46"] = f"='지표_연간'!{src_col}{row_map['영업이익']}"
            if row_map["당기순이익"]:
                ws[f"{col_letter}49"] = f"='지표_연간'!{src_col}{row_map['당기순이익']}"
            if row_map["자산총계"]:
                ws[f"{col_letter}53"] = f"='지표_연간'!{src_col}{row_map['자산총계']}"
            if row_map["부채총계"]:
                ws[f"{col_letter}54"] = f"='지표_연간'!{src_col}{row_map['부채총계']}"
            if row_map["자본총계"]:
                ws[f"{col_letter}55"] = f"='지표_연간'!{src_col}{row_map['자본총계']}"
            ws[f"{col_letter}47"] = f"=IFERROR({col_letter}46/{col_letter}44,\"\")"
            ws[f"{col_letter}50"] = f"=IFERROR({col_letter}49/{col_letter}44,\"\")"
            ws[f"{col_letter}52"] = f"=IFERROR({col_letter}49/{col_letter}55,\"\")"
        else:
            p = i - n_hist
            prev_col = get_column_letter(TABLE_START_COL + i - 1)
            g = rev_g[p] if p < len(rev_g) else rev_g[-1] if rev_g else 0.1
            om = op_m[p] if p < len(op_m) else op_m[-1] if op_m else 0.1
            nm = net_m[p] if p < len(net_m) else net_m[-1] if net_m else 0.1
            ws[f"{col_letter}44"] = f"={prev_col}44*(1+{g})"
            ws[f"{col_letter}46"] = f"={col_letter}44*{om}"
            ws[f"{col_letter}47"] = om
            ws[f"{col_letter}49"] = f"={col_letter}44*{nm}"
            ws[f"{col_letter}50"] = nm
            # 자산/부채/자본은 단순화: 자본은 전기 자본+당기순이익(배당 가정 없이 전액 유보) 누적, 부채는 전기와 동일 유지, 자산=부채+자본
            ws[f"{col_letter}55"] = f"={prev_col}55+{col_letter}49"
            ws[f"{col_letter}54"] = f"={prev_col}54"
            ws[f"{col_letter}53"] = f"={col_letter}54+{col_letter}55"
            ws[f"{col_letter}52"] = f"=IFERROR({col_letter}49/{col_letter}55,\"\")"

        # Growth(전기 대비 증가율)는 과거·예측 구간 구분 없이 앞 열을 참조하는
        # 수식으로 통일한다(첫 연도는 전기가 없어 비워둔다).
        if i > 0:
            prev_col_g = get_column_letter(TABLE_START_COL + i - 1)
            ws[f"{col_letter}45"] = f"=IFERROR(({col_letter}44-{prev_col_g}44)/{prev_col_g}44,\"\")"
            ws[f"{col_letter}48"] = f"=IFERROR(({col_letter}46-{prev_col_g}46)/{prev_col_g}46,\"\")"
            ws[f"{col_letter}51"] = f"=IFERROR(({col_letter}49-{prev_col_g}49)/{prev_col_g}49,\"\")"

        for r in (44, 46, 49, 53, 54, 55):
            ws.cell(row=r, column=TABLE_START_COL + i).number_format = "#,##0.0"
        for r in (45, 47, 48, 50, 51, 52):
            ws.cell(row=r, column=TABLE_START_COL + i).number_format = "0.0%"

    ws.cell(row=len(all_years) + 57, column=2, value=(
        "※ 굵게 표시되지 않은 연도(마지막 실적연도 이후)는 사용자/Claude가 제시한 성장률 가정에 따른 추정치입니다. "
        "실적이 아니라 시나리오이므로 투자 판단 시 가정을 직접 검토하세요."
    )).font = NOTE_FONT

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 14
    for col in range(3, 30):
        ws.column_dimensions[get_column_letter(col)].width = 11

    if outdir is None:
        # 기본 동작(v0.9.1부터): 새 파일을 만들지 않고 원본 파일 그대로 덮어써서
        # "재무제표+투자분석+투자판단 종합"이 전부 한 파일에 담기게 한다.
        outfile = Path(src_path)
    else:
        today = __import__("datetime").date.today().strftime("%Y%m%d")
        outdir_path = Path(outdir)
        outdir_path.mkdir(parents=True, exist_ok=True)
        outfile = outdir_path / f"{content.get('company_name','기업')}_투자판단종합_{today}.xlsx"
    wb.save(outfile)
    return str(outfile)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("src_xlsx", help="기존 투자분석 포함 xlsx 경로 (예: 연간 파일)")
    ap.add_argument("content_json", help="정성 콘텐츠 JSON 파일 경로")
    ap.add_argument(
        "--outdir", default=None,
        help="지정하면 새 파일(...''_투자판단종합_'YYYYMMDD.xlsx)로 따로 저장한다. "
             "지정하지 않으면(기본값) src_xlsx 파일 자체에 시트를 추가해 덮어쓴다 "
             "— 재무제표+투자분석+투자판단 종합이 파일 하나로 합쳐진다.",
    )
    args = ap.parse_args()

    outfile = build_thesis_sheet(args.src_xlsx, args.content_json, args.outdir)
    print(json.dumps({"saved": outfile}, ensure_ascii=False))


if __name__ == "__main__":
    main()
