"""
KRX Open API로 유가증권(코스피) 종목의 특정일 종가/시가총액 등을 가져와 캐시에 저장한다.

⚠️ 중요 — 첫 실행 시 반드시 확인할 것:
KRX Open API의 정확한 JSON 응답 필드명은 회사마다/시점마다 문서가 조금씩 다를 수
있어, 이 스크립트는 실제 응답을 신뢰할 수 있는 소식통(velog 등 3rd-party 예제)을
바탕으로 최대한 정확히 작성했지만 100% 검증되지는 않았다. 처음 실행했을 때
--raw 옵션으로 원본 응답을 한 번 찍어보고, 아래 FIELD_CANDIDATES에서 실제 필드명과
다르면 맞게 고쳐야 한다.

사용법:
    python fetch_stock_price.py --auth-key <키> <stock_code> <YYYYMMDD>
    python fetch_stock_price.py --auth-key <키> <stock_code> <YYYYMMDD> --raw   # 원본 응답 확인용

<stock_code>는 DART corpCode.xml의 stock_code(6자리, 예: 005930)를 그대로 쓴다.
지정한 날짜가 휴장일이면 최대 10일 전까지 거슬러 올라가며 재시도한다.

캐시 위치: cache/price_{stock_code}_{YYYYMMDD}.json
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import requests

# KRX Open API 실제 호출 도메인은 openapi.krx.co.kr(포털)이 아니라
# data-dbg.krx.co.kr(API 서버)이다. Cowork 네트워크 허용 목록에도 이 도메인을
# 추가해야 한다.
BASE_URL = "https://data-dbg.krx.co.kr/svc/apis/sto"
ENDPOINT_KOSPI_DAILY = "/stk_bydd_trd"  # 유가증권 일별매매정보

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

# 실제 응답에서 확인되면 이 후보 목록 순서를 조정하거나 정확한 필드명으로 좁힌다.
FIELD_CANDIDATES = {
    "stock_code": ["ISU_SRT_CD", "ISU_CD", "SRTN_CD"],
    "close_price": ["TDD_CLSPRC", "CLSPRC"],
    "market_cap": ["MKTCAP", "MKT_CAP"],
    "listed_shares": ["LIST_SHRS", "LIST_SHARE_CNT"],
    "trade_date": ["BAS_DD", "TRD_DD"],
}


def pick(row: dict, candidates: list[str]):
    for c in candidates:
        if c in row and row[c] not in (None, ""):
            return row[c]
    return None


def to_number(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fetch_one_day(auth_key: str, bas_dd: str) -> list[dict]:
    resp = requests.post(
        BASE_URL + ENDPOINT_KOSPI_DAILY,
        headers={
            "AUTH_KEY": auth_key.strip(),
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        json={"basDd": bas_dd},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    # 공공/거래소 오픈API에 흔한 패턴들을 순서대로 시도
    for key in ("OutBlock_1", "output", "response"):
        if isinstance(data, dict) and key in data and isinstance(data[key], list):
            return data[key]
    if isinstance(data, list):
        return data
    # 못 찾으면 원본을 그대로 보여줘서 사람이 구조를 파악하게 한다
    print("WARNING: 예상한 목록 필드를 못 찾았습니다. 원본 응답:", file=sys.stderr)
    print(json.dumps(data, ensure_ascii=False)[:2000], file=sys.stderr)
    return []


def find_stock_row(rows: list[dict], stock_code: str) -> dict | None:
    for row in rows:
        code = pick(row, FIELD_CANDIDATES["stock_code"])
        if code and str(code).strip().lstrip("A") == stock_code:
            return row
    return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("stock_code", help="6자리 종목코드 (예: 005930)")
    ap.add_argument("date", help="YYYYMMDD, 조회 기준일(보통 연말/최근 영업일)")
    ap.add_argument("--auth-key", required=True)
    ap.add_argument("--raw", action="store_true", help="원본 응답을 그대로 stderr에 출력하고 종료")
    ap.add_argument("--max-back", type=int, default=10, help="휴장일이면 최대 며칠 전까지 재시도할지")
    args = ap.parse_args()

    target = dt.datetime.strptime(args.date, "%Y%m%d").date()
    rows = None
    used_date = None
    for i in range(args.max_back + 1):
        d = target - dt.timedelta(days=i)
        if d.weekday() >= 5:  # 토/일 스킵
            continue
        bas_dd = d.strftime("%Y%m%d")
        try:
            rows = fetch_one_day(args.auth_key, bas_dd)
        except requests.RequestException as e:
            print(f"ERROR: KRX 호출 실패 ({bas_dd}): {e}", file=sys.stderr)
            sys.exit(1)

        if args.raw:
            print(json.dumps(rows[:3], ensure_ascii=False, indent=2))
            return

        if rows:
            row = find_stock_row(rows, args.stock_code)
            if row:
                used_date = bas_dd
                break
        rows = None

    if not rows or used_date is None:
        print(
            f"ERROR: {args.date} 기준 {args.max_back}일 이내에서 {args.stock_code} 데이터를 찾지 못했습니다. "
            "휴장일 범위를 늘려보거나(--max-back), --raw로 원본 응답을 확인하세요.",
            file=sys.stderr,
        )
        sys.exit(1)

    row = find_stock_row(rows, args.stock_code)
    result = {
        "stock_code": args.stock_code,
        "requested_date": args.date,
        "used_date": used_date,
        "close_price": to_number(pick(row, FIELD_CANDIDATES["close_price"])),
        "market_cap": to_number(pick(row, FIELD_CANDIDATES["market_cap"])),
        "listed_shares": to_number(pick(row, FIELD_CANDIDATES["listed_shares"])),
        "raw": row,
    }

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"price_{args.stock_code}_{args.date}.json"
    cache_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"saved": str(cache_path), **{k: v for k, v in result.items() if k != "raw"}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
