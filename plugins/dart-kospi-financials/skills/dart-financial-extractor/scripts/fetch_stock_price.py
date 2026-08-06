"""
공공데이터포털의 "금융위원회_주식시세정보"(getStockPriceInfo) API로 특정일 종가·
시가총액·상장주식수를 가져와 캐시에 저장한다.

KRX Open API(openapi.krx.co.kr) 대신 이 API를 쓰는 이유:
- data.go.kr 서비스는 서비스별 개별 승인 없이(또는 훨씬 빠르게) 쓸 수 있는 경우가 많다
  (KRX Open API는 "유가증권 일별매매정보"처럼 상품 단위로 별도 승인이 필요해 막히는 사례가 있었음).
- 응답에 종가(clpr)뿐 아니라 상장주식수(lstgStCnt)·시가총액(mrktTotAmt)까지 바로 들어있어
  별도로 곱셈할 필요가 없다.
- GET + 쿼리파라미터 방식이라 POST + 커스텀 헤더 방식보다 다루기 쉽다.

참고: 이 API는 실시간이 아니고 기준일 다음 영업일 오후 1시 이후 갱신된다. 연말 종가처럼
과거 특정일 조회 용도(이 스킬의 목적)에는 문제없다.

사용법:
    python fetch_stock_price.py --service-key <데이터포털 인증키(Decoding)> <stock_code> <YYYYMMDD>
    python fetch_stock_price.py --service-key <키> <stock_code> <YYYYMMDD> --raw   # 원본 응답 확인용

<stock_code>는 DART corpCode.xml의 stock_code(6자리, 예: 005930)를 그대로 쓴다.
지정한 날짜에 데이터가 없으면(휴장일/미갱신) 최대 --max-back일 전까지 거슬러 올라간다.

캐시 위치: cache/price_{stock_code}_{YYYYMMDD}.json
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import requests

BASE_URL = "https://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo"

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"


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


def fetch_one_day(service_key: str, bas_dd: str, num_of_rows: int = 10000) -> list[dict]:
    resp = requests.get(
        BASE_URL,
        params={
            "serviceKey": service_key,
            "resultType": "json",
            "basDt": bas_dd,
            "numOfRows": num_of_rows,
            "pageNo": 1,
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()

    header = data.get("response", {}).get("header", {})
    result_code = header.get("resultCode")
    if result_code not in (None, "00", "0"):
        raise RuntimeError(f"data.go.kr 오류 (resultCode={result_code}): {header.get('resultMsg')}")

    body = data.get("response", {}).get("body", {})
    items = body.get("items", {})
    item = items.get("item", []) if isinstance(items, dict) else []
    if isinstance(item, dict):  # 결과가 1건이면 리스트가 아니라 dict로 오는 경우가 있다
        item = [item]
    return item


def find_stock_row(rows: list[dict], stock_code: str) -> dict | None:
    for row in rows:
        if str(row.get("srtnCd", "")).strip() == stock_code:
            return row
    return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("stock_code", help="6자리 종목코드 (예: 005930)")
    ap.add_argument("date", help="YYYYMMDD, 조회 기준일(보통 연말/최근 영업일)")
    ap.add_argument("--service-key", required=True, help="공공데이터포털 '일반 인증키(Decoding)'")
    ap.add_argument("--raw", action="store_true", help="원본 응답을 그대로 출력하고 종료")
    ap.add_argument("--max-back", type=int, default=10, help="데이터 없으면 최대 며칠 전까지 재시도할지")
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
            rows = fetch_one_day(args.service_key, bas_dd)
        except (requests.RequestException, RuntimeError) as e:
            print(f"ERROR: 호출 실패 ({bas_dd}): {e}", file=sys.stderr)
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
            "--max-back을 늘리거나 --raw로 원본 응답을 확인하세요.",
            file=sys.stderr,
        )
        sys.exit(1)

    row = find_stock_row(rows, args.stock_code)
    result = {
        "stock_code": args.stock_code,
        "requested_date": args.date,
        "used_date": used_date,
        "close_price": to_number(row.get("clpr")),
        "market_cap": to_number(row.get("mrktTotAmt")),
        "listed_shares": to_number(row.get("lstgStCnt")),
        "raw": row,
    }

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"price_{args.stock_code}_{args.date}.json"
    cache_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"saved": str(cache_path), **{k: v for k, v in result.items() if k != "raw"}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
