"""
KRX Open API "유가증권 일별매매정보"(stk_bydd_trd)로 특정일 종가·시가총액·
상장주식수를 가져와 캐시에 저장한다.

이 스크립트는 KRX가 공식 배포한 API Spec 문서(요청/응답 필드 전부 명시)를 그대로
반영했으므로 필드명 추측이 아니라 확정값이다.

- Endpoint: https://data-dbg.krx.co.kr/svc/apis/sto/stk_bydd_trd
- 인증: HTTP 헤더 AUTH_KEY
- 요청: POST, JSON body {"basDd": "YYYYMMDD"}
- 응답: {"OutBlock_1": [{"ISU_CD":..., "TDD_CLSPRC":..., "MKTCAP":..., "LIST_SHRS":..., ...}, ...]}

"유가증권 일별매매정보" 서비스에 대한 이용 신청이 KRX 마이페이지에서 승인되어 있어야
한다(인증키 자체 발급과는 별개 절차). 승인 전에는 "Unauthorized API Call" 에러가 난다.

사용법:
    python fetch_stock_price.py --auth-key <KRX 인증키> <stock_code> <YYYYMMDD>
    python fetch_stock_price.py --auth-key <키> <stock_code> <YYYYMMDD> --raw   # 원본 응답 확인용

<stock_code>는 DART corpCode.xml의 stock_code(6자리, 예: 005930)를 그대로 쓴다.
지정한 날짜에 데이터가 없으면(휴장일 등) 최대 --max-back일 전까지 거슬러 올라간다.

캐시 위치: cache/price_{stock_code}_{YYYYMMDD}.json
"""
import argparse
import datetime as dt
import json
import sys
from pathlib import Path

import requests

# 포털(openapi.krx.co.kr)이 아니라 실제 API 서버 도메인. Cowork 네트워크 허용
# 목록에도 이 도메인을 추가해야 한다.
BASE_URL = "https://data-dbg.krx.co.kr/svc/apis/sto/stk_bydd_trd"

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


def fetch_one_day(auth_key: str, bas_dd: str) -> list[dict]:
    resp = requests.post(
        BASE_URL,
        headers={
            "AUTH_KEY": auth_key.strip(),
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        json={"basDd": bas_dd},
        timeout=30,
    )
    if resp.status_code == 401 or "Unauthorized" in resp.text[:200]:
        raise RuntimeError(
            "인증/권한 오류입니다. 인증키 자체는 맞아도 '유가증권 일별매매정보' 서비스에 "
            "대한 이용 신청이 KRX 마이페이지에서 아직 승인되지 않았을 수 있습니다. "
            f"응답: {resp.text[:300]}"
        )
    resp.raise_for_status()
    data = resp.json()
    return data.get("OutBlock_1", [])


def find_stock_row(rows: list[dict], stock_code: str) -> dict | None:
    for row in rows:
        code = str(row.get("ISU_CD", "")).strip()
        # KRX 종목코드가 'A005930'처럼 접두어가 붙어 오는 경우까지 대비
        if code == stock_code or code.lstrip("A") == stock_code:
            return row
    return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("stock_code", help="6자리 종목코드 (예: 005930)")
    ap.add_argument("date", help="YYYYMMDD, 조회 기준일(보통 연말/최근 영업일)")
    ap.add_argument("--auth-key", required=True, help="KRX Open API 인증키")
    ap.add_argument("--raw", action="store_true", help="원본 응답을 그대로 출력하고 종료")
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
        except (requests.RequestException, RuntimeError) as e:
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
        "close_price": to_number(row.get("TDD_CLSPRC")),
        "market_cap": to_number(row.get("MKTCAP")),
        "listed_shares": to_number(row.get("LIST_SHRS")),
        "raw": row,
    }

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"price_{args.stock_code}_{args.date}.json"
    cache_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"saved": str(cache_path), **{k: v for k, v in result.items() if k != "raw"}}, ensure_ascii=False))


if __name__ == "__main__":
    main()

