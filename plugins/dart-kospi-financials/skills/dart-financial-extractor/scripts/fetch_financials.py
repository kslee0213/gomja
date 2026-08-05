"""
fnlttSinglAcntAll.json을 호출해 특정 (연도, 보고서) 조합의 전체 재무제표를 가져오고
로컬 캐시에 저장한다. CFS(연결) 응답이 비어 있으면 OFS(별도)로 자동 재시도한다.

사용법:
    python fetch_financials.py --api-key <키> <corp_code> <bsns_year> <reprt_code> [fs_div]

    fs_div 생략 시 CFS로 먼저 시도하고, 빈 응답이면 OFS로 재시도한다.

API 키:
    --api-key 인자로 전달하는 것을 우선한다 (Cowork 등 OS 환경 변수가
    전달되지 않는 격리 환경에서도 동작). --api-key가 없으면 환경 변수
    DART_API_KEY를 폴백으로 사용한다 (로컬 Claude Code 등 환경 변수가
    정상적으로 전달되는 경우).

출력:
    표준출력으로 JSON을 반환한다:
    {
      "bsns_year": ..., "reprt_code": ..., "fs_div_used": "CFS" | "OFS",
      "status": "000", "items": { "BS": [...], "IS": [...], "CIS": [...], "CF": [...] }
    }
    캐시 파일: cache/{corp_code}_{bsns_year}_{reprt_code}_{fs_div}.json
"""
import json
import os
import sys
from pathlib import Path

import requests

BASE_URL = "https://opendart.fss.or.kr/api"
SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

RELEVANT_SJ_DIV = {"BS", "IS", "CIS", "CF"}

REPRT_NAMES = {
    "11013": "1분기보고서",
    "11012": "반기보고서",
    "11014": "3분기보고서",
    "11011": "사업보고서",
}


def get_api_key(cli_key: str | None) -> str:
    key = cli_key or os.environ.get("DART_API_KEY")
    if not key:
        print(
            "ERROR: API 키가 없습니다. --api-key 인자로 전달하거나 "
            "환경 변수 DART_API_KEY를 설정하세요.",
            file=sys.stderr,
        )
        sys.exit(1)
    return key


def cache_path(corp_code: str, bsns_year: str, reprt_code: str, fs_div: str) -> Path:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return CACHE_DIR / f"{corp_code}_{bsns_year}_{reprt_code}_{fs_div}.json"


def call_api(api_key: str, corp_code: str, bsns_year: str, reprt_code: str, fs_div: str) -> dict:
    resp = requests.get(
        f"{BASE_URL}/fnlttSinglAcntAll.json",
        params={
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "bsns_year": bsns_year,
            "reprt_code": reprt_code,
            "fs_div": fs_div,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def group_by_sj_div(items: list[dict]) -> dict:
    grouped: dict[str, list] = {k: [] for k in RELEVANT_SJ_DIV}
    for item in items:
        sj = item.get("sj_div")
        if sj in RELEVANT_SJ_DIV:
            grouped[sj].append(item)
    for sj in grouped:
        grouped[sj].sort(key=lambda x: int(x.get("ord") or 0))
    return grouped


def fetch_one(api_key: str, corp_code: str, bsns_year: str, reprt_code: str, fs_div: str | None) -> dict:
    tried_fs_divs = [fs_div] if fs_div else ["CFS", "OFS"]

    for fdiv in tried_fs_divs:
        cp = cache_path(corp_code, bsns_year, reprt_code, fdiv)
        if cp.exists():
            cached = json.loads(cp.read_text(encoding="utf-8"))
            if cached.get("status") == "000" and cached.get("items", {}).get("BS"):
                return cached

        data = call_api(api_key, corp_code, bsns_year, reprt_code, fdiv)
        status = data.get("status")

        if status == "000" and data.get("list"):
            result = {
                "corp_code": corp_code,
                "bsns_year": bsns_year,
                "reprt_code": reprt_code,
                "reprt_name": REPRT_NAMES.get(reprt_code, reprt_code),
                "fs_div_used": fdiv,
                "status": status,
                "items": group_by_sj_div(data["list"]),
            }
            cp.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            return result

        # status 013 = 데이터 없음 → 다음 fs_div로 재시도 (fs_div가 명시 지정된 경우는 재시도 없이 반환)
        if fs_div:
            return {
                "corp_code": corp_code,
                "bsns_year": bsns_year,
                "reprt_code": reprt_code,
                "reprt_name": REPRT_NAMES.get(reprt_code, reprt_code),
                "fs_div_used": fdiv,
                "status": status,
                "message": data.get("message"),
                "items": {},
            }

    # 모든 fs_div 시도 후에도 데이터 없음
    return {
        "corp_code": corp_code,
        "bsns_year": bsns_year,
        "reprt_code": reprt_code,
        "reprt_name": REPRT_NAMES.get(reprt_code, reprt_code),
        "fs_div_used": None,
        "status": "013",
        "message": "CFS/OFS 모두 데이터 없음 (해당 보고서 미제출 가능성)",
        "items": {},
    }


def main() -> None:
    args = sys.argv[1:]
    cli_key = None
    if "--api-key" in args:
        idx = args.index("--api-key")
        if idx + 1 >= len(args):
            print("ERROR: --api-key 다음에 키 값이 필요합니다.", file=sys.stderr)
            sys.exit(1)
        cli_key = args[idx + 1]
        args = args[:idx] + args[idx + 2:]

    if len(args) < 3:
        print(
            "사용법: python fetch_financials.py --api-key <키> <corp_code> <bsns_year> <reprt_code> [fs_div]",
            file=sys.stderr,
        )
        sys.exit(1)

    corp_code, bsns_year, reprt_code = args[0], args[1], args[2]
    fs_div = args[3] if len(args) > 3 else None

    api_key = get_api_key(cli_key)
    result = fetch_one(api_key, corp_code, bsns_year, reprt_code, fs_div)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
