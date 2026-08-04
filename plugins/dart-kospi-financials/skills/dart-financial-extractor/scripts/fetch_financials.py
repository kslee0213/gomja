"""
fnlttSinglAcntAll.json을 호출해 특정 (연도, 보고서) 조합의 전체 재무제표를 가져오고
로컬 캐시에 저장한다. CFS(연결) 응답이 비어 있으면 OFS(별도)로 자동 재시도한다.

사용법:
    python fetch_financials.py <corp_code> <bsns_year> <reprt_code> [fs_div]

    fs_div 생략 시 CFS로 먼저 시도하고, 빈 응답이면 OFS로 재시도한다.

환경 변수:
    DART_API_KEY  필수.

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


def get_api_key() -> str:
    key = os.environ.get("DART_API_KEY")
    if not key:
        print("ERROR: 환경 변수 DART_API_KEY가 설정되어 있지 않습니다.", file=sys.stderr)
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


def fetch_one(corp_code: str, bsns_year: str, reprt_code: str, fs_div: str | None) -> dict:
    api_key = get_api_key()
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
    if len(sys.argv) < 4:
        print(
            "사용법: python fetch_financials.py <corp_code> <bsns_year> <reprt_code> [fs_div]",
            file=sys.stderr,
        )
        sys.exit(1)

    corp_code, bsns_year, reprt_code = sys.argv[1], sys.argv[2], sys.argv[3]
    fs_div = sys.argv[4] if len(sys.argv) > 4 else None

    result = fetch_one(corp_code, bsns_year, reprt_code, fs_div)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
