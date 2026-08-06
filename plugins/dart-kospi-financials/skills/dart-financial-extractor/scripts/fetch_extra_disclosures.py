"""
DART Open API의 추가 공시정보 4종을 가져와 캐시에 저장한다.
fnlttSinglAcntAll(재무제표)과 달리 이 4종은 "투자분석" 시트 전용이다.

- 주식총수현황 (stockTotqySttus): 발행주식총수 → 시가총액 계산에 필요
- 배당에 관한 사항 (alotMatter): 배당성향, 주당배당금
- 최대주주 현황 (hyslrSttus): 대주주 명단
- 자기주식 취득/처분 현황 (tesstkAcqsDspsSttus): 자사주 매입/소각 이력

사용법:
    python fetch_extra_disclosures.py --api-key <키> <corp_code> <bsns_year> <reprt_code>

캐시 위치: cache/extra_{corp_code}_{bsns_year}_{reprt_code}.json
"""
import json
import os
import sys
from pathlib import Path

import requests

BASE_URL = "https://opendart.fss.or.kr/api"
SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

ENDPOINTS = {
    "주식총수현황": "stockTotqySttus",
    "배당": "alotMatter",
    "최대주주현황": "hyslrSttus",
    "자기주식현황": "tesstkAcqsDspsSttus",
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


def call_api(api_key: str, endpoint: str, corp_code: str, bsns_year: str, reprt_code: str) -> dict:
    resp = requests.get(
        f"{BASE_URL}/{endpoint}.json",
        params={
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "bsns_year": bsns_year,
            "reprt_code": reprt_code,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


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
            "사용법: python fetch_extra_disclosures.py --api-key <키> <corp_code> <bsns_year> <reprt_code>",
            file=sys.stderr,
        )
        sys.exit(1)

    corp_code, bsns_year, reprt_code = args[0], args[1], args[2]
    api_key = get_api_key(cli_key)

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    result = {}
    for label, endpoint in ENDPOINTS.items():
        try:
            data = call_api(api_key, endpoint, corp_code, bsns_year, reprt_code)
        except requests.RequestException as e:
            data = {"status": "ERR", "message": str(e)}
        result[label] = data

    cache_path = CACHE_DIR / f"extra_{corp_code}_{bsns_year}_{reprt_code}.json"
    cache_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

    summary = {
        label: (data.get("status"), data.get("message"))
        for label, data in result.items()
    }
    print(json.dumps({"saved": str(cache_path), "status_by_endpoint": summary}, ensure_ascii=False))


if __name__ == "__main__":
    main()
