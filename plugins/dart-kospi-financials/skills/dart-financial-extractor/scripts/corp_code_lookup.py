"""
DART corpCode.xml을 다운로드/캐시하고 기업명으로 corp_code를 조회한다.

사용법:
    python corp_code_lookup.py "삼성전자"

환경 변수:
    DART_API_KEY  필수. DART Open API 인증키.

동작:
    1. cache/CORPCODE.xml이 없거나 7일 이상 오래되었으면 재다운로드한다.
    2. 기업명으로 정확히 일치하는 후보를 찾는다(공백 제거 후 비교).
    3. 정확히 일치하는 후보가 여러 개면 전부, 하나도 없으면 부분일치 후보를 출력한다.
    4. company.json으로 법인구분(corp_cls)을 조회해 KOSPI(Y) 여부를 함께 보여준다.

주의:
    - API 키 값 자체는 어떤 print/log/파일에도 출력하지 않는다.
    - 네트워크 호출 실패 시 명확한 에러 메시지와 함께 종료 코드 1을 반환한다.
"""
import io
import json
import os
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import requests

BASE_URL = "https://opendart.fss.or.kr/api"
SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"
CORP_CODE_XML = CACHE_DIR / "CORPCODE.xml"
CACHE_MAX_AGE_SEC = 7 * 24 * 3600


def get_api_key() -> str:
    key = os.environ.get("DART_API_KEY")
    if not key:
        print(
            "ERROR: 환경 변수 DART_API_KEY가 설정되어 있지 않습니다. "
            "Cowork 환경 변수에 DART_API_KEY를 등록한 뒤 다시 실행하세요.",
            file=sys.stderr,
        )
        sys.exit(1)
    return key


def refresh_corp_code_cache(api_key: str) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if CORP_CODE_XML.exists():
        age = time.time() - CORP_CODE_XML.stat().st_mtime
        if age < CACHE_MAX_AGE_SEC:
            return

    resp = requests.get(
        f"{BASE_URL}/corpCode.xml", params={"crtfc_key": api_key}, timeout=30
    )
    resp.raise_for_status()

    # 정상 응답은 ZIP 바이너리. 에러 시 XML 오류 메시지가 그대로 온다.
    content_type = resp.headers.get("Content-Type", "")
    if "xml" in content_type and b"<result>" in resp.content[:200]:
        tree = ET.fromstring(resp.content)
        status = tree.findtext("status")
        message = tree.findtext("message")
        print(f"ERROR: corpCode.xml 다운로드 실패 (status={status}): {message}", file=sys.stderr)
        sys.exit(1)

    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        zf.extract("CORPCODE.xml", path=CACHE_DIR)


def load_corp_list() -> list[dict]:
    tree = ET.parse(CORP_CODE_XML)
    root = tree.getroot()
    corps = []
    for node in root.findall("list"):
        corps.append(
            {
                "corp_code": (node.findtext("corp_code") or "").strip(),
                "corp_name": (node.findtext("corp_name") or "").strip(),
                "stock_code": (node.findtext("stock_code") or "").strip(),
                "modify_date": (node.findtext("modify_date") or "").strip(),
            }
        )
    return corps


def find_candidates(corps: list[dict], name: str) -> list[dict]:
    target = name.replace(" ", "")
    exact = [c for c in corps if c["corp_name"].replace(" ", "") == target and c["stock_code"]]
    if exact:
        return exact
    partial = [c for c in corps if target in c["corp_name"].replace(" ", "") and c["stock_code"]]
    return partial[:10]


def check_market(api_key: str, corp_code: str) -> str | None:
    resp = requests.get(
        f"{BASE_URL}/company.json",
        params={"crtfc_key": api_key, "corp_code": corp_code},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("status") != "000":
        return None
    return data.get("corp_cls")  # Y=코스피, K=코스닥, N=코넥스, E=기타


def main() -> None:
    if len(sys.argv) < 2:
        print("사용법: python corp_code_lookup.py <기업명>", file=sys.stderr)
        sys.exit(1)

    company_name = sys.argv[1]
    api_key = get_api_key()

    refresh_corp_code_cache(api_key)
    corps = load_corp_list()
    candidates = find_candidates(corps, company_name)

    if not candidates:
        print(json.dumps({"found": False, "candidates": []}, ensure_ascii=False))
        return

    market_labels = {"Y": "코스피(KOSPI)", "K": "코스닥(KOSDAQ)", "N": "코넥스(KONEX)", "E": "기타"}
    result = []
    for c in candidates:
        corp_cls = check_market(api_key, c["corp_code"])
        result.append(
            {
                "corp_code": c["corp_code"],
                "corp_name": c["corp_name"],
                "stock_code": c["stock_code"],
                "market": market_labels.get(corp_cls, "확인불가"),
                "corp_cls": corp_cls,
            }
        )

    print(json.dumps({"found": True, "candidates": result}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
