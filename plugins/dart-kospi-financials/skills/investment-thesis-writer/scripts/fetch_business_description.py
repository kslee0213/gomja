"""
DART "공시서류원본파일" API로 최신 사업보고서 원문을 내려받아, "II. 사업의 내용"
섹션 텍스트만 추출해 캐시에 저장한다.

다른 스크립트들과 달리 이 데이터는 깔끔한 JSON 필드가 아니라 사업보고서 전체
문서(XML)에서 텍스트를 긁어내는 방식이라, 회사마다 문서 구조가 조금씩 달라
완벽하게 정제되지 않을 수 있다. 그래서 이 스크립트는 "완벽한 파싱"을 목표로
하지 않고, Claude가 읽고 요약할 수 있는 수준의 원문 텍스트 블록을 뽑아내는
것을 목표로 한다.

절차:
    1. list.json (공시검색)으로 최신 사업보고서의 접수번호(rcept_no)를 찾는다.
    2. document.xml (공시서류원본파일)로 그 보고서의 원문 ZIP을 받는다.
    3. ZIP 안의 XML에서 "사업의 내용" 또는 "II. 사업의 내용" 제목 이후 ~
       다음 대제목(로마숫자, 보통 "III.") 이전까지의 텍스트를 뽑는다.

사용법:
    python fetch_business_description.py --api-key <DART_API_KEY> <corp_code>

캐시 위치: cache/bizdesc_{corp_code}.json
    {"corp_name":..., "rcept_no":..., "rcept_dt":..., "text": "...", "truncated": bool}
"""
import argparse
import io
import json
import re
import sys
import zipfile
from pathlib import Path

import requests

BASE_URL = "https://opendart.fss.or.kr/api"
SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

# Claude가 읽고 요약하기에 충분하면서도 대화 컨텍스트를 과도하게 잡아먹지
# 않도록 텍스트 길이를 제한한다. 필요하면 사용자가 스크립트 인자로 늘릴 수 있다.
DEFAULT_MAX_CHARS = 12000


def get_api_key(cli_key: str | None) -> str:
    key = cli_key or __import__("os").environ.get("DART_API_KEY")
    if not key:
        print("ERROR: API 키가 없습니다. --api-key 인자로 전달하세요.", file=sys.stderr)
        sys.exit(1)
    return key


def find_latest_annual_report(api_key: str, corp_code: str) -> dict | None:
    """list.json으로 가장 최근 사업보고서(정기공시, 사업보고서 유형)의 접수번호를 찾는다."""
    resp = requests.get(
        f"{BASE_URL}/list.json",
        params={
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "pblntf_ty": "A",  # 정기공시
            "page_count": 20,
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("status") != "000":
        return None
    for item in data.get("list", []):
        # report_nm 예: "사업보고서 (2023.12)" — "사업보고서"로 시작하는 것만 (반기/분기 제외)
        if str(item.get("report_nm", "")).strip().startswith("사업보고서"):
            return item
    return None


def fetch_document_zip(api_key: str, rcept_no: str) -> bytes:
    resp = requests.get(
        f"{BASE_URL}/document.xml",
        params={"crtfc_key": api_key, "rcept_no": rcept_no},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.content


def strip_xml_tags(raw: str) -> str:
    # 표(TABLE) 안 셀 구분을 공백으로 살려두고 태그만 제거한다.
    text = re.sub(r"<[^>]+>", " ", raw)
    text = text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    text = text.replace("&nbsp;", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def extract_business_section(full_text: str) -> str | None:
    """"II. 사업의 내용" 제목부터 다음 로마숫자 대제목("III." 등) 전까지 추출한다."""
    # 문서마다 "사업의 내용", "II. 사업의 내용", "2. 사업의 내용" 등 표기가 다를 수 있다.
    start_pat = re.compile(r"(사업의\s*내용)")
    m_start = start_pat.search(full_text)
    if not m_start:
        return None
    tail = full_text[m_start.start():]
    # 다음 대제목(보통 "III. 재무에 관한 사항" 류) 앞까지만 자른다.
    end_pat = re.compile(r"(III\s*\.\s*재무|3\s*\.\s*재무에\s*관한|재무에\s*관한\s*사항)")
    m_end = end_pat.search(tail[50:])  # 자기 자신 근처 오탐 방지로 조금 건너뛰고 탐색
    if m_end:
        tail = tail[: 50 + m_end.start()]
    return tail.strip()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("corp_code")
    ap.add_argument("--api-key", required=True)
    ap.add_argument("--max-chars", type=int, default=DEFAULT_MAX_CHARS)
    args = ap.parse_args()

    report = find_latest_annual_report(args.api_key, args.corp_code)
    if not report:
        print(f"ERROR: {args.corp_code}의 사업보고서를 찾지 못했습니다.", file=sys.stderr)
        sys.exit(1)

    rcept_no = report["rcept_no"]
    try:
        zip_bytes = fetch_document_zip(args.api_key, rcept_no)
    except requests.RequestException as e:
        print(f"ERROR: 원문 다운로드 실패: {e}", file=sys.stderr)
        sys.exit(1)

    combined_text = ""
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if not name.lower().endswith((".xml", ".htm", ".html")):
                    continue
                raw = zf.read(name).decode("utf-8", errors="ignore")
                text = strip_xml_tags(raw)
                section = extract_business_section(text)
                if section:
                    combined_text = section
                    break
    except zipfile.BadZipFile:
        print("ERROR: 응답이 ZIP 형식이 아닙니다(인증키/접수번호 확인 필요).", file=sys.stderr)
        sys.exit(1)

    if not combined_text:
        print(
            "WARNING: '사업의 내용' 섹션을 자동으로 찾지 못했습니다. "
            "빈 텍스트로 저장하니, 필요하면 DART 사이트에서 직접 확인하세요.",
            file=sys.stderr,
        )

    truncated = len(combined_text) > args.max_chars
    if truncated:
        combined_text = combined_text[: args.max_chars]

    result = {
        "corp_code": args.corp_code,
        "rcept_no": rcept_no,
        "rcept_dt": report.get("rcept_dt"),
        "report_nm": report.get("report_nm"),
        "text": combined_text,
        "truncated": truncated,
    }

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_path = CACHE_DIR / f"bizdesc_{args.corp_code}.json"
    cache_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(
        {"saved": str(cache_path), "rcept_no": rcept_no, "text_length": len(combined_text), "truncated": truncated},
        ensure_ascii=False,
    ))


if __name__ == "__main__":
    main()
