#!/usr/bin/env python3
"""ticker_lookup.py — 기업명/Ticker를 SEC CIK(Central Index Key)로 변환한다.

v3.0.0: SEC EDGAR로 전환하면서 CIK 확보가 필수가 됐다. SEC가 공식으로 제공하는
전체 티커→CIK 매핑 파일(company_tickers.json, API 키 불필요)을 한 번 받아
캐시해두고 재사용한다 — 회사마다 개별 조회할 필요가 없다.

⚠️ SEC 요청 시 User-Agent 헤더에 "앱이름 연락처이메일" 형식이 반드시 필요하다
(SEC 정책 — 인증키는 아니지만 요청 시 빠지면 차단될 수 있다). 이 스크립트는
--contact 인자로 받거나 환경변수 SEC_CONTACT_EMAIL을 쓴다.
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"
TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"

# 자주 쓰는 기업의 한글/영문 라벨 → Ticker. SEC의 공식 매핑(전체 상장사)과
# 별개로, 사용자가 한글로 물어봤을 때 바로 대응하기 위한 보조 사전이다.
COMMON_TICKERS = {
    "애플": "AAPL", "apple": "AAPL",
    "마이크로소프트": "MSFT", "microsoft": "MSFT",
    "구글": "GOOGL", "google": "GOOGL", "알파벳": "GOOGL", "alphabet": "GOOGL",
    "아마존": "AMZN", "amazon": "AMZN",
    "메타": "META", "meta": "META", "페이스북": "META", "facebook": "META",
    "테슬라": "TSLA", "tesla": "TSLA",
    "엔비디아": "NVDA", "nvidia": "NVDA",
    "인텔": "INTC", "intel": "INTC",
    "amd": "AMD",
    "넷플릭스": "NFLX", "netflix": "NFLX",
    "어도비": "ADBE", "adobe": "ADBE",
    "세일즈포스": "CRM", "salesforce": "CRM",
    "오라클": "ORCL", "oracle": "ORCL",
    "IBM": "IBM", "아이비엠": "IBM",
    "퀄컴": "QCOM", "qualcomm": "QCOM",
    "브로드컴": "AVGO", "broadcom": "AVGO",
    "jp모건": "JPM", "jpmorgan": "JPM",
    "뱅크오브아메리카": "BAC", "bank of america": "BAC",
    "웰스파고": "WFC", "wells fargo": "WFC",
    "골드만삭스": "GS", "goldman sachs": "GS",
    "씨티그룹": "C", "citigroup": "C",
    "모건스탠리": "MS", "morgan stanley": "MS",
    "비자": "V", "visa": "V",
    "마스터카드": "MA", "mastercard": "MA",
    "페이팔": "PYPL", "paypal": "PYPL",
    "버크셔해서웨이": "BRK-B", "berkshire hathaway": "BRK-B",
    "코카콜라": "KO", "coca-cola": "KO", "coca cola": "KO",
    "펩시": "PEP", "pepsi": "PEP", "pepsico": "PEP",
    "맥도날드": "MCD", "mcdonald's": "MCD", "mcdonalds": "MCD",
    "나이키": "NKE", "nike": "NKE",
    "스타벅스": "SBUX", "starbucks": "SBUX",
    "디즈니": "DIS", "disney": "DIS",
    "코스트코": "COST", "costco": "COST",
    "월마트": "WMT", "walmart": "WMT",
    "존슨앤존슨": "JNJ", "johnson & johnson": "JNJ",
    "화이자": "PFE", "pfizer": "PFE",
    "보잉": "BA", "boeing": "BA",
    "엑슨모빌": "XOM", "exxonmobil": "XOM", "exxon": "XOM",
}


def _explain_error(e: Exception) -> str:
    msg = str(e)
    if "403" in msg:
        return f"{msg} — SEC가 요청을 차단했습니다. User-Agent 헤더(연락처 이메일 포함)가 올바른지 확인하세요."
    if "429" in msg:
        return f"{msg} — SEC 레이트리밋(초당 10회 제한). 요청 간격을 늘리세요."
    return msg


def _headers(contact_email: str) -> dict:
    if not contact_email or "@" not in contact_email:
        raise ValueError(
            "SEC EDGAR는 User-Agent에 유효한 연락처 이메일이 필요합니다. "
            "--contact your@email.com 형태로 전달하세요(SEC 정책)."
        )
    return {"User-Agent": f"us-stocks-financials-plugin {contact_email}"}


def fetch_all_tickers(contact_email: str, cache_dir: Path = CACHE_DIR) -> dict:
    """SEC의 전체 티커→CIK 매핑을 받아 캐시한다(하루에 한 번 정도면 충분 —
    상장 변동이 매일 나는 게 아니므로 캐시가 있으면 재사용한다)."""
    import requests

    cache_file = cache_dir / "sec_company_tickers.json"
    if cache_file.exists():
        try:
            cached = json.loads(cache_file.read_text(encoding="utf-8"))
            return cached
        except Exception:
            pass

    try:
        resp = requests.get(TICKERS_URL, headers=_headers(contact_email), timeout=30)
        resp.raise_for_status()
        raw = resp.json()
    except Exception as e:
        print(f"FAIL SEC 티커 매핑 다운로드 실패: {_explain_error(e)}", file=sys.stderr)
        return {}

    # raw 형태: {"0": {"cik_str": 320193, "ticker": "AAPL", "title": "Apple Inc."}, ...}
    by_ticker = {}
    for entry in raw.values():
        t = entry.get("ticker", "").upper()
        if t:
            by_ticker[t] = {"cik": str(entry["cik_str"]).zfill(10), "title": entry.get("title", "")}

    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file.write_text(json.dumps(by_ticker, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"SAVE SEC 티커 매핑 {len(by_ticker)}개 저장: {cache_file}")
    return by_ticker


def lookup_ticker(company_name: str) -> str | None:
    normalized = company_name.strip().lower()
    return COMMON_TICKERS.get(normalized)


def resolve_cik(ticker: str, contact_email: str) -> tuple[str | None, str | None]:
    """Ticker로 CIK와 공식 회사명을 찾는다. 반환: (cik, title) 또는 (None, None)."""
    all_tickers = fetch_all_tickers(contact_email)
    entry = all_tickers.get(ticker.upper())
    if entry:
        return entry["cik"], entry["title"]
    return None, None


def save_company_info(ticker: str, cik: str, title: str, cache_dir: Path = CACHE_DIR):
    cache_dir.mkdir(parents=True, exist_ok=True)
    company_info = {
        "ticker": ticker.upper(), "cik": cik, "title": title,
        "fetched_at": datetime.now().isoformat(),
    }
    cache_file = cache_dir / f"company_{ticker.upper()}.json"
    cache_file.write_text(json.dumps(company_info, ensure_ascii=False, indent=2), encoding="utf-8")
    return company_info


def main() -> None:
    ap = argparse.ArgumentParser(description="기업명/Ticker → SEC CIK 변환")
    ap.add_argument("company_name", help="기업명(한글/영문) 또는 Ticker 심볼")
    ap.add_argument("--contact", required=True, help="SEC User-Agent에 넣을 연락처 이메일(필수, SEC 정책)")
    args = ap.parse_args()

    ticker = lookup_ticker(args.company_name)
    if ticker is None:
        ticker = args.company_name.strip().upper()

    cik, title = resolve_cik(ticker, args.contact)
    if cik is None:
        print(
            f"FAIL '{args.company_name}'(Ticker: {ticker})을(를) SEC 매핑에서 찾을 수 없습니다. "
            "정확한 Ticker 심볼을 확인해주세요.",
            file=sys.stderr,
        )
        print(json.dumps({"found": False, "input": args.company_name}, ensure_ascii=False))
        sys.exit(1)

    save_company_info(ticker, cik, title)
    print(f"OK {args.company_name} -> {ticker} (CIK {cik}, {title})")
    print(json.dumps({"found": True, "ticker": ticker, "cik": cik, "title": title}, ensure_ascii=False))


if __name__ == "__main__":
    main()
