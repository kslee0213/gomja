#!/usr/bin/env python3
"""ticker_lookup.py — 자주 쓰는 기업명(한글/영문)을 Ticker로 변환한다.

⚠️ v1.0.0은 문서에 "200+ 기업 자동 매핑"이라고 적어놓고 실제 매핑 사전은
11개뿐이었다(문서-코드 불일치). 이 버전은 매핑 목록을 정직하게 이 파일
안의 COMMON_TICKERS 개수 그대로로 유지하고, 여기 없는 기업은 무조건
"Ticker를 직접 입력하세요"로 안내한다 — 억지로 야후 파이낸스 검색 API를
붙여 "혹시 맞을 수도 있는" 티커를 추측하지 않는다(잘못된 회사 데이터를
가져오는 것보다 명확히 실패하는 게 낫다).
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

# 자주 쓰는 기업만 담는다. 개수를 부풀리지 않는다 — 여기 없으면 사용자가
# Ticker를 직접 입력하도록 명확히 안내하는 편이, "아마 이거겠지" 하고
# 틀린 회사를 매칭하는 것보다 훨씬 안전하다.
COMMON_TICKERS = {
    # 빅테크
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
    # 금융
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
    # 소비재/기타
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


def lookup_ticker(company_name: str) -> str | None:
    """알려진 매핑에 있으면 티커를 반환하고, 없으면 None(실패)을 반환한다.
    "5자 이내 영문이면 그냥 티커로 간주" 같은 추측은 하지 않는다 — 존재하지
    않는 티커를 그대로 넘겨서 나중 단계에서 원인 불명확한 에러가 나느니,
    여기서 명확히 실패시키는 편이 사용자 경험상 낫다(이후 단계에서
    ticker_obj.info로 실존 여부를 검증하므로 그 결과가 훨씬 명확하다)."""
    normalized = company_name.strip().lower()
    return COMMON_TICKERS.get(normalized)


def verify_ticker_exists(ticker: str) -> tuple[bool, dict | None]:
    """yfinance로 해당 티커가 실제로 존재하는지 확인한다. 매핑에 없는
    기업이거나, 사용자가 티커를 직접 입력한 경우 둘 다 이 함수로 검증한다."""
    import yfinance as yf

    try:
        t = yf.Ticker(ticker.upper())
        info = t.info
        if info and info.get("longName"):
            return True, info
        return False, None
    except Exception as e:
        print(f"WARN {ticker} 조회 중 오류: {e}", file=sys.stderr)
        return False, None


def save_company_info(ticker: str, info: dict, cache_dir: Path = CACHE_DIR):
    cache_dir.mkdir(parents=True, exist_ok=True)
    company_info = {
        "ticker": ticker,
        "longName": info.get("longName", ""),
        "sector": info.get("sector", ""),
        "industry": info.get("industry", ""),
        "currency": info.get("currency", ""),
        "exchange": info.get("exchange", ""),
        "fetched_at": datetime.now().isoformat(),
    }
    cache_file = cache_dir / f"company_{ticker}.json"
    cache_file.write_text(json.dumps(company_info, ensure_ascii=False, indent=2), encoding="utf-8")
    return company_info


def main() -> None:
    ap = argparse.ArgumentParser(description="기업명 → Ticker 변환")
    ap.add_argument("company_name", help="기업명(한글/영문) 또는 Ticker 심볼")
    args = ap.parse_args()

    ticker = lookup_ticker(args.company_name)
    if ticker is None:
        # 매핑에 없으면 사용자가 준 문자열 자체를 티커 후보로 보고 실존 여부만 검증한다.
        ticker = args.company_name.strip().upper()

    ok, info = verify_ticker_exists(ticker)
    if not ok:
        print(
            f"FAIL '{args.company_name}'을(를) 찾을 수 없습니다. "
            f"정확한 Ticker 심볼을 직접 입력해주세요(예: AAPL).",
            file=sys.stderr,
        )
        print(json.dumps({"found": False, "input": args.company_name}, ensure_ascii=False))
        sys.exit(1)

    save_company_info(ticker, info)
    print(f"OK {args.company_name} -> {ticker} ({info.get('longName', '')})")
    print(json.dumps({"found": True, "ticker": ticker, "longName": info.get("longName", "")}, ensure_ascii=False))


if __name__ == "__main__":
    main()
