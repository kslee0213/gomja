#!/usr/bin/env python3
"""fetch_extra_info.py — 현재 주가·배당·섹터·개황 정보를 수집해 캐시에 저장한다.

v1.0.0의 bare `except:` 를 전부 없애고, 실패 시 어떤 필드를 못 가져왔는지
stderr에 정확히 남긴다(조용히 빈 dict를 돌려주면 나중에 investment_analysis에서
"왜 비었는지" 추적이 안 된다).
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"

# yfinance .info 딕셔너리에서 쓸 필드. 값이 없으면 None으로 남기고 지어내지 않는다.
INFO_FIELDS = {
    "longName": "longName",
    "shortName": "shortName",
    "sector": "sector",
    "industry": "industry",
    "currency": "currency",
    "exchange": "exchange",
    "currentPrice": "currentPrice",
    "regularMarketPrice": "regularMarketPrice",
    "previousClose": "previousClose",
    "marketCap": "marketCap",
    "sharesOutstanding": "sharesOutstanding",
    "trailingPE": "trailingPE",
    "forwardPE": "forwardPE",
    "priceToBook": "priceToBook",
    "priceToSalesTrailing12Months": "priceToSalesTrailing12Months",
    "trailingEps": "trailingEps",
    "bookValue": "bookValue",
    "dividendYield": "dividendYield",
    "dividendRate": "dividendRate",
    "payoutRatio": "payoutRatio",
    "returnOnEquity": "returnOnEquity",
    "returnOnAssets": "returnOnAssets",
    "debtToEquity": "debtToEquity",
    "beta": "beta",
    "fiftyTwoWeekHigh": "fiftyTwoWeekHigh",
    "fiftyTwoWeekLow": "fiftyTwoWeekLow",
}


def fetch_extra_info(ticker: str) -> dict:
    import yfinance as yf

    result = {"ticker": ticker, "info": {}, "dividends": {}, "errors": []}

    try:
        t = yf.Ticker(ticker)
        info = t.info or {}
        for out_key, src_key in INFO_FIELDS.items():
            result["info"][out_key] = info.get(src_key)
    except Exception as e:
        result["errors"].append(f"info 조회 실패: {e}")
        print(f"WARN {ticker} info 조회 실패: {e}", file=sys.stderr)

    try:
        t = yf.Ticker(ticker)
        div = t.dividends
        if div is not None and not div.empty:
            result["dividends"] = {
                "last_dividend": float(div.iloc[-1]),
                "last_ex_date": str(div.index[-1].date()),
                "trailing_4_sum": float(div.tail(4).sum()),
            }
        else:
            result["dividends"] = {"no_dividends": True}
    except Exception as e:
        result["errors"].append(f"dividends 조회 실패: {e}")
        print(f"WARN {ticker} dividends 조회 실패: {e}", file=sys.stderr)

    return result


def save_extra_info_cache(data: dict, cache_dir: Path = CACHE_DIR):
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"extra_info_{data['ticker']}.json"
    payload = dict(data)
    payload["fetched_at"] = datetime.now().isoformat()
    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"SAVE 캐시 저장: {cache_file}")
    return cache_file


def load_extra_info_cache(ticker: str, cache_dir: Path = CACHE_DIR):
    cache_file = cache_dir / f"extra_info_{ticker}.json"
    if not cache_file.exists():
        return None
    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except Exception:
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description="주가/배당/섹터 정보 수집")
    ap.add_argument("ticker")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    cached = None if args.force else load_extra_info_cache(args.ticker)
    if cached:
        print(f"CACHE 캐시 재사용: {args.ticker}")
        print(json.dumps({"cached": True, "ticker": args.ticker}, ensure_ascii=False))
        return

    data = fetch_extra_info(args.ticker)
    save_extra_info_cache(data)
    print(json.dumps({"cached": False, "ticker": args.ticker, "errors": data["errors"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
