#!/usr/bin/env python3
"""fetch_extra_info.py — 현재 주가·시가총액만 yfinance로 최소한 가져온다.

v3.0.0: 재무제표는 전부 SEC EDGAR로 옮겼으므로, yfinance는 SEC가 제공하지
않는 "시장 주가" 하나만 위해 남겨뒀다. 의존 범위를 최소화해서, yfinance가
또 막히더라도(Yahoo의 403/429는 여전히 발생할 수 있음) 재무제표 자체는
영향받지 않고 "투자분석"의 주가 연동 지표(PER·PBR 등)만 비게 만든다
(그 부분만 실패해도 전체 워크북이 죽지 않도록 build_workbook.py에서 처리).
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"


def _explain_error(e: Exception) -> str:
    msg = str(e)
    if "403" in msg:
        return f"{msg} — Yahoo Finance 403 차단(클라우드 IP대역 봇 차단 가능성). 재무제표(SEC EDGAR)는 영향 없음, 주가만 빈다."
    if "429" in msg:
        return f"{msg} — Yahoo Finance 레이트리밋. 잠시 후 재시도."
    return msg


def fetch_price_info(ticker: str) -> dict:
    """딱 4개 필드만 가져온다: 현재가, 시가총액, 상장주식수, 52주 고저.
    PER/PBR/PSR은 SEC 재무제표 값(EPS·BVPS·매출)과 이 주가를 조합해
    build_workbook.py가 직접 수식으로 계산한다 — yfinance의 trailingPE 같은
    "이미 계산된 값"을 그대로 믿지 않고 우리가 재계산해 감사 가능하게 만든다."""
    import yfinance as yf

    result = {"ticker": ticker, "price": {}, "errors": []}
    try:
        t = yf.Ticker(ticker)
        info = t.info or {}
        result["price"] = {
            "currentPrice": info.get("currentPrice") or info.get("regularMarketPrice"),
            "marketCap": info.get("marketCap"),
            "sharesOutstanding": info.get("sharesOutstanding"),
            "fiftyTwoWeekHigh": info.get("fiftyTwoWeekHigh"),
            "fiftyTwoWeekLow": info.get("fiftyTwoWeekLow"),
            "dividendYield": info.get("dividendYield"),
            "payoutRatio": info.get("payoutRatio"),
        }
    except Exception as e:
        result["errors"].append(_explain_error(e))
        print(f"WARN {ticker} 주가 조회 실패: {_explain_error(e)}", file=sys.stderr)
    return result


def save_price_cache(data: dict, cache_dir: Path = CACHE_DIR):
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / f"price_{data['ticker'].upper()}.json"
    payload = dict(data)
    payload["fetched_at"] = datetime.now().isoformat()
    cache_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"SAVE 캐시 저장: {cache_file}")
    return cache_file


def load_price_cache(ticker: str, cache_dir: Path = CACHE_DIR):
    cache_file = cache_dir / f"price_{ticker.upper()}.json"
    if not cache_file.exists():
        return None
    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except Exception:
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description="현재 주가·시가총액 수집(yfinance, 최소 사용)")
    ap.add_argument("ticker")
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    cached = None if args.force else load_price_cache(args.ticker)
    if cached:
        print(f"CACHE 캐시 재사용: {args.ticker}")
        print(json.dumps({"cached": True, "ticker": args.ticker}, ensure_ascii=False))
        return

    data = fetch_price_info(args.ticker)
    save_price_cache(data)
    print(json.dumps({"cached": False, "ticker": args.ticker, "errors": data["errors"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
