#!/usr/bin/env python3
"""fetch_financials.py — yfinance로 재무제표(BS/IS/CF)를 수집해 캐시에 저장한다.

⚠️ v2.0.0 재작성 시 실제로 고친 버그(v1.0.0 기준):
  - quarterly/annual 캐시 파일명이 동일해서(`financials_{ticker}_{날짜}.json`)
    --period both로 실행하면 나중 것이 먼저 것을 덮어써 분기 데이터가
    통째로 사라지는 버그가 있었다. 파일명에 frequency를 반드시 포함한다.
  - DataFrame.to_dict()의 결과에는 pandas Timestamp(컬럼 인덱스)와 numpy
    타입(np.float64, np.int64, NaN)이 섞여 있어 표준 json.dump()가
    "Object of type Timestamp is not JSON serializable" 로 실패한다.
    반드시 이 파일의 _json_safe()를 거쳐서 저장한다.
  - yfinance 무료 데이터는 분기 4~5개, 연간 4개 정도만 제공되는 게 보통이다
    (한국 DART처럼 12분기/5개년을 항상 다 채울 수 있다고 가정하면 안 된다).
    "있는 만큼만" 채우고, 사용자에게 몇 개나 채웠는지 정직하게 보고한다.
"""
import argparse
import json
import math
import sys
from datetime import datetime, date
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_DIR = SCRIPT_DIR.parent / "cache"


def _json_safe(obj):
    """pandas/numpy 값을 표준 json이 직렬화 가능한 형태로 재귀 변환한다."""
    if obj is None:
        return None
    if isinstance(obj, dict):
        return {_json_safe_key(k): _json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_json_safe(v) for v in obj]
    if isinstance(obj, (datetime, date)):
        return obj.strftime("%Y-%m-%d")
    if hasattr(obj, "isoformat"):  # pandas.Timestamp 등
        try:
            return obj.strftime("%Y-%m-%d")
        except Exception:
            return str(obj)
    if isinstance(obj, float):
        return None if math.isnan(obj) else float(obj)
    # numpy 스칼라(np.float64, np.int64, np.bool_ 등): .item()으로 파이썬 기본형 변환
    if hasattr(obj, "item") and not isinstance(obj, (str, bytes)):
        try:
            v = obj.item()
            return None if (isinstance(v, float) and math.isnan(v)) else v
        except Exception:
            pass
    return obj


def _json_safe_key(k):
    if hasattr(k, "strftime"):
        try:
            return k.strftime("%Y-%m-%d")
        except Exception:
            return str(k)
    return str(k)


def _df_to_records(df):
    """DataFrame(행=계정명, 열=기간)을 {계정명: {날짜문자열: 값}} 형태로 변환한다.
    빈 DataFrame이면 {}를 반환한다."""
    if df is None or df.empty:
        return {}
    out = {}
    for account in df.index:
        row = {}
        for col in df.columns:
            val = df.loc[account, col]
            row[_json_safe_key(col)] = _json_safe(val)
        out[str(account)] = row
    return out


def fetch_quarterly_financials(ticker: str, periods: int = 12):
    """분기 재무제표 수집. yfinance가 실제로 제공하는 개수가 12보다 적을 수 있다
    (보통 4~5개) — 있는 만큼만 가져온다."""
    import yfinance as yf

    try:
        t = yf.Ticker(ticker)
        q_income = t.quarterly_financials
        q_balance = t.quarterly_balance_sheet
        q_cashflow = t.quarterly_cashflow

        n_available = max(
            q_income.shape[1] if q_income is not None else 0,
            q_balance.shape[1] if q_balance is not None else 0,
            q_cashflow.shape[1] if q_cashflow is not None else 0,
        )
        n = min(periods, n_available) if n_available else 0

        data = {
            "ticker": ticker,
            "frequency": "quarterly",
            "requested_periods": periods,
            "available_periods": n_available,
            "income_statement": _df_to_records(q_income.iloc[:, :n] if q_income is not None and not q_income.empty else None),
            "balance_sheet": _df_to_records(q_balance.iloc[:, :n] if q_balance is not None and not q_balance.empty else None),
            "cash_flow": _df_to_records(q_cashflow.iloc[:, :n] if q_cashflow is not None and not q_cashflow.empty else None),
        }
        print(f"OK 분기 재무제표 수집 완료: {ticker} ({n_available}개 분기 확보, {periods}개 요청)")
        return data
    except Exception as e:
        print(f"FAIL 분기 재무제표 수집 실패: {ticker} — {e}", file=sys.stderr)
        return None


def fetch_annual_financials(ticker: str, periods: int = 5):
    """연간 재무제표 수집. yfinance는 보통 최근 4개년 정도만 제공한다."""
    import yfinance as yf

    try:
        t = yf.Ticker(ticker)
        a_income = t.financials
        a_balance = t.balance_sheet
        a_cashflow = t.cashflow

        n_available = max(
            a_income.shape[1] if a_income is not None else 0,
            a_balance.shape[1] if a_balance is not None else 0,
            a_cashflow.shape[1] if a_cashflow is not None else 0,
        )
        n = min(periods, n_available) if n_available else 0

        data = {
            "ticker": ticker,
            "frequency": "annual",
            "requested_periods": periods,
            "available_periods": n_available,
            "income_statement": _df_to_records(a_income.iloc[:, :n] if a_income is not None and not a_income.empty else None),
            "balance_sheet": _df_to_records(a_balance.iloc[:, :n] if a_balance is not None and not a_balance.empty else None),
            "cash_flow": _df_to_records(a_cashflow.iloc[:, :n] if a_cashflow is not None and not a_cashflow.empty else None),
        }
        print(f"OK 연간 재무제표 수집 완료: {ticker} ({n_available}개년 확보, {periods}개 요청)")
        return data
    except Exception as e:
        print(f"FAIL 연간 재무제표 수집 실패: {ticker} — {e}", file=sys.stderr)
        return None


def save_financials_cache(data: dict, cache_dir: Path = CACHE_DIR):
    """frequency(quarterly/annual)를 반드시 파일명에 포함해 서로 덮어쓰지 않게 한다.
    날짜 대신 데이터 자체를 식별하는 키만 쓰므로, 같은 날 재실행하면 캐시를 그대로 재사용한다."""
    cache_dir.mkdir(parents=True, exist_ok=True)
    ticker = data["ticker"]
    freq = data["frequency"]  # "quarterly" | "annual"
    cache_file = cache_dir / f"financials_{ticker}_{freq}.json"
    payload = dict(data)
    payload["fetched_at"] = datetime.now().isoformat()
    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"SAVE 캐시 저장: {cache_file}")
    return cache_file


def load_financials_cache(ticker: str, frequency: str, cache_dir: Path = CACHE_DIR):
    cache_file = cache_dir / f"financials_{ticker}_{frequency}.json"
    if not cache_file.exists():
        return None
    try:
        return json.loads(cache_file.read_text(encoding="utf-8"))
    except Exception:
        return None


def main() -> None:
    ap = argparse.ArgumentParser(description="yfinance 재무제표 수집")
    ap.add_argument("ticker")
    ap.add_argument("--period", choices=["quarterly", "annual", "both"], default="both")
    ap.add_argument("--quarters", type=int, default=12)
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--force", action="store_true", help="캐시가 있어도 강제로 다시 조회한다")
    args = ap.parse_args()

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    result = {"ticker": args.ticker}

    if args.period in ("quarterly", "both"):
        cached = None if args.force else load_financials_cache(args.ticker, "quarterly")
        if cached:
            print(f"CACHE 분기 캐시 재사용: {args.ticker} ({cached.get('available_periods')}개 분기)")
            result["quarterly"] = {"saved": True, "available_periods": cached.get("available_periods")}
        else:
            q = fetch_quarterly_financials(args.ticker, args.quarters)
            if q:
                save_financials_cache(q)
                result["quarterly"] = {"saved": True, "available_periods": q["available_periods"]}
            else:
                result["quarterly"] = {"saved": False}

    if args.period in ("annual", "both"):
        cached = None if args.force else load_financials_cache(args.ticker, "annual")
        if cached:
            print(f"CACHE 연간 캐시 재사용: {args.ticker} ({cached.get('available_periods')}개년)")
            result["annual"] = {"saved": True, "available_periods": cached.get("available_periods")}
        else:
            a = fetch_annual_financials(args.ticker, args.years)
            if a:
                save_financials_cache(a)
                result["annual"] = {"saved": True, "available_periods": a["available_periods"]}
            else:
                result["annual"] = {"saved": False}

    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
