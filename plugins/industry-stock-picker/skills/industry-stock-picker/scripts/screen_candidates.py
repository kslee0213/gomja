#!/usr/bin/env python3
"""
industry-stock-picker: 후보 기업들을 일괄 수집·평가해 산업 스크리닝 순위를 만든다.

사용:
    python screen_candidates.py candidates.json \
        [--fetch] [--dart-key <키>] [--krx-key <키>] [--contact <이메일>] \
        [--assumptions-kr a_kr.json] [--assumptions-us a_us.json] \
        [--years 5] [--out screening.json]

candidates.json 스키마는 examples/candidates.example.json 참고.
- KR 후보는 corp_code + stock_code, US 후보는 ticker + cik(10자리)가 미리 채워져 있어야
  한다(식별은 사람이/Claude가 corp_code_lookup·ticker_lookup으로 확정 — 스크립트가
  이름만 보고 다른 회사를 집어오는 사고 방지).
- --fetch를 주면 기존 추출 플러그인 스크립트를 호출해 캐시를 먼저 채운다(있으면 재사용).
- API 키는 인자로만 받고 출력·저장물에 남기지 않는다.

출력(screening.json): 후보별 지표·가치평가 요약·점수 분해·결측/경고와 순위.
점수 규칙(정규화): 각 항목은 (점수, 만점)으로 계산하되 데이터가 없는 항목은 0점 처리가
아니라 분모에서 제외하고 '결측'으로 기록한다. 종합점수 = Σ점수/Σ가용만점 × 100 − 감점.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PLUGINS_DIR = SCRIPT_DIR.parents[3]          # .../plugins
DART_SKILL = PLUGINS_DIR / "dart-kospi-financials" / "skills" / "dart-financial-extractor"
US_SKILL = PLUGINS_DIR / "us-stocks-financials" / "skills" / "us-stock-financial-extractor"
VV_SCRIPTS = PLUGINS_DIR / "valuation-verdict" / "skills" / "valuation-verdict" / "scripts"
sys.path.insert(0, str(VV_SCRIPTS))

import adapters  # noqa: E402
import valuation_verdict as vv  # noqa: E402

REPRT_QUARTERS = ["11013", "11012", "11014"]  # 1분기·반기·3분기


def _run(cmd: list[str], label: str, timeout: int = 120) -> tuple[int, str]:
    """하위 스크립트 실행. 키가 섞인 전체 명령은 출력하지 않는다."""
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return p.returncode, (p.stdout or "") + (p.stderr or "")
    except subprocess.TimeoutExpired:
        return 124, f"{label}: timeout"


# --------------------------------------------------------------------------
# 수집 (--fetch)
# --------------------------------------------------------------------------

def fetch_kr(c: dict, dart_key: str | None, krx_key: str | None, years: int, log: list[str]) -> None:
    if not dart_key:
        log.append("KR: --dart-key가 없어 수집 생략(캐시만 사용)")
        return
    corp, stock = c["corp_code"], c.get("stock_code", "")
    fetch_py = str(DART_SKILL / "scripts" / "fetch_financials.py")
    this_year = dt.date.today().year
    ok = 0
    for y in range(this_year - years, this_year):
        rc, out = _run([sys.executable, fetch_py, "--api-key", dart_key, corp, str(y), "11011"], f"dart {y}")
        ok += 1 if '"status": "000"' in out else 0
    log.append(f"연간 사업보고서 수집: {ok}/{years}개년")
    # 진행연도 분기/반기(있는 것만) + 전년 동기(추정(E) 계산용)
    got_q = []
    for code in REPRT_QUARTERS:
        rc, out = _run([sys.executable, fetch_py, "--api-key", dart_key, corp, str(this_year), code], f"dart q{code}")
        if '"status": "000"' in out:
            got_q.append(code)
            _run([sys.executable, fetch_py, "--api-key", dart_key, corp, str(this_year - 1), code], f"dart q-1 {code}")
    if got_q:
        log.append(f"진행연도 보고서: {','.join(got_q)}")
    # 배당 등 부가 공시(실패해도 계속)
    extra_py = str(DART_SKILL / "scripts" / "fetch_extra_disclosures.py")
    _run([sys.executable, extra_py, "--api-key", dart_key, corp, str(this_year - 1), "11011"], "dart extra")
    # 주가(KRX)
    if krx_key and stock:
        price_py = str(DART_SKILL / "scripts" / "fetch_stock_price.py")
        today = dt.date.today().strftime("%Y%m%d")
        rc, out = _run([sys.executable, price_py, stock, today, "--auth-key", krx_key, "--max-back", "10"], "krx price")
        log.append("주가 수집: 성공" if rc == 0 else "주가 수집 실패(현재가 없이 진행 — 판정 불가)")
    elif not krx_key:
        log.append("KR: --krx-key가 없어 주가 수집 생략")


def fetch_us(c: dict, contact: str | None, log: list[str]) -> None:
    if not contact:
        log.append("US: --contact가 없어 수집 생략(캐시만 사용)")
        return
    ticker = c["ticker"].upper()
    cik = c.get("cik")
    if not cik:
        rc, out = _run([sys.executable, str(US_SKILL / "scripts" / "ticker_lookup.py"), ticker, "--contact", contact], "ticker_lookup")
        try:
            found = json.loads(out.strip().splitlines()[-1])
            cik = found.get("cik")
            c["cik"] = cik
        except (json.JSONDecodeError, IndexError):
            log.append("ticker_lookup 실패 — cik를 candidates.json에 직접 넣어주세요")
            return
    rc, out = _run([sys.executable, str(US_SKILL / "scripts" / "fetch_financials.py"), ticker, cik, "--contact", contact], "sec fetch", timeout=180)
    log.append("SEC 재무제표 수집: " + ("성공" if rc == 0 else f"실패 rc={rc}"))
    rc, out = _run([sys.executable, str(US_SKILL / "scripts" / "fetch_extra_info.py"), ticker], "yfinance", timeout=120)
    log.append("주가(yfinance) 수집: " + ("성공" if rc == 0 else "실패(현재가 없이 진행 — 판정 불가)"))


# --------------------------------------------------------------------------
# 지표·점수
# --------------------------------------------------------------------------

def _avg(vals: list[float | None], n: int, min_count: int = 2) -> float | None:
    xs = [v for v in vals[-n:] if v is not None]
    return sum(xs) / len(xs) if len(xs) >= min_count else None


def compute_metrics(fi, res: dict) -> dict:
    ni_p = vv._parent(fi, "ni")
    eq_p = vv._parent(fi, "equity")
    liab = fi.series("liabilities")
    ocf = fi.series("ocf")
    capex = fi.series("capex")
    # 마지막 실적 연도 인덱스((E) 추정연도는 배수 계산에서 제외)
    last = len(fi.years) - (2 if fi.estimated_last else 1)
    if last < 0:
        last = 0
    m: dict = {"basis_year": fi.years[last] if fi.years else None}

    # ROE 3년 평균(%, 지배 우선)
    roes = [ni / eq * 100 for ni, eq in zip(ni_p, eq_p) if ni is not None and eq]
    m["roe3_pct"] = _avg(roes, 3) if roes else None
    # 부채비율(%, 최근 실적 연도)
    m["debt_ratio_pct"] = (liab[last] / eq_p[last] * 100) if (last < len(liab) and liab[last] is not None and eq_p[last]) else None
    # 이익의 질: 최근 3개 실적연도 ΣOCF/Σ순이익(순이익 합>0일 때만)
    pairs = [(o, n) for o, n in list(zip(ocf, ni_p))[: last + 1] if o is not None and n is not None][-3:]
    m["ocf_ni3"] = (sum(o for o, _ in pairs) / sum(n for _, n in pairs)) if pairs and sum(n for _, n in pairs) > 0 else None
    # FCF 양수 연도 비중
    fcf = [(o - c) for o, c in list(zip(ocf, capex))[: last + 1] if o is not None and c is not None]
    m["fcf_pos_share"] = (sum(1 for v in fcf if v > 0) / len(fcf)) if len(fcf) >= 3 else None
    m["fcf_years"] = len(fcf)
    # 매출 성장(3년 CAGR, %) — 엔진이 계산한 값 재사용
    rc = res.get("params", {}).get("rev_cagr")
    m["rev_cagr3_pct"] = rc * 100 if rc is not None else None
    # 흑자 연속성(최근 5개 실적연도)
    nis = [v for v in ni_p[: last + 1] if v is not None][-5:]
    m["loss_years"] = sum(1 for v in nis if v <= 0) if len(nis) >= 3 else None
    m["ni_years"] = len(nis)
    # 현재 배수(마지막 실적 연도 기준)
    if fi.price and fi.shares:
        mcap = fi.price * fi.shares
        ni_l = ni_p[last] if last < len(ni_p) else None
        eq_l = eq_p[last] if last < len(eq_p) else None
        m["per"] = mcap / (ni_l * fi.amount_unit) if ni_l and ni_l > 0 else None
        m["pbr"] = mcap / (eq_l * fi.amount_unit) if eq_l and eq_l > 0 else None
        m["graham_mult"] = m["per"] * m["pbr"] if (m.get("per") and m.get("pbr")) else None
    else:
        m["per"] = m["pbr"] = m["graham_mult"] = None
    m["dividend_yield_pct"] = res.get("dividend_yield_pct")
    return m


def _lin(v: float, lo: float, hi: float, maxpts: float) -> float:
    """lo 이하 0점, hi 이상 만점, 사이 선형."""
    if hi == lo:
        return maxpts
    t = (v - lo) / (hi - lo)
    return maxpts * min(1.0, max(0.0, t))


def score_candidate(m: dict, res: dict) -> dict:
    comp: list[dict] = []

    def add(name: str, pts: float | None, maxpts: float, note: str = ""):
        comp.append({"name": name, "points": round(pts, 2) if pts is not None else None,
                     "max": maxpts, "missing": pts is None, "note": note})

    up = res.get("upside_pct")
    add("가치(상승여력)", _lin(up, -30, 50, 40) if up is not None else None, 40,
        "가치평가_결론 기준 상승여력 −30%→0점, +50%→만점")
    add("수익성(ROE 3평균)", _lin(m["roe3_pct"], 0, 15, 15) if m.get("roe3_pct") is not None else None, 15,
        "0%→0점, 15% 이상 만점")
    add("건전성(부채비율)", _lin(-m["debt_ratio_pct"], -300, -100, 10) if m.get("debt_ratio_pct") is not None else None, 10,
        "100% 이하 만점, 300% 이상 0점")
    add("이익의 질(OCF/NI)", _lin(m["ocf_ni3"], 0.4, 1.0, 8) if m.get("ocf_ni3") is not None else None, 8,
        "0.4→0점, 1.0 이상 만점")
    add("현금창출(FCF 양수비중)", m["fcf_pos_share"] * 7 if m.get("fcf_pos_share") is not None else None, 7,
        f"FCF 계산 가능 {m.get('fcf_years', 0)}개년 중 양수 비중")
    add("성장(매출 3년 CAGR)", _lin(m["rev_cagr3_pct"], 0, 10, 10) if m.get("rev_cagr3_pct") is not None else None, 10,
        "0%→0점, 10% 이상 만점")
    if m.get("loss_years") is not None:
        pts = 10.0 if m["loss_years"] == 0 else (4.0 if m["loss_years"] == 1 else 0.0)
        add("흑자 연속성", pts, 10, f"최근 {m['ni_years']}개 실적연도 중 적자 {m['loss_years']}회")
    else:
        add("흑자 연속성", None, 10, "실적연도 3개 미만")

    avail = [c for c in comp if not c["missing"]]
    max_avail = sum(c["max"] for c in avail)
    raw = sum(c["points"] for c in avail)
    normalized = raw / max_avail * 100 if max_avail else None

    penalty = 0.0
    reasons = []
    if res.get("confidence") == "낮음":
        penalty += 8; reasons.append("가치평가 신뢰도 낮음 −8")
    elif res.get("confidence") == "보통":
        penalty += 4; reasons.append("가치평가 신뢰도 보통 −4")
    if res.get("price") is None:
        reasons.append("현재가 없음 — 가치 항목 결측(감점 대신 분모 제외)")
    final = (normalized - penalty) if normalized is not None else None
    return {"components": comp, "available_max": max_avail, "raw": round(raw, 2),
            "normalized": round(normalized, 1) if normalized is not None else None,
            "penalty": penalty, "penalty_reasons": reasons,
            "final": round(final, 1) if final is not None else None,
            "missing": [c["name"] for c in comp if c["missing"]]}


# --------------------------------------------------------------------------
# 메인
# --------------------------------------------------------------------------

def evaluate_candidate(c: dict, args) -> dict:
    market = c["market"].upper()
    if market == "KR":
        fi = adapters.load_dart(str(DART_SKILL / "cache"), c["corp_code"], args.years)
        a = vv.Assumptions.from_file(args.assumptions_kr)
    elif market == "US":
        fi = adapters.load_sec(str(US_SKILL / "cache"), c["ticker"], args.years)
        a = vv.Assumptions.from_file(args.assumptions_us)
    elif market == "JSON":  # 테스트/외부 연동용 정규화 입력
        fi = adapters.load_json(c["input"])
        a = vv.Assumptions.from_file(c.get("assumptions"))
    else:
        raise ValueError(f"market은 KR/US/JSON만 지원: {market}")
    res = vv.evaluate(fi, a)
    m = compute_metrics(fi, res)
    sc = score_candidate(m, res)
    return {
        "name": c.get("name") or fi.company, "company": fi.company, "market": fi.market,
        "id": c.get("corp_code") or c.get("ticker") or c.get("input"),
        "currency": fi.currency, "price": res["price"], "price_date": res["price_date"],
        "fair_value": res["fair_value"], "upside_pct": res["upside_pct"],
        "target_range": res["target_range"], "verdict": res["verdict"],
        "confidence": res["confidence"], "confidence_reasons": res["confidence_reasons"],
        "metrics": m, "score": sc,
        "defaults_used": res["defaults_used"], "warnings": res["warnings"],
        "estimated_last": res["estimated_last"], "years": res["years"],
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="산업 후보 일괄 스크리닝")
    ap.add_argument("candidates", help="candidates.json 경로")
    ap.add_argument("--fetch", action="store_true", help="기존 플러그인 스크립트로 캐시를 먼저 채움")
    ap.add_argument("--dart-key", help="DART API 키(KR --fetch)")
    ap.add_argument("--krx-key", help="KRX API 키(KR 주가 --fetch)")
    ap.add_argument("--contact", help="SEC User-Agent 이메일(US --fetch)")
    ap.add_argument("--assumptions-kr", help="KR 공통 가정 JSON")
    ap.add_argument("--assumptions-us", help="US 공통 가정 JSON")
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--out", help="결과 JSON 경로(기본: candidates.json 옆 screening.json)")
    args = ap.parse_args()

    spec = json.loads(Path(args.candidates).read_text(encoding="utf-8"))
    cands = spec.get("candidates", [])
    if not cands:
        sys.exit("ERROR: candidates가 비어 있습니다")

    results, errors = [], []
    for c in cands:
        name = c.get("name", "?")
        log: list[str] = []
        try:
            if args.fetch:
                if c["market"].upper() == "KR":
                    fetch_kr(c, args.dart_key, args.krx_key, args.years, log)
                elif c["market"].upper() == "US":
                    fetch_us(c, args.contact, log)
            r = evaluate_candidate(c, args)
            r["fetch_log"] = log
            results.append(r)
            print(f"[OK] {name}: 점수 {r['score']['final']} / 판정 {r['verdict']} / 신뢰도 {r['confidence']}")
        except SystemExit as e:
            errors.append({"name": name, "error": str(e), "fetch_log": log})
            print(f"[FAIL] {name}: {e}")
        except Exception as e:  # 한 후보 실패가 전체를 죽이지 않게
            errors.append({"name": name, "error": f"{type(e).__name__}: {e}", "fetch_log": log})
            print(f"[FAIL] {name}: {type(e).__name__}: {e}")

    ranked = sorted(results, key=lambda r: (r["score"]["final"] is None, -(r["score"]["final"] or 0)))
    for i, r in enumerate(ranked, 1):
        r["rank"] = i

    out = {
        "industry": spec.get("industry", ""),
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "scoring_note": ("항목별 (점수/만점)로 계산하고 데이터 결측 항목은 분모에서 제외해 정규화(0~100). "
                         "신뢰도 감점 후 최종점수. 후보 간 동일 가정(--assumptions-*)을 적용했으므로 "
                         "상대 비교용이며, 절대 적정주가는 종목별 정밀 가정이 아니라는 한계가 있다."),
        "disclaimer": ("이 순위는 공시 재무데이터와 명시된 가정에 종속된 조건부 추정이며 투자 권유가 아니다. "
                       "기본값을 쓴 가정이 있는 후보는 그 부분에 한해 '이 정보들은 정확하지 않습니다'."),
        "candidates": ranked,
        "errors": errors,
    }
    out_path = args.out or str(Path(args.candidates).with_name("screening.json"))
    Path(out_path).write_text(json.dumps(out, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps({"saved": out_path, "ranked": [(r["rank"], r["name"], r["score"]["final"]) for r in ranked],
                      "errors": len(errors)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
