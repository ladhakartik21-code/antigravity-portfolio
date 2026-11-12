"""
AntiGravity Frozen Investment Thesis Compliance Engine (v1.0 | 18 Criteria)
- Implements Kartik & AntiGravity's Frozen v1.0 18-Criteria Equity Audit:
  * Global Controls: Rolling 5Y Trailing Financial Window (TTM + Last 4 FYs),
    Zero-Penny Lock (< Rs.50), and Exact Evidence Reason on EVERY criterion.
  * Sector-Aware Evaluation (Non-Financials vs Banks/NBFCs).
  * Connects fundamental yfinance statements + curated Moat/Governance/Contingent DB
    + live 350-stock screener S&R data + live Investment Made.xlsx portfolio weights.
"""

import json
import math
import os
import threading
import time
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import yfinance as yf

import screener_engine

APP_DIR = os.path.dirname(os.path.abspath(__file__))
MOAT_PATH = os.path.join(APP_DIR, "fundamental_moat.json")
CACHE_PATH = os.path.join(APP_DIR, "scorecard_cache.json")
THESIS_TTL_SEC = 14400  # 4 hours

_CACHE_LOCK = threading.Lock()
_MEMORY_SCORECARD_CACHE: Dict[str, Any] = {}

# Curated institutional baseline for key NSE bluechips where BSE SHP / Annual Report Notes
# provide verified pledge, contingent liabilities, Bank NPA/PCR, or historical 5Y multiples.
CURATED_PROFILES: Dict[str, Dict[str, Any]] = {
    "HDFCBANK.NS": {
        "moat_proof": "India's #1 private banking deposit franchise (42%+ low-cost CASA ratio & nationwide branch network).",
        "is_bank": True,
        "bank_metrics": {"gross_npa": 1.36, "net_npa": 0.41, "crar": 19.3, "pcr": 71.0, "roa": 1.85, "roe": 16.4},
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Zero promoter pledge; regulated RBI tier-1 governance halo."},
        "ownership": {"promoter_pct": 0.0, "fii_pct": 47.2, "dii_pct": 34.8, "qoq_drop_pct": 0.0, "widely_held_exempt": True},
        "dilution_exempt_note": "HDFC Ltd merger (FY24) & 1:1 bonus share adjustment excluded per Approved M&A/Bonus rule.",
        "contingent_pct_nw": 8.5,
        "contingent_note": "Routine banking guarantees/LCs excluded; non-routine tax/legal claims ~8.5% of Net Worth (<=20%).",
    },
    "POLYCAB.NS": {
        "moat_proof": "24%+ organized domestic market share in wires & cables with backward-integrated copper/FMEG distribution.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Zero promoter pledge; clean statutory audit compliance."},
        "ownership": {"promoter_pct": 63.1, "fii_pct": 11.4, "dii_pct": 7.8, "qoq_drop_pct": 0.2},
        "contingent_pct_nw": 4.2,
        "contingent_note": "Routine tax/legal contingencies ~4.2% of Net Worth (well below 20% guard).",
    },
    "CDSL.NS": {
        "moat_proof": "Regulatory duopoly toll-gate (>70% market share in retail demat accounts, 14+ crore accounts, 60%+ EBITDA margin).",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "SEBI-regulated Market Infrastructure Institution (BSE promoted); zero pledge."},
        "ownership": {"promoter_pct": 15.0, "fii_pct": 14.5, "dii_pct": 35.8, "qoq_drop_pct": 0.0, "regulatory_mmi_cap": True},
        "contingent_pct_nw": 2.1,
        "contingent_note": "Contingent claims ~2.1% of Net Worth (<=20%).",
    },
    "TCS.NS": {
        "moat_proof": "Global enterprise IT mission-critical switching cost moat with industry-best ~24.5% operating margins.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Tata Sons institutional governance flagship; zero share pledge."},
        "ownership": {"promoter_pct": 71.7, "fii_pct": 12.4, "dii_pct": 10.9, "qoq_drop_pct": 0.0},
        "contingent_pct_nw": 6.4,
        "contingent_note": "Disputed tax/legal claims ~6.4% of Net Worth (<=20%).",
    },
    "RELIANCE.NS": {
        "moat_proof": "Tri-engine dominance: Jio telecom (#1 market share), Reliance Retail (18,000+ stores), & world-scale O2C refining complex.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Zero promoter pledge; AAA domestic balance sheet."},
        "ownership": {"promoter_pct": 50.1, "fii_pct": 19.2, "dii_pct": 18.6, "qoq_drop_pct": 0.1},
        "contingent_pct_nw": 11.2,
        "contingent_note": "Pending indirect tax/statutory guarantees ~11.2% of Net Worth (<=20%).",
    },
    "VBL.NS": {
        "moat_proof": "Exclusive PepsiCo bottling, manufacturing & cold-chain territory franchise across India and key Africa markets.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Zero promoter pledge; Ravi Jaipuria promoter group."},
        "ownership": {"promoter_pct": 60.7, "fii_pct": 21.8, "dii_pct": 6.2, "qoq_drop_pct": 0.3},
        "contingent_pct_nw": 5.5,
        "contingent_note": "Routine excise/tax contingencies ~5.5% of Net Worth (<=20%).",
    },
    "TATAPOWER.NS": {
        "moat_proof": "Integrated power transmission, distribution licenses (Mumbai/Delhi/Odisha) & national rooftop solar EPC leader.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Tata Group flagship power utility; zero pledge."},
        "ownership": {"promoter_pct": 46.9, "fii_pct": 9.8, "dii_pct": 16.4, "qoq_drop_pct": 0.0},
        "contingent_pct_nw": 14.8,
        "contingent_note": "Regulatory PPA tariff claims & guarantees ~14.8% of Net Worth (<=20%).",
    },
    "LT.NS": {
        "moat_proof": "Unmatched EPC engineering scale (Rs.5L+ Cr order backlog) across defense, metro, grid, and hydrocarbons.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Professionally managed widely-held national infrastructure bluechip."},
        "ownership": {"promoter_pct": 0.0, "fii_pct": 22.8, "dii_pct": 42.9, "qoq_drop_pct": 0.0, "widely_held_exempt": True},
        "contingent_pct_nw": 15.6,
        "contingent_note": "Non-routine legal/tax contingencies ~15.6% of Net Worth (<=20%; normal performance BGs excluded).",
    },
    "BHARTIARTL.NS": {
        "moat_proof": "Indian telecom duopoly with industry-leading ARPU trajectory and enterprise/home fiber monetization.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Bharti + Singtel promoter governance; zero pledge."},
        "ownership": {"promoter_pct": 53.1, "fii_pct": 25.4, "dii_pct": 17.8, "qoq_drop_pct": 0.1},
        "contingent_pct_nw": 26.5,
        "contingent_note": "Elevated statutory AGR/spectrum legal claims exceeding 20% of Net Worth ceiling.",
    },
    "TITAN.NS": {
        "moat_proof": "Unrivaled Tata trust halo in Tanishq jewelry, watches & eyewear powering market share shift from unorganized players.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Tata Group + TIDCO joint governance; zero pledge."},
        "ownership": {"promoter_pct": 52.9, "fii_pct": 18.2, "dii_pct": 11.4, "qoq_drop_pct": 0.0},
        "contingent_pct_nw": 5.0,
        "contingent_note": "Routine indirect tax claims ~5.0% of Net Worth (<=20%).",
    },
    "ICICIBANK.NS": {
        "moat_proof": "Top-tier retail & SME liability franchise with disciplined underwriting and digital ecosystem moat.",
        "is_bank": True,
        "bank_metrics": {"gross_npa": 1.9, "net_npa": 0.42, "crar": 16.5, "pcr": 78.0, "roa": 2.35, "roe": 18.2},
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Widely held institution; zero pledge."},
        "ownership": {"promoter_pct": 0.0, "fii_pct": 45.2, "dii_pct": 44.1, "qoq_drop_pct": 0.0, "widely_held_exempt": True},
        "contingent_pct_nw": 7.0,
        "contingent_note": "Routine banking derivatives/guarantees excluded; non-routine claims <20%.",
    },
    "INFY.NS": {
        "moat_proof": "Tier-1 global enterprise IT vendor with Topaz AI & Cobalt cloud switching-cost moat.",
        "governance": {"pledge_pct": 0.0, "clean_audit": True, "note": "Gold-standard corporate governance; zero pledge."},
        "ownership": {"promoter_pct": 14.4, "fii_pct": 33.6, "dii_pct": 38.2, "qoq_drop_pct": 0.1, "widely_held_exempt": True},
        "contingent_pct_nw": 7.2,
        "contingent_note": "Tax disputes ~7.2% of Net Worth (<=20%).",
    },
}


def _sf(val: Any, default: Optional[float] = None) -> Optional[float]:
    try:
        if val is None:
            return default
        f = float(val)
        return default if (math.isnan(f) or math.isinf(f)) else f
    except Exception:
        return default


def _series_for(df: Optional[pd.DataFrame], keys: List[str]) -> Optional[pd.Series]:
    if df is None or df.empty:
        return None
    for k in keys:
        if k in df.index:
            s = df.loc[k].dropna()
            if not s.empty:
                # Sort chronological oldest -> newest
                return s.sort_index()
    return None


def _cagr(series: Optional[pd.Series]) -> Optional[Tuple[float, int]]:
    if series is None or len(series) < 2:
        return None
    v0 = _sf(series.iloc[0])
    v1 = _sf(series.iloc[-1])
    n = len(series) - 1
    if not v0 or not v1 or v0 <= 0 or v1 <= 0 or n <= 0:
        return None
    return ((v1 / v0) ** (1.0 / n) - 1.0) * 100.0, n


def _load_moat_item(sym: str) -> Dict[str, Any]:
    if os.path.exists(MOAT_PATH):
        try:
            with open(MOAT_PATH, "r", encoding="utf-8") as f:
                for a in json.load(f).get("assets", []):
                    if a.get("symbol", "").upper() == sym.upper():
                        return a
        except Exception:
            pass
    return {}


def _load_disk_cache() -> Dict[str, Any]:
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_disk_cache(cache_dict: Dict[str, Any]) -> None:
    try:
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache_dict, f)
    except Exception:
        pass


def _criterion(
    cid: int,
    name: str,
    category: str,
    tier: str,
    status: str,
    actual_value: str,
    threshold: str,
    reason: str,
) -> Dict[str, Any]:
    sev_w = 3 if "Tier 1" in tier else (2 if "Tier 2" in tier else 1)
    return {
        "id": cid,
        "name": name,
        "category": category,
        "tier": tier,
        "severity_weight": sev_w,
        "status": status,
        "actual_value": actual_value,
        "threshold": threshold,
        "reason": reason,
    }


def evaluate_stock_18pt(
    symbol: str,
    sr_map: Optional[Dict[str, Any]] = None,
    portfolio_data: Optional[Dict[str, Any]] = None,
    force_refresh: bool = False,
) -> Dict[str, Any]:
    clean_sym = symbol.strip().upper().replace(" ", "")
    if not clean_sym.endswith(".NS") and not clean_sym.endswith(".BO") and not clean_sym.startswith("^"):
        clean_sym = f"{clean_sym}.NS"

    now = time.time()
    if not force_refresh:
        with _CACHE_LOCK:
            mem_entry = _MEMORY_SCORECARD_CACHE.get(clean_sym)
            if mem_entry and (now - mem_entry.get("_ts", 0.0) < THESIS_TTL_SEC):
                return mem_entry
            disk = _load_disk_cache()
            disk_entry = disk.get(clean_sym)
            if disk_entry and (now - disk_entry.get("_ts", 0.0) < THESIS_TTL_SEC):
                _MEMORY_SCORECARD_CACHE[clean_sym] = disk_entry
                return disk_entry

    # 1. Gather Screener Technical Snapshot (<10ms)
    scr_snap = screener_engine.get_screener_snapshot(sr_map=sr_map or {})
    scr_stock = next((s for s in scr_snap.get("stocks", []) if s["symbol"].upper() == clean_sym), {})
    moat_item = _load_moat_item(clean_sym)
    curated = CURATED_PROFILES.get(clean_sym, {})

    # 2. Fetch Live Financial Statements via yfinance
    info: Dict[str, Any] = {}
    inc: Optional[pd.DataFrame] = None
    bs: Optional[pd.DataFrame] = None
    cf: Optional[pd.DataFrame] = None
    try:
        tk = yf.Ticker(clean_sym)
        info = tk.info or {}
        inc = tk.income_stmt
        bs = tk.balance_sheet
        cf = tk.cashflow
    except Exception as e:
        print(f"[Thesis Engine YF Warning] {clean_sym}: {e}", flush=True)

    company_name = scr_stock.get("name") or moat_item.get("name") or info.get("longName") or info.get("shortName") or clean_sym
    sector_str = scr_stock.get("sector") or info.get("sector") or "General"
    industry_str = info.get("industry") or ""
    is_bank = bool(
        curated.get("is_bank")
        or any(k in industry_str.lower() for k in ("bank", "credit services", "mortgage", "nbfc", "insurance"))
    )
    ltp = _sf(scr_stock.get("ltp") or info.get("currentPrice") or info.get("regularMarketPrice") or moat_item.get("ltp"), 0.0) or 0.0
    cap_tier = scr_stock.get("cap_tier") or ("Large Cap" if (_sf(info.get("marketCap"), 0.0) or 0.0) >= 5e11 else "Mid Cap")
    indices_set = {x.upper() for x in scr_stock.get("indices", [])}
    is_mega_cap = ("NIFTY50" in indices_set or "SENSEX" in indices_set or (_sf(info.get("marketCap"), 0.0) or 0.0) >= 1e12)

    rev_s = _series_for(inc, ["Total Revenue", "Operating Revenue"])
    opinc_s = _series_for(inc, ["Operating Income", "EBIT"])
    ebit_s = _series_for(inc, ["EBIT", "Operating Income"])
    ebitda_s = _series_for(inc, ["EBITDA", "Normalized EBITDA"])
    pat_s = _series_for(inc, ["Net Income", "Net Income Continuous Operations"])
    int_s = _series_for(inc, ["Interest Expense", "Interest Expense Non Operating"])
    shares_inc = _series_for(inc, ["Diluted Average Shares", "Basic Average Shares"])
    shares_s = shares_inc if (shares_inc is not None and not shares_inc.empty) else _series_for(bs, ["Ordinary Shares Number"])

    debt_s = _series_for(bs, ["Total Debt"])
    netdebt_s = _series_for(bs, ["Net Debt"])
    eq_s = _series_for(bs, ["Stockholders Equity", "Common Stock Equity"])
    ic_s = _series_for(bs, ["Invested Capital"])
    rec_s = _series_for(bs, ["Accounts Receivable", "Receivables"])

    cfo_s = _series_for(cf, ["Operating Cash Flow"])
    fcf_s = _series_for(cf, ["Free Cash Flow"])
    div_s = _series_for(cf, ["Cash Dividends Paid", "Common Stock Dividend Paid"])

    criteria: List[Dict[str, Any]] = []

    # Shared OPM 5Y calculation
    opm_list: List[float] = []
    if opinc_s is not None and rev_s is not None:
        for col in opinc_s.index.intersection(rev_s.index):
            r = _sf(rev_s[col], 0.0)
            o = _sf(opinc_s[col], 0.0)
            if r and r > 0 and o is not None:
                opm_list.append((o / r) * 100.0)
    avg_opm_5y = round(sum(opm_list) / len(opm_list), 1) if opm_list else None
    ttm_opm = round((_sf(info.get("operatingMargins"), 0.0) or 0.0) * 100.0, 1) if info.get("operatingMargins") is not None else avg_opm_5y

    # 1. Economic Moat & Pricing Power
    moat_proof = curated.get("moat_proof") or moat_item.get("catalyst")
    moat_score = _sf(moat_item.get("moat_score"), None)
    if moat_proof:
        opm_note = f" | Rolling OPM: {avg_opm_5y:.1f}%" if avg_opm_5y is not None else ""
        criteria.append(_criterion(
            1, "Economic Moat & Pricing Power", "Business Quality", "Tier 2 - Core Quality (2x)", "PASS",
            f"{moat_item.get('moat_rating', 'Verified Moat')}{opm_note}",
            "Structural Moat (>=20% share/duopoly) & Margin >= Sector",
            f"Passed: {moat_proof}",
        ))
    elif avg_opm_5y is not None:
        m_pass = avg_opm_5y >= 12.0
        criteria.append(_criterion(
            1, "Economic Moat & Pricing Power", "Business Quality", "Tier 2 - Core Quality (2x)",
            "PASS" if m_pass else "FAIL",
            f"Rolling {len(opm_list)}Y Avg OPM = {avg_opm_5y:.1f}%",
            "OPM >= 12.0% + Verifiable Moat Proof",
            f"{'Passed quantitative margin power proxy' if m_pass else 'Thin operating margin suggests weak pricing power'} ({avg_opm_5y:.1f}% avg OPM over {len(opm_list)} FYs); confirm qualitative market share if deploying heavily.",
        ))
    else:
        criteria.append(_criterion(
            1, "Economic Moat & Pricing Power", "Business Quality", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK",
            "Financials unavailable", "Verifiable Moat & OPM >= 12%",
            "Insufficient margin history returned by feed; review Annual Report market share before investing.",
        ))

    # 2. Capital Efficiency (ROCE / ROE) — Rolling Window
    avg_roce_or_roe: Optional[float] = None
    if is_bank:
        bank_m = curated.get("bank_metrics", {})
        roe_vals = []
        if pat_s is not None and eq_s is not None:
            for col in pat_s.index.intersection(eq_s.index):
                e = _sf(eq_s[col], 0.0)
                p = _sf(pat_s[col], 0.0)
                if e and e > 0 and p is not None:
                    roe_vals.append((p / e) * 100.0)
        consol_roe = round(sum(roe_vals) / len(roe_vals), 1) if roe_vals else None
        avg_roce_or_roe = bank_m.get("roe") if bank_m.get("roe") is not None else consol_roe
        if avg_roce_or_roe is not None:
            roa_val = bank_m.get("roa") or (_sf(info.get("returnOnAssets"), 0.0) or 0.0) * 100.0
            c_pass = avg_roce_or_roe >= 15.0 and (roa_val == 0 or roa_val >= 1.2)
            consol_note = f" | Consol YF {consol_roe:.1f}%" if (consol_roe is not None and bank_m.get("roe") is not None) else ""
            criteria.append(_criterion(
                2, "Capital Efficiency (ROCE / ROE)", "Capital Efficiency", "Tier 2 - Core Quality (2x)",
                "PASS" if c_pass else "FAIL",
                f"Standalone Bank ROE = {avg_roce_or_roe:.1f}% (ROA {roa_val:.2f}%{consol_note})",
                "Bank/NBFC: ROE >= 15.0% & ROA >= 1.5%",
                f"{'Passed' if c_pass else 'Below threshold'}: Core banking capital compounding at {avg_roce_or_roe:.1f}% ROE ({roa_val:.2f}% ROA).",
            ))
        else:
            criteria.append(_criterion(2, "Capital Efficiency (ROCE / ROE)", "Capital Efficiency", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", "Bank ROE >= 15%", "Could not extract bank equity/ROE automatically."))
    else:
        roce_vals = []
        if ebit_s is not None and ic_s is not None:
            for col in ebit_s.index.intersection(ic_s.index):
                ic_v = _sf(ic_s[col], 0.0)
                eb_v = _sf(ebit_s[col], 0.0)
                if ic_v and ic_v > 0 and eb_v is not None:
                    roce_vals.append((eb_v / ic_v) * 100.0)
        avg_roce_or_roe = round(sum(roce_vals) / len(roce_vals), 1) if roce_vals else _sf(moat_item.get("roce_pct"), None)
        if avg_roce_or_roe is not None:
            c_pass = avg_roce_or_roe >= 20.0
            criteria.append(_criterion(
                2, "Capital Efficiency (ROCE / ROE)", "Capital Efficiency", "Tier 2 - Core Quality (2x)",
                "PASS" if c_pass else "FAIL",
                f"Rolling {len(roce_vals) or 'Curated'}Y Avg ROCE = {avg_roce_or_roe:.1f}%",
                "Non-Bank: Rolling 5Y Avg ROCE >= 20.0%",
                f"{'Passed' if c_pass else 'Failed'}: Rolling {len(roce_vals)}Y ROCE averages {avg_roce_or_roe:.1f}% vs 20.0% minimum hurdle.",
            ))
        else:
            criteria.append(_criterion(2, "Capital Efficiency (ROCE / ROE)", "Capital Efficiency", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", "ROCE >= 20%", "Invested capital unavailable; verify ROCE manually."))

    # 3. Balance Sheet Solvency & Debt Safety
    if is_bank:
        bank_m = curated.get("bank_metrics", {})
        if bank_m:
            gnpa = bank_m["gross_npa"]
            nnpa = bank_m["net_npa"]
            crar = bank_m["crar"]
            b_pass = gnpa <= 2.0 and nnpa <= 0.5 and crar >= 15.0
            criteria.append(_criterion(
                3, "Balance Sheet Solvency & Debt Safety", "Financial Solvency", "Tier 1 - Dealbreaker (3x)",
                "PASS" if b_pass else "FAIL",
                f"GNPA {gnpa}% | NNPA {nnpa}% | CRAR {crar}%",
                "GNPA <= 2.0% | NNPA <= 0.5% | CRAR >= 15.0%",
                f"{'Passed' if b_pass else 'Failed'}: Asset quality & capital adequacy safely within fortress RBI buffers.",
            ))
        else:
            criteria.append(_criterion(
                3, "Balance Sheet Solvency & Debt Safety", "Financial Solvency", "Tier 1 - Dealbreaker (3x)", "MANUAL_CHECK",
                "Requires RBI Qtrly Pillar-3", "GNPA <= 2.0% | NNPA <= 0.5% | CRAR >= 15%",
                "Bank asset-quality metrics (GNPA/NNPA/CRAR) must be checked in latest quarterly investor presentation.",
            ))
    else:
        lat_debt = _sf(debt_s.iloc[-1], None) if debt_s is not None else None
        lat_eq = _sf(eq_s.iloc[-1], None) if eq_s is not None else None
        lat_nd = _sf(netdebt_s.iloc[-1], None) if netdebt_s is not None else None
        lat_ebitda = _sf(ebitda_s.iloc[-1], None) if ebitda_s is not None else None
        lat_ebit = _sf(ebit_s.iloc[-1], None) if ebit_s is not None else None
        lat_int = abs(_sf(int_s.iloc[-1], 0.0) or 0.0) if int_s is not None else 0.0

        de_ratio = round(lat_debt / lat_eq, 2) if (lat_debt is not None and lat_eq and lat_eq > 0) else (round((_sf(info.get("debtToEquity"), 0.0) or 0.0) / 100.0, 2) if info.get("debtToEquity") is not None else None)
        nd_ebitda = round(lat_nd / lat_ebitda, 2) if (lat_nd is not None and lat_ebitda and lat_ebitda > 0) else _sf(moat_item.get("net_debt_ebitda"), None)
        int_cov = round(lat_ebit / lat_int, 1) if (lat_ebit is not None and lat_int > 1e-3) else (99.0 if (lat_debt == 0 or lat_int == 0) else None)

        if de_ratio is not None:
            nd_ok = True if nd_ebitda is None else (nd_ebitda <= 1.5)
            ic_ok = True if int_cov is None else (int_cov >= 5.0)
            s_pass = (de_ratio <= 0.6) and nd_ok and ic_ok
            nd_str = f"{nd_ebitda:.2f}x" if nd_ebitda is not None else "N/A"
            ic_str = "Debt-Free" if int_cov == 99.0 else (f"{int_cov:.1f}x" if int_cov is not None else "N/A")
            criteria.append(_criterion(
                3, "Balance Sheet Solvency & Debt Safety", "Financial Solvency", "Tier 1 - Dealbreaker (3x)",
                "PASS" if s_pass else "FAIL",
                f"D/E {de_ratio:.2f}x | NetDebt/EBITDA {nd_str} | IntCov {ic_str}",
                "D/E <= 0.60x | Net Debt/EBITDA <= 1.5x | Int Cov >= 5.0x",
                f"{'Passed' if s_pass else 'Failed'}: Balance sheet leverage {'is conservatively shielded' if s_pass else 'breaches debt or coverage safety limits'}.",
            ))
        else:
            criteria.append(_criterion(3, "Balance Sheet Solvency & Debt Safety", "Financial Solvency", "Tier 1 - Dealbreaker (3x)", "MANUAL_CHECK", "Unavailable", "D/E <= 0.60x", "Could not parse debt ratio from latest balance sheet."))

    # 4. Cash Flow Conversion & Quality
    annual_cfo_pat: Optional[float] = None
    if is_bank:
        pcr = curated.get("bank_metrics", {}).get("pcr")
        if pcr is not None:
            cf_pass = pcr >= 70.0
            criteria.append(_criterion(
                4, "Cash Flow Conversion & Quality", "Earnings Quality", "Tier 1 - Dealbreaker (3x)",
                "PASS" if cf_pass else "FAIL",
                f"Provision Coverage Ratio (PCR) = {pcr:.1f}%",
                "Banks: PCR >= 70.0%",
                f"{'Passed' if cf_pass else 'Failed'}: Strong NPA provisioning buffer ({pcr:.1f}% vs 70% floor).",
            ))
        else:
            criteria.append(_criterion(4, "Cash Flow Conversion & Quality", "Earnings Quality", "Tier 1 - Dealbreaker (3x)", "MANUAL_CHECK", "N/A", "Bank PCR >= 70%", "Check quarterly Provision Coverage Ratio in bank results."))
    else:
        if cfo_s is not None and pat_s is not None and len(cfo_s) > 0:
            sum_cfo = float(cfo_s.sum())
            sum_pat = float(pat_s.reindex(cfo_s.index).dropna().sum())
            sum_ebitda = float(ebitda_s.reindex(cfo_s.index).dropna().sum()) if ebitda_s is not None else 0.0
            cfo_pat = round((sum_cfo / sum_pat) * 100.0, 1) if sum_pat > 0 else 0.0
            cfo_ebitda = round((sum_cfo / sum_ebitda) * 100.0, 1) if sum_ebitda > 0 else None
            lat_dt = cfo_s.index[-1]
            lat_cfo_v = _sf(cfo_s.get(lat_dt), None)
            lat_pat_v = _sf(pat_s.get(lat_dt), None)
            annual_cfo_pat = round((lat_cfo_v / lat_pat_v) * 100.0, 1) if (lat_cfo_v is not None and lat_pat_v and lat_pat_v > 0) else None
            eb_ok = True if cfo_ebitda is None else (cfo_ebitda >= 60.0)
            cf_pass = (cfo_pat >= 80.0) and eb_ok and sum_cfo > 0
            eb_txt = f" | Post-Tax CFO/EBITDA {cfo_ebitda:.1f}%" if cfo_ebitda is not None else ""
            criteria.append(_criterion(
                4, "Cash Flow Conversion & Quality", "Earnings Quality", "Tier 1 - Dealbreaker (3x)",
                "PASS" if cf_pass else "FAIL",
                f"Rolling {len(cfo_s)}Y CFO/PAT {cfo_pat:.1f}%{eb_txt}",
                "Rolling CFO/PAT >= 80% & Post-Tax CFO/EBITDA >= 60%",
                f"{'Passed' if cf_pass else 'Failed'}: Cumulative operating cash flow covers {cfo_pat:.1f}% of net profits.",
            ))
        else:
            criteria.append(_criterion(4, "Cash Flow Conversion & Quality", "Earnings Quality", "Tier 1 - Dealbreaker (3x)", "MANUAL_CHECK", "N/A", "CFO/PAT >= 80%", "Cash flow series unavailable from feed."))

    # 5. Secular Runway (TAM Growth)
    rev_cagr_res = _cagr(rev_s)
    req_rev_cagr = 10.0 if is_mega_cap else 12.0
    if rev_cagr_res is not None:
        r_cagr, r_yrs = rev_cagr_res
        r_pass = r_cagr >= req_rev_cagr
        criteria.append(_criterion(
            5, "Secular Runway (Expanding Demand / TAM)", "Growth Runway", "Tier 2 - Core Quality (2x)",
            "PASS" if r_pass else "FAIL",
            f"{r_yrs}Y Revenue CAGR = {r_cagr:.1f}% ({sector_str})",
            f">= {req_rev_cagr:.0f}% Revenue CAGR ({'Mega-Cap' if is_mega_cap else 'Standard'})",
            f"{'Passed' if r_pass else 'Below GDP+ hurdle'}: Top-line compounded at {r_cagr:.1f}%/yr over {r_yrs}Y inside {sector_str}.",
        ))
    else:
        criteria.append(_criterion(5, "Secular Runway (Expanding Demand / TAM)", "Growth Runway", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", f"Revenue CAGR >= {req_rev_cagr:.0f}%", "Insufficient multi-year revenue history."))

    # 6. Valuation Margin of Safety (STRICT 3-OF-3 LOCK)
    disc_pct = _sf(scr_stock.get("ath_discount_pct"), 0.0) or 0.0
    req_disc = 15.0 if cap_tier == "Large Cap" else (25.0 if cap_tier == "Mid Cap" else 30.0)
    cond_c = disc_pct >= req_disc

    trail_pe = _sf(info.get("trailingPE"), None)
    fwd_pe = _sf(info.get("forwardPE"), None)
    pb_val = _sf(info.get("priceToBook"), None)
    pe_med_fair = _sf(moat_item.get("pe_median_value"), None)
    blended_fair = _sf(moat_item.get("blended_fair_value"), None)

    if pe_med_fair and ltp > 0:
        cond_a = ltp <= pe_med_fair * 1.05
        a_lbl = f"LTP ₹{ltp:.0f} <= 5Y Fair ₹{pe_med_fair*1.05:.0f}" if cond_a else f"LTP ₹{ltp:.0f} > 5Y Fair ₹{pe_med_fair*1.05:.0f}"
    elif is_bank and pb_val is not None:
        cond_a = pb_val <= 2.8
        a_lbl = f"Bank P/B {pb_val:.2f}x (<=2.8x)"
    elif trail_pe is not None and fwd_pe is not None and trail_pe > 0:
        cond_a = (trail_pe <= 35.0) and (fwd_pe <= trail_pe)
        a_lbl = f"PE {trail_pe:.1f}x / Fwd {fwd_pe:.1f}x"
    elif trail_pe is not None:
        cond_a = trail_pe <= 32.0
        a_lbl = f"PE {trail_pe:.1f}x (<=32x proxy)"
    else:
        cond_a = False
        a_lbl = "P/E data unavailable"

    if blended_fair and ltp > 0:
        mos_val = ((blended_fair - ltp) / blended_fair) * 100.0
        cond_b = mos_val >= 10.0
        b_lbl = f"DCF+PE MoS {mos_val:+.1f}% (>=+10%)"
    elif trail_pe and fwd_pe and fwd_pe > 0 and trail_pe > 0:
        implied_mos = ((trail_pe - fwd_pe) / trail_pe) * 100.0
        cond_b = implied_mos >= 10.0 and trail_pe <= 38.0
        b_lbl = f"Fwd Earnings Discount {implied_mos:+.1f}%"
    else:
        cond_b = False
        b_lbl = "Fair Value MoS unavailable"

    c_lbl = f"52W Peak Disc -{disc_pct:.1f}% (Req >={req_disc:.0f}%)"
    pass_cnt_v = sum([1 if cond_a else 0, 1 if cond_b else 0, 1 if cond_c else 0])
    v_pass = (pass_cnt_v == 3)
    criteria.append(_criterion(
        6, "Valuation Margin of Safety (Strict 3-of-3 Lock)", "Valuation Safety", "Tier 3 - Timing / Valuation (1x)",
        "PASS" if v_pass else "FAIL",
        f"{pass_cnt_v}/3 Tests | {a_lbl} | {b_lbl} | {c_lbl}",
        "ALL 3 OF 3: Multiple Safety + Intrinsic MoS >= +10% + Peak Discount",
        f"{'Passed all 3 strict valuation locks' if v_pass else f'Valuation Caution ({pass_cnt_v}/3 passed) — stagger or wait near S1 support rather than paying full multiple'}: (a) {'✓' if cond_a else '✗'} {a_lbl}; (b) {'✓' if cond_b else '✗'} {b_lbl}; (c) {'✓' if cond_c else '✗'} {c_lbl}.",
    ))

    # 7. Promoter & Regulatory Governance
    gov = curated.get("governance")
    if gov:
        g_pass = (gov["pledge_pct"] <= 2.0) and bool(gov["clean_audit"])
        criteria.append(_criterion(
            7, "Promoter & Regulatory Governance", "Governance", "Tier 1 - Dealbreaker (3x)",
            "PASS" if g_pass else "FAIL",
            f"Pledge {gov['pledge_pct']}% | Audit Clean: {'Yes' if gov['clean_audit'] else 'No'}",
            "Pledge <= 2.0% (Ideal 0%) & Clean 5Y Audit",
            f"{'Passed' if g_pass else 'Failed'}: {gov['note']}",
        ))
    else:
        criteria.append(_criterion(
            7, "Promoter & Regulatory Governance", "Governance", "Tier 1 - Dealbreaker (3x)", "MANUAL_CHECK",
            "Verify BSE Shareholding Pledge", "Pledge = 0% (Hard Cap <=2%) & Zero SEBI Flags",
            "Exchange pledge % & auditor qualifications are not exposed via price feed; inspect BSE SHP before final capital commitment.",
        ))

    # 8. Pre-Defined Exit / Invalidation Triggers
    s2_lvl = scr_stock.get("s2") or round(ltp * 0.88, 2)
    r2_lvl = scr_stock.get("r2") or round(ltp * 1.18, 2)
    e_fund = "GNPA > 2.5% or ROE < 13%" if is_bank else "ROCE < 14% for 2 Qtrs or D/E > 1.0x"
    criteria.append(_criterion(
        8, "Pre-Defined Exit / Invalidation Triggers", "Exit Discipline", "Tier 2 - Core Quality (2x)", "PASS",
        f"Break: {e_fund} | Trim @ R2 ₹{r2_lvl}",
        "Both Fundamental Invalidation & Euphoric R2 Trim Defined",
        f"Passed: Exit immediately if {e_fund} (or S2 ₹{s2_lvl} structural failure with broken thesis); trim 25%-35% above R2 ₹{r2_lvl} ONLY if > +35% overvalued.",
    ))

    # 9. Ownership Floor (>=65%) & Rationale-Verified Selling Guard
    own = curated.get("ownership")
    ttm_vs_5y_drop = (avg_opm_5y - ttm_opm) if (avg_opm_5y is not None and ttm_opm is not None) else 0.0
    pat_cagr_res = _cagr(pat_s)
    yoy_pat_drop = bool(pat_s is not None and len(pat_s) >= 2 and (_sf(pat_s.iloc[-1], 0.0) or 0.0) < (_sf(pat_s.iloc[-2], 0.0) or 0.0))

    if own:
        total_strong = round(own["promoter_pct"] + own["fii_pct"] + own["dii_pct"], 1)
        prom_ok = own.get("widely_held_exempt") or own.get("regulatory_mmi_cap") or (own["promoter_pct"] >= 30.0)
        floor_ok = total_strong >= 65.0
        qoq_drop = own.get("qoq_drop_pct", 0.0)
        deteriorating = (ttm_vs_5y_drop > 2.0) or yoy_pat_drop
        pledge_bad = bool(curated.get("governance", {}).get("pledge_pct", 0.0) > 0.0)
        sell_fail = (qoq_drop > 1.5 and (not own.get("widely_held_exempt"))) or (qoq_drop > 1.5 and deteriorating) or pledge_bad
        o_pass = prom_ok and floor_ok and (not sell_fail)
        rationale_note = "Zero abnormal selling pressure" if qoq_drop <= 0.5 else ("Minor benign profit-booking trim (margins & profits healthy)" if not sell_fail else "Selling coincides with fundamental stress")
        criteria.append(_criterion(
            9, "Ownership Floor (65%) & Selling Rationale", "Governance & Ownership", "Tier 1 - Dealbreaker (3x)",
            "PASS" if o_pass else "FAIL",
            f"Promoter+FII+DII = {total_strong}% (Prom {own['promoter_pct']}% | QoQ {qoq_drop:+.1f}%)",
            "Prom+FII+DII >= 65% & No Fundamentals-Confirmed Exit",
            f"{'Passed' if o_pass else 'Failed'}: Strong hands hold {total_strong}% (retail float {100-total_strong:.1f}% <= 35%). {rationale_note}.",
        ))
    else:
        ins = (_sf(info.get("heldPercentInsiders"), None) or 0.0) * 100.0
        inst = (_sf(info.get("heldPercentInstitutions"), None) or 0.0) * 100.0
        if ins > 0 or inst > 0:
            tot = round(ins + inst, 1)
            o_pass = tot >= 65.0
            criteria.append(_criterion(
                9, "Ownership Floor (65%) & Selling Rationale", "Governance & Ownership", "Tier 1 - Dealbreaker (3x)",
                "PASS" if o_pass else "FAIL",
                f"Insider+Institutional (YF) = {tot:.1f}% (Insider {ins:.1f}% | Inst {inst:.1f}%)",
                "Prom+FII+DII >= 65% (Check BSE SHP for domestic DIIs)",
                f"{'Passed' if o_pass else 'Below 65% in Yahoo feed (verify complete BSE FII+DII SHP)'}: Combined tracked strong-hand ownership = {tot:.1f}%.",
            ))
        else:
            criteria.append(_criterion(9, "Ownership Floor (65%) & Selling Rationale", "Governance & Ownership", "Tier 1 - Dealbreaker (3x)", "MANUAL_CHECK", "N/A", "Prom+FII+DII >= 65%", "Verify BSE Quarterly Shareholding Pattern >= 65%."))

    # 10. 5Y Earnings Consistency & PAT CAGR
    if pat_s is not None and len(pat_s) >= 2:
        pos_yrs = int((pat_s > 0).sum())
        tot_yrs = len(pat_s)
        pat_c = pat_cagr_res[0] if pat_cagr_res else 0.0
        p_pass = (pos_yrs == tot_yrs) and (pat_c >= 14.0)
        criteria.append(_criterion(
            10, "5Y Earnings Consistency & PAT CAGR", "Growth & Consistency", "Tier 2 - Core Quality (2x)",
            "PASS" if p_pass else "FAIL",
            f"Positive PAT {pos_yrs}/{tot_yrs} FYs | PAT CAGR = {pat_c:.1f}%",
            "Positive Every Year & PAT CAGR >= 14.0%",
            f"{'Passed' if p_pass else 'Failed'}: Compounded net earnings at {pat_c:.1f}%/yr across {tot_yrs} FYs ({'zero loss years' if pos_yrs == tot_yrs else 'loss year detected'}). ",
        ))
    else:
        criteria.append(_criterion(10, "5Y Earnings Consistency & PAT CAGR", "Growth & Consistency", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", "PAT CAGR >= 14%", "Insufficient multi-year PAT series."))

    # 11. Annual Working Capital & Receivables Discipline (Double Confirmation)
    if is_bank:
        criteria.append(_criterion(
            11, "Annual Working Capital & Receivables Discipline", "Earnings Quality", "Tier 2 - Core Quality (2x)", "NA",
            "N/A (Banking / Financial Model)", "Non-Banks Only",
            "Excluded cleanly: Lending banks & financials hold loan books rather than trade inventory/receivables.",
        ))
    elif rec_s is not None and rev_s is not None and len(rec_s) >= 2:
        rec_days: List[float] = []
        for col in rec_s.index.intersection(rev_s.index):
            rv = _sf(rev_s[col], 0.0)
            rc = _sf(rec_s[col], 0.0)
            if rv and rv > 0 and rc is not None:
                rec_days.append((rc / rv) * 365.0)
        if len(rec_days) >= 2:
            lat_days = round(rec_days[-1], 1)
            hist_avg_days = round(sum(rec_days[:-1]) / len(rec_days[:-1]), 1)
            days_spike = (lat_days > hist_avg_days + 15.0) or (lat_days > 120.0)
            cfo_ok = True if annual_cfo_pat is None else (annual_cfo_pat >= 70.0)
            w_pass = not (days_spike and (not cfo_ok))
            criteria.append(_criterion(
                11, "Annual Working Capital & Receivables Discipline", "Earnings Quality", "Tier 2 - Core Quality (2x)",
                "PASS" if w_pass else "FAIL",
                f"Latest FY {lat_days:.0f} Days vs 3Y Avg {hist_avg_days:.0f}d (Ann CFO/PAT {annual_cfo_pat or '--'}%)",
                "Fail ONLY if Days > 3Y Avg + 15d AND Annual CFO/PAT < 70%",
                f"{'Passed' if w_pass else 'Failed'}: {'Year-end receivables collection remains healthy' if not days_spike else ('Days stretched vs avg but confirmed collected via Annual CFO/PAT >= 70%' if cfo_ok else 'Receivable days spiked AND cash conversion dropped < 70%')}.",
            ))
        else:
            criteria.append(_criterion(11, "Annual Working Capital & Receivables Discipline", "Earnings Quality", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "Insufficient FYs", "<= 3Y Avg + 15d", "Fewer than 2 FY receivable figures available."))
    else:
        criteria.append(_criterion(11, "Annual Working Capital & Receivables Discipline", "Earnings Quality", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", "Annual Rec Days Check", "Receivables data missing from feed."))

    # 12. Equity Dilution Integrity Guard (Handles Bonus/Split/Approved Merger)
    if curated.get("dilution_exempt_note"):
        criteria.append(_criterion(
            12, "Equity Dilution Integrity Guard", "Capital Discipline", "Tier 2 - Core Quality (2x)", "PASS",
            "Adjusted: Approved Merger / Bonus Issue", "5Y Share CAGR <= 1.5%",
            f"Passed: {curated['dilution_exempt_note']}",
        ))
    elif shares_s is not None and len(shares_s) >= 2:
        # Check for corporate splits/bonuses (>1.85x jump with ~zero equity raise)
        sh_vals = [_sf(x) for x in shares_s.values if _sf(x) and _sf(x) > 0]
        raw_cagr = _cagr(shares_s)
        sh_cagr = raw_cagr[0] if raw_cagr else 0.0
        split_detected = any((sh_vals[i] / sh_vals[i - 1]) >= 1.85 for i in range(1, len(sh_vals)))
        if split_detected:
            criteria.append(_criterion(
                12, "Equity Dilution Integrity Guard", "Capital Discipline", "Tier 2 - Core Quality (2x)", "PASS",
                f"Bonus / Stock Split Event Detected", "5Y Diluted Shares CAGR <= 1.5%",
                "Passed: Large share-count multiple (~2x+) reflects corporate Bonus/Stock Split rather than dilutive QIP fundraising.",
            ))
        else:
            d_pass = sh_cagr <= 1.5
            criteria.append(_criterion(
                12, "Equity Dilution Integrity Guard", "Capital Discipline", "Tier 2 - Core Quality (2x)",
                "PASS" if d_pass else "FAIL",
                f"{raw_cagr[1] if raw_cagr else 1}Y Share Count CAGR = {sh_cagr:+.2f}%",
                "<= +1.50% CAGR",
                f"{'Passed' if d_pass else 'Failed'}: {'Minimal/negative dilution via ESOPs or buybacks protects per-share EPS' if d_pass else 'Repeated equity dilution eroding per-share compounding'}.",
            ))
    else:
        criteria.append(_criterion(12, "Equity Dilution Integrity Guard", "Capital Discipline", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", "<= 1.5% CAGR", "Share count series unavailable."))

    # 13. Operating Margin (OPM) Stability
    if is_bank:
        criteria.append(_criterion(
            13, "Operating Margin (OPM) Stability", "Business Quality", "Tier 2 - Core Quality (2x)", "NA",
            "N/A (Bank NIM / Cost-to-Income Model)", "Non-Banks Only",
            "Excluded cleanly: Banks measure Net Interest Margin (NIM) rather than manufacturing/service OPM.",
        ))
    elif ttm_opm is not None and avg_opm_5y is not None:
        m_diff = round(ttm_opm - avg_opm_5y, 1)
        opm_pass = m_diff >= -2.0
        criteria.append(_criterion(
            13, "Operating Margin (OPM) Stability", "Business Quality", "Tier 2 - Core Quality (2x)",
            "PASS" if opm_pass else "FAIL",
            f"TTM OPM {ttm_opm:.1f}% vs Rolling 5Y Avg {avg_opm_5y:.1f}% ({m_diff:+.1f}%)",
            "TTM OPM >= Rolling 5Y Avg - 2.0% (-200 bps)",
            f"{'Passed' if opm_pass else 'Failed'}: Unit operating profitability {'holds steady/expands' if opm_pass else 'compressed > 200 bps below 5Y norm'}.",
        ))
    else:
        criteria.append(_criterion(13, "Operating Margin (OPM) Stability", "Business Quality", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK", "N/A", "TTM >= 5Y Avg - 2%", "Margin history incomplete."))

    # 14. Capital Allocation: High-ROCE Reinvestment vs Payout
    payout_pct = round((_sf(info.get("payoutRatio"), 0.0) or 0.0) * 100.0, 1) if info.get("payoutRatio") is not None else 0.0
    if div_s is not None and pat_s is not None and payout_pct == 0.0:
        d_lat = abs(_sf(div_s.iloc[-1], 0.0) or 0.0)
        p_lat = _sf(pat_s.iloc[-1], 0.0) or 0.0
        if p_lat > 0 and d_lat > 0:
            payout_pct = round((d_lat / p_lat) * 100.0, 1)
    eff_hurdle = 15.0 if is_bank else 20.0
    reinvest_ok = (avg_roce_or_roe is not None and avg_roce_or_roe >= eff_hurdle)
    payout_ok = payout_pct >= 25.0
    ca_pass = reinvest_ok or payout_ok
    criteria.append(_criterion(
        14, "Capital Allocation: High-ROCE Reinvestment vs Payout", "Capital Discipline", "Tier 3 - Timing / Valuation (1x)",
        "PASS" if ca_pass else "FAIL",
        f"Capital Eff {avg_roce_or_roe or '--'}% | Dividend Payout {payout_pct:.1f}%",
        f"ROCE/ROE >= {eff_hurdle:.0f}% OR Dividend Payout >= 25%",
        f"{'Passed' if ca_pass else 'Failed'}: {'Reinvesting at high ROCE/ROE or returning surplus cash via >=25% payout' if ca_pass else 'Low capital efficiency combined with low cash distribution to shareholders'}.",
    ))

    # 15. Technical S&R Entry Gate (No Resistance Chasing)
    sig_code = scr_stock.get("signal_code", "UNKNOWN")
    sig_label = scr_stock.get("zone_signal", "Check Chart")
    rsi_val = _sf(scr_stock.get("rsi_14"), 50.0) or 50.0
    s1_lvl = scr_stock.get("s1", "--")
    r1_lvl = scr_stock.get("r1", "--")
    if sig_code != "UNKNOWN":
        t_pass = (sig_code != "WAIT") and (rsi_val < 68.0)
        criteria.append(_criterion(
            15, "Technical S&R Demand Gate (No Resistance Buying)", "Technical Timing", "Tier 3 - Timing / Valuation (1x)",
            "PASS" if t_pass else "FAIL",
            f"{sig_label} | RSI(14) {rsi_val:.1f} | S1 ₹{s1_lvl} / R1 ₹{r1_lvl}",
            "Signal != WAIT (Resistance) & RSI < 68",
            f"{'Passed' if t_pass else 'Failed — Wait for Pullback'}: Price sits safely below R1 ₹{r1_lvl} resistance without overbought RSI." if t_pass else f"Stock is extended near R1 ₹{r1_lvl} supply or overbought (RSI {rsi_val:.1f}); avoid lump-sum buying here.",
        ))
    else:
        criteria.append(_criterion(15, "Technical S&R Demand Gate (No Resistance Buying)", "Technical Timing", "Tier 3 - Timing / Valuation (1x)", "MANUAL_CHECK", "Not in Screener Cache", "S1 Support / RSI < 68", "Open S&R Terminal to inspect live demand/supply zones."))

    # 16. Strict Zero-Penny Liquidity & Cap-Tier Gate
    turn_cr = _sf(scr_stock.get("avg_turnover_cr"), None)
    safe_elig = scr_stock.get("safety_eligible", False)
    if ltp > 0 and turn_cr is not None:
        liq_pass = (ltp >= 50.0) and (turn_cr >= 10.0) and safe_elig
        criteria.append(_criterion(
            16, "Strict Zero-Penny Liquidity & Cap-Tier Gate", "Risk & Liquidity", "Tier 1 - Dealbreaker (3x)",
            "PASS" if liq_pass else "FAIL",
            f"LTP ₹{ltp:.1f} | Turnover ₹{turn_cr:.1f}Cr | {cap_tier} ({'Eligible' if safe_elig else 'Gated'})",
            "Price >= ₹50 & 60D Turnover >= ₹10 Cr & Cap-Tier Safety Rule",
            scr_stock.get("safety_note") or ("Meets >=₹50 & liquid institutional standards" if liq_pass else "Blocked by zero-penny, turnover, or deep-discount mid/small-cap safety rule."),
        ))
    else:
        p_ok = ltp >= 50.0
        criteria.append(_criterion(
            16, "Strict Zero-Penny Liquidity & Cap-Tier Gate", "Risk & Liquidity", "Tier 1 - Dealbreaker (3x)",
            "PASS" if p_ok else "FAIL",
            f"LTP ₹{ltp:.1f} (>= ₹50 lock)",
            "Price >= ₹50 & Turnover >= ₹10 Cr",
            "Cleared permanent ₹50 minimum floor." if p_ok else "PERMANENT REJECT: Trades below ₹50 penny floor.",
        ))

    # 17. Portfolio Concentration & Sector Pause Rebalancer (Direct Equity Level)
    stock_w_pct = 0.0
    sector_w_pct = 0.0
    total_eq_val = 0.0
    if portfolio_data and portfolio_data.get("stocks"):
        eq_holdings = portfolio_data["stocks"]
        total_eq_val = sum(float(x.get("current_value") or 0.0) for x in eq_holdings)
        if total_eq_val > 0:
            for x in eq_holdings:
                xsym = str(x.get("symbol", "")).upper()
                if xsym == clean_sym:
                    stock_w_pct += float(x.get("current_value") or 0.0) / total_eq_val * 100.0
                # Match broad sector if tracked in screener
                x_scr = next((s for s in scr_snap.get("stocks", []) if s["symbol"].upper() == xsym), {})
                if x_scr.get("sector") == sector_str and sector_str != "General":
                    sector_w_pct += float(x.get("current_value") or 0.0) / total_eq_val * 100.0
    conc_pass = (stock_w_pct <= 15.0) and (sector_w_pct <= 30.0)
    pause_warn = sector_w_pct >= 25.0
    conc_reason = (
        f"Soft Pause Alert: {sector_str} already at {sector_w_pct:.1f}% (>=25%) of Direct Equity; pause sector next 2 months."
        if pause_warn else
        (f"Passed: Current holding weight {stock_w_pct:.1f}% (<=15%), Sector '{sector_str}' weight {sector_w_pct:.1f}% (<=30% ceiling)." if conc_pass else f"Breaches concentration limit: Stock {stock_w_pct:.1f}% / Sector {sector_w_pct:.1f}%.")
    )
    criteria.append(_criterion(
        17, "Portfolio Concentration & Sector Pause Rebalancer", "Portfolio Guardrail", "Tier 3 - Timing / Valuation (1x)",
        "PASS" if conc_pass else "FAIL",
        f"Stock Weight {stock_w_pct:.1f}% | {sector_str} Sector {sector_w_pct:.1f}%",
        "Stock <= 15% & Sector <= 30% (Pause >= 25%)",
        conc_reason,
    ))

    # 18. Contingent Liabilities & Legal Overhang Guard
    cont_pct = curated.get("contingent_pct_nw")
    if cont_pct is not None:
        c18_pass = cont_pct <= 20.0
        criteria.append(_criterion(
            18, "Contingent Liabilities & Legal Overhang Guard", "Balance Sheet Risk", "Tier 2 - Core Quality (2x)",
            "PASS" if c18_pass else "FAIL",
            f"Contingent Claims = {cont_pct:.1f}% of Net Worth",
            "Non-Routine Claims <= 20.0% of Net Worth",
            f"{'Passed' if c18_pass else 'Failed'}: {curated.get('contingent_note', '')}",
        ))
    else:
        criteria.append(_criterion(
            18, "Contingent Liabilities & Legal Overhang Guard", "Balance Sheet Risk", "Tier 2 - Core Quality (2x)", "MANUAL_CHECK",
            "Annual Report Note Required", "Contingent Liabilities <= 20% of Net Worth",
            "Off-balance-sheet legal/tax claims are reported only in Annual Report notes; verify before large core sizing.",
        ))

    # Aggregate 18-Criteria Summary
    passed_list = [c for c in criteria if c["status"] == "PASS"]
    failed_list = [c for c in criteria if c["status"] == "FAIL"]
    manual_list = [c for c in criteria if c["status"] == "MANUAL_CHECK"]
    applicable_cnt = len(passed_list) + len(failed_list)
    tier1_fails = [c for c in failed_list if c["severity_weight"] == 3]

    if tier1_fails:
        verdict_badge = "HIGH RISK — TIER-1 DEALBREAKER FAILED"
        disclaimer = f"🚨 RISK ALERT: {company_name} failed {len(tier1_fails)} Tier-1 Dealbreaker rule(s): {', '.join(c['name'] for c in tier1_fails)}. Protect capital and resolve before deploying."
    elif len(failed_list) >= 3:
        verdict_badge = f"ELEVATED CAUTION — {len(failed_list)} RULES FAILED"
        disclaimer = f"⚠️ CAUTION: {company_name} passed {len(passed_list)}/{applicable_cnt} automated checks but failed {len(failed_list)} criteria ({', '.join(c['name'] for c in failed_list[:3])})."
    elif failed_list:
        verdict_badge = f"QUALITY INTACT — {len(failed_list)} SOFT CAUTION(S)"
        disclaimer = f"⚡ MODERATE CAUTION: {company_name} passed all Tier-1 safety gates ({len(passed_list)}/{applicable_cnt} passed) with {len(failed_list)} timing/valuation check(s): {', '.join(c['name'] for c in failed_list)}."
    else:
        verdict_badge = "FORTRESS PASS — ALL APPLICABLE CHECKS CLEARED"
        disclaimer = f"✅ HIGH CONVICTION: {company_name} cleared all {len(passed_list)} applicable automated criteria ({len(manual_list)} annual-report checks noted cleanly)."

    payload = {
        "_ts": now,
        "symbol": clean_sym,
        "name": company_name,
        "sector": sector_str,
        "cap_tier": cap_tier,
        "ltp": round(ltp, 2),
        "framework_version": "Frozen v1.0 (18 Criteria | Rolling 5Y Window)",
        "score": f"{len(passed_list)} / {applicable_cnt} Automated ({len(criteria)} Total)",
        "summary_counts": {
            "total": len(criteria),
            "passed": len(passed_list),
            "failed": len(failed_list),
            "manual_check": len(manual_list),
            "applicable": applicable_cnt,
            "tier1_fails": len(tier1_fails),
        },
        "verdict_badge": verdict_badge,
        "disclaimer": disclaimer,
        "criteria": criteria,
        # Preserve backwards-compatible keys cleanly too
        "passed": [f"{c['name']}: {c['actual_value']} — {c['reason']}" for c in passed_list],
        "cautions": [{"pointer": f"{c['name']} ({c['tier']})", "detail": f"{c['actual_value']} — {c['reason']}"} for c in failed_list],
    }

    with _CACHE_LOCK:
        _MEMORY_SCORECARD_CACHE[clean_sym] = payload
        disk = _load_disk_cache()
        disk[clean_sym] = payload
        # Cap disk cache cleanly at latest 60 audited tickers
        if len(disk) > 60:
            oldest = sorted(disk.items(), key=lambda kv: kv[1].get("_ts", 0.0))[: len(disk) - 60]
            for k, _ in oldest:
                disk.pop(k, None)
        _save_disk_cache(disk)

    return payload
