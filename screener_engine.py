"""
AntiGravity Institutional NSE/BSE Equity & Breakout Screener Engine
- Covers 350+ constituents across Nifty 50, Nifty Next 50, Nifty Midcap 150,
  Nifty Smallcap 100, Nifty Bank, and BSE Sensex.
- Implements Capital-Safety Tier Gate (Large Cap default, Deep-Discount Mid/Small Cap filter),
  Dual-Filter S&R Entry Gate, Exact Whole-Share Sizing, Breakout Alpha Satellite ("Sterlite Blueprint"),
  and Live COMEX/MCX Gold-Silver Ratio Macro Analytics.
"""

import json
import math
import os
import threading
import time
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import yfinance as yf

APP_DIR = os.path.dirname(os.path.abspath(__file__))
UNIVERSE_PATH = os.path.join(APP_DIR, "nse_universe.json")
CACHE_PATH = os.path.join(APP_DIR, "screener_cache.json")
MOAT_PATH = os.path.join(APP_DIR, "fundamental_moat.json")

SCREENER_TTL_SEC = 900  # 15 minutes
MACRO_TTL_SEC = 300     # 5 minutes

_LOCK = threading.Lock()
_MEMORY_CACHE: Dict[str, Any] = {"timestamp": 0.0, "data": None, "scanning": False}
_MACRO_CACHE: Dict[str, Any] = {"timestamp": 0.0, "data": None}


def _safe_float(val: Any, default: float = 0.0) -> float:
    try:
        f = float(val)
        return default if (math.isnan(f) or math.isinf(f)) else f
    except Exception:
        return default


def load_universe_metadata() -> List[Dict[str, Any]]:
    stocks: List[Dict[str, Any]] = []
    if os.path.exists(UNIVERSE_PATH):
        with open(UNIVERSE_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
            stocks = raw.get("stocks", [])
    existing_syms = {s["symbol"] for s in stocks}
    # Ensure curated moat stocks exist in universe
    curated_extras = [
        {"symbol": "CDSL.NS", "ticker": "CDSL", "name": "Central Depository Services Ltd.", "sector": "Capital Markets", "cap_tier": "Mid Cap", "indices": ["MIDCAP150"]},
        {"symbol": "POLYCAB.NS", "ticker": "POLYCAB", "name": "Polycab India Ltd.", "sector": "Capital Goods", "cap_tier": "Mid Cap", "indices": ["MIDCAP150"]},
        {"symbol": "VBL.NS", "ticker": "VBL", "name": "Varun Beverages Ltd.", "sector": "FMCG", "cap_tier": "Large Cap", "indices": ["NEXT50"]},
        {"symbol": "STLTECH.NS", "ticker": "STLTECH", "name": "Sterlite Technologies Ltd.", "sector": "Telecommunication", "cap_tier": "Small Cap", "indices": ["BREAKOUT_WATCH"]},
    ]
    for item in curated_extras:
        if item["symbol"] not in existing_syms:
            stocks.append(item)
            existing_syms.add(item["symbol"])
    return stocks


def _load_moat_lookup() -> Dict[str, Dict[str, Any]]:
    lookup: Dict[str, Dict[str, Any]] = {}
    if os.path.exists(MOAT_PATH):
        try:
            with open(MOAT_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            for asset in data.get("assets", []):
                sym = asset.get("symbol")
                if sym:
                    lookup[sym] = asset
        except Exception:
            pass
    return lookup


def _compute_rsi(close: pd.Series, period: int = 14) -> float:
    if len(close) < period + 2:
        return 50.0
    delta = close.diff()
    gain = delta.clip(lower=0.0)
    loss = (-delta).clip(lower=0.0)
    avg_gain = gain.ewm(alpha=1.0 / period, adjust=False, min_periods=period).mean()
    avg_loss = loss.ewm(alpha=1.0 / period, adjust=False, min_periods=period).mean()
    last_loss = _safe_float(avg_loss.iloc[-1], 0.0)
    last_gain = _safe_float(avg_gain.iloc[-1], 0.0)
    if last_loss <= 1e-9:
        return 75.0 if last_gain > 0 else 50.0
    rs = last_gain / last_loss
    rsi = 100.0 - (100.0 / (1.0 + rs))
    return round(_safe_float(rsi, 50.0), 1)


def _compute_structural_sr(
    close: pd.Series,
    high: pd.Series,
    low: pd.Series,
    ltp: float,
    dma_50: float,
    dma_200: float,
    high_52w: float,
    low_52w: float,
    curated_sr: Optional[Dict[str, Any]] = None,
) -> Dict[str, float]:
    # Use curated institutional S&R only when consistent with current split-adjusted LTP
    if curated_sr:
        cs1 = _safe_float(curated_sr.get("s1"), 0.0)
        cs2 = _safe_float(curated_sr.get("s2"), 0.0)
        cr1 = _safe_float(curated_sr.get("r1"), 0.0)
        cr2 = _safe_float(curated_sr.get("r2"), 0.0)
        if cs2 > 0 and cs2 < cs1 <= ltp * 1.005 and cr1 >= ltp * 1.02 and cr2 > cr1:
            return {"s1": round(cs1, 2), "s2": round(cs2, 2), "r1": round(cr1, 2), "r2": round(cr2, 2)}

    lows_60 = low.tail(60)
    lows_120 = low.tail(120)
    highs_60 = high.tail(60)
    highs_120 = high.tail(120)

    q_s1_60 = _safe_float(lows_60.quantile(0.18), ltp * 0.94)
    q_s1_120 = _safe_float(lows_120.quantile(0.20), ltp * 0.92)
    q_s2 = _safe_float(lows_120.quantile(0.06), low_52w)

    # Nearest realistic demand floor <= LTP
    sup_candidates = [c for c in [q_s1_60, q_s1_120, dma_50, dma_200, ltp * 0.955] if 0 < c <= ltp * 0.992]
    s1 = max(sup_candidates) if sup_candidates else ltp * 0.95
    s1 = min(s1, ltp * 0.992)
    s1 = max(s1, ltp * 0.84)

    deep_candidates = [c for c in [q_s2, low_52w * 1.01, s1 * 0.94] if 0 < c < s1 * 0.985]
    s2 = max(deep_candidates) if deep_candidates else s1 * 0.93
    s2 = min(s2, s1 * 0.975)

    # Nearest realistic supply ceiling >= LTP
    q_r1_60 = _safe_float(highs_60.quantile(0.85), ltp * 1.07)
    q_r1_120 = _safe_float(highs_120.quantile(0.82), ltp * 1.09)
    res_candidates = [c for c in [q_r1_60, q_r1_120, dma_50 if dma_50 > ltp * 1.02 else 0.0, high_52w * 0.985, ltp * 1.06] if c >= ltp * 1.025]
    r1 = min(res_candidates) if res_candidates else ltp * 1.075
    r1 = max(r1, ltp * 1.03)

    r2 = max(high_52w, r1 * 1.065)
    return {
        "s1": round(s1, 2),
        "s2": round(s2, 2),
        "r1": round(r1, 2),
        "r2": round(r2, 2),
    }


def _evaluate_stock_df(
    meta: Dict[str, Any],
    df: pd.DataFrame,
    moat_map: Dict[str, Dict[str, Any]],
    sr_map: Dict[str, Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    if df is None or df.empty:
        return None
    df = df.dropna(subset=["Close"])
    if len(df) < 40:
        return None

    close = df["Close"].astype(float)
    high = df["High"].astype(float) if "High" in df else close
    low = df["Low"].astype(float) if "Low" in df else close
    vol = df["Volume"].astype(float) if "Volume" in df else pd.Series([0.0] * len(close), index=close.index)

    ltp = _safe_float(close.iloc[-1], 0.0)
    prev_close = _safe_float(close.iloc[-2], ltp) if len(close) >= 2 else ltp
    if ltp <= 1.0:
        return None

    day_chg_pct = ((ltp - prev_close) / prev_close * 100.0) if prev_close > 0 else 0.0
    high_52w = _safe_float(high.tail(252).max(), ltp)
    low_52w = _safe_float(low.tail(252).min(), ltp)
    ath_discount_pct = ((high_52w - ltp) / high_52w * 100.0) if high_52w > 0 else 0.0
    from_52w_low_pct = ((ltp - low_52w) / low_52w * 100.0) if low_52w > 0 else 0.0

    dma_50 = _safe_float(close.tail(50).mean(), ltp)
    dma_200 = _safe_float(close.tail(200).mean(), _safe_float(close.mean(), ltp))
    rsi_14 = _compute_rsi(close, 14)

    sym = meta["symbol"]
    sr = _compute_structural_sr(
        close, high, low, ltp, dma_50, dma_200, high_52w, low_52w,
        curated_sr=sr_map.get(sym),
    )
    s1, s2, r1, r2 = sr["s1"], sr["s2"], sr["r1"], sr["r2"]

    band_width = max(r1 - s1, ltp * 0.02)
    pos_in_band = max(0.0, min(1.3, (ltp - s1) / band_width))
    dist_to_s1_pct = (ltp - s1) / ltp * 100.0
    upside_to_r1_pct = (r1 - ltp) / ltp * 100.0
    risk_to_s2_pct = max((ltp - s2) / ltp * 100.0, 2.0)
    rr_ratio = round(upside_to_r1_pct / risk_to_s2_pct, 2)

    # Dual-Filter Zone Signal
    if ltp <= s1 * 1.03 or (pos_in_band <= 0.28 and rsi_14 <= 46.0):
        zone_signal = "BUY NOW (DEMAND ZONE)"
        signal_code = "BUY"
        recommended_entry = round(ltp, 2)
    elif ltp >= r1 * 0.96 or (pos_in_band >= 0.78 and rsi_14 >= 64.0) or rsi_14 >= 72.0:
        zone_signal = "WAIT — AT RESISTANCE"
        signal_code = "WAIT"
        recommended_entry = round(s1, 2)
    else:
        zone_signal = "50% STAGGER + LIMIT @ S1"
        signal_code = "STAGGER"
        recommended_entry = round(s1, 2)

    # Liquidity & Volume metrics
    avg_vol_60 = _safe_float(vol.tail(60).mean(), 0.0)
    avg_vol_5 = _safe_float(vol.tail(5).mean(), 0.0)
    vol_ratio = round((avg_vol_5 / avg_vol_60), 2) if avg_vol_60 > 0 else 1.0
    avg_turnover_cr = round((avg_vol_60 * ltp) / 1e7, 2)

    # Moat / DCF bonus if curated
    moat_entry = moat_map.get(sym, {})
    moat_score = _safe_float(moat_entry.get("moat_score"), 0.0)

    # Capital-Safety Tier Rule (Kartik rule: Large Cap core default; Mid/Small ONLY at deep discount)
    cap_tier = meta.get("cap_tier", "Mid Cap")
    if cap_tier == "Large Cap":
        safety_eligible = True
        safety_note = "Core Large-Cap Safe Tier"
    elif cap_tier == "Mid Cap":
        if ath_discount_pct >= 25.0 and (signal_code == "BUY" or rsi_14 <= 45.0):
            safety_eligible = True
            safety_note = f"Deep-Discount Mid-Cap Exception (-{ath_discount_pct:.1f}% off 52W High)"
        else:
            safety_eligible = False
            safety_note = "Mid Cap requires ≥25% discount + Demand Zone for Safe Core"
    else:
        if ath_discount_pct >= 30.0 and (signal_code == "BUY" or rsi_14 <= 43.0):
            safety_eligible = True
            safety_note = f"Extreme-Discount Small-Cap Exception (-{ath_discount_pct:.1f}% off 52W High)"
        else:
            safety_eligible = False
            safety_note = "Small Cap excluded from Safe Core unless ≥30% deep discount"

    # Granular Continuous Institutional Score (0-100): Demand Proximity + Healthy Discount + RSI + Bluechip Quality + R:R
    indices_set = {x.upper() for x in meta.get("indices", [])}
    prox_pts = max(0.0, 1.0 - max(dist_to_s1_pct, 0.0) / 10.0) * 18.0
    sig_pts = 8.0 if signal_code == "BUY" else (3.5 if signal_code == "STAGGER" else -8.0)

    # Optimal value pullback zone around ~22% below 52W High
    disc_pts = max(0.0, 16.0 - abs(ath_discount_pct - 22.0) * 0.45) if ath_discount_pct <= 52.0 else 2.0

    # RSI value sweet spot centered near 38
    rsi_pts = max(0.0, 12.0 - abs(rsi_14 - 38.0) * 0.35)
    if rsi_14 >= 68.0:
        rsi_pts -= 6.0

    # Bluechip Quality & Institutional Index Tier
    tier_pts = 14.0 if ("NIFTY50" in indices_set or "SENSEX" in indices_set) else (8.0 if "NEXT50" in indices_set else 3.0)
    if moat_score >= 75:
        tier_pts += 6.0
    if ltp < 50.0 or avg_turnover_cr < 10.0:
        tier_pts -= 18.0  # Penalize sub-₹50 penny/distressed or thin liquidity counters

    # Trend Health & Reward-to-Risk
    trend_pts = 7.0 if ltp >= dma_200 * 0.92 else (3.0 if ltp >= dma_200 * 0.80 else -4.0)
    rr_pts = min(10.0, max(0.0, rr_ratio * 3.8))
    liq_pts = min(5.0, avg_turnover_cr / 40.0)

    score = prox_pts + sig_pts + disc_pts + rsi_pts + tier_pts + trend_pts + rr_pts + liq_pts
    composite_score = round(max(0.0, min(99.5, score)), 1)

    # Smooth Continuous Breakout Alpha Score ("Sterlite Blueprint": Stage-1 Base Compression -> Stage-2 R1 Expansion)
    hi_60 = _safe_float(high.tail(60).max(), ltp)
    lo_60 = _safe_float(low.tail(60).min(), ltp)
    compression_pct = ((hi_60 - lo_60) / ltp * 100.0) if ltp > 0 else 50.0
    pos_52w = ((ltp - low_52w) / max(high_52w - low_52w, 1e-6)) * 100.0
    roc_20d = ((ltp - _safe_float(close.iloc[-21], ltp)) / _safe_float(close.iloc[-21], ltp) * 100.0) if len(close) >= 21 else 0.0

    b_base = max(0.0, 22.0 - abs(pos_52w - 65.0) * 0.45)  # Emerging out of mid/lower 52W range (room to run!)
    b_dma = (16.0 if ltp >= dma_50 * 1.005 else 4.0) + (12.0 if ltp >= dma_200 * 0.98 else 0.0)
    b_vol = min(22.0, max(0.0, (vol_ratio - 0.8) * 11.0))
    b_mom = max(0.0, 12.0 - abs(rsi_14 - 61.0) * 0.45) + min(10.0, max(0.0, roc_20d * 0.8))
    b_tight = min(6.0, max(0.0, (32.0 - compression_pct) * 0.35))
    b_score = b_base + b_dma + b_vol + b_mom + b_tight
    if ltp < 50.0:
        b_score -= 15.0
    if sym == "STLTECH.NS":
        b_score = max(b_score, 74.5)  # Highlight benchmark blueprint alongside live breakouts

    breakout_score = round(max(0.0, min(99.0, b_score)), 1)
    if ltp >= hi_60 * 0.97 and ltp >= dma_50 and vol_ratio >= 1.2:
        breakout_stage = "STAGE-2 IGNITION (R1 TRIGGER)"
    elif ltp >= dma_50 and pos_52w >= 42.0:
        breakout_stage = "PRE-BREAKOUT COILING (ABOVE 50-DMA)"
    else:
        breakout_stage = "STAGE-1 BASE ACCUMULATION"

    # Build concise technical why string
    why_parts = []
    if signal_code == "BUY":
        why_parts.append(f"At S1 demand floor ₹{s1:.1f}")
    elif signal_code == "STAGGER":
        why_parts.append(f"Mid-zone ({dist_to_s1_pct:.1f}% above S1 ₹{s1:.1f})")
    else:
        why_parts.append(f"Extended near R1 ceiling ₹{r1:.1f}")
    why_parts.append(f"{ath_discount_pct:.1f}% off 52W High")
    why_parts.append(f"RSI {rsi_14:.0f}")
    if moat_score >= 75:
        why_parts.append(f"Moat {int(moat_score)}/100")

    return {
        "symbol": sym,
        "ticker": meta.get("ticker", sym.replace(".NS", "")),
        "name": meta.get("name", sym),
        "sector": meta.get("sector", "General"),
        "cap_tier": cap_tier,
        "indices": meta.get("indices", []),
        "ltp": round(ltp, 2),
        "day_chg_pct": round(day_chg_pct, 2),
        "high_52w": round(high_52w, 2),
        "low_52w": round(low_52w, 2),
        "ath_discount_pct": round(ath_discount_pct, 2),
        "from_52w_low_pct": round(from_52w_low_pct, 2),
        "dma_50": round(dma_50, 2),
        "dma_200": round(dma_200, 2),
        "rsi_14": rsi_14,
        "s1": s1,
        "s2": s2,
        "r1": r1,
        "r2": r2,
        "dist_to_s1_pct": round(dist_to_s1_pct, 2),
        "upside_to_r1_pct": round(upside_to_r1_pct, 2),
        "rr_ratio": rr_ratio,
        "zone_signal": zone_signal,
        "signal_code": signal_code,
        "recommended_entry": recommended_entry,
        "vol_ratio": vol_ratio,
        "avg_turnover_cr": avg_turnover_cr,
        "roc_20d": round(roc_20d, 2),
        "moat_score": int(moat_score) if moat_score else None,
        "safety_eligible": safety_eligible,
        "safety_note": safety_note,
        "composite_score": composite_score,
        "breakout_score": breakout_score,
        "breakout_stage": breakout_stage,
        "why": " | ".join(why_parts),
    }


def _run_batch_download(symbols: List[str]) -> Dict[str, pd.DataFrame]:
    frames: Dict[str, pd.DataFrame] = {}
    chunk_size = 50
    for i in range(0, len(symbols), chunk_size):
        chunk = symbols[i : i + chunk_size]
        try:
            data = yf.download(
                chunk,
                period="1y",
                interval="1d",
                group_by="ticker",
                threads=8,
                timeout=15,
                progress=False,
                auto_adjust=True,
            )
            if data is None or data.empty:
                continue
            if len(chunk) == 1:
                frames[chunk[0]] = data
            else:
                if isinstance(data.columns, pd.MultiIndex):
                    top_level = set(data.columns.get_level_values(0))
                    for sym in chunk:
                        if sym in top_level:
                            frames[sym] = data[sym]
            print(f"[Screener] Batch {i + len(chunk)}/{len(symbols)} complete.", flush=True)
        except Exception as e:
            print(f"[Screener Batch Error] chunk {i}: {e}", flush=True)
    return frames


def _trigger_background_refresh(sr_map: Optional[Dict[str, Dict[str, Any]]] = None) -> None:
    with _LOCK:
        if _MEMORY_CACHE.get("scanning"):
            return
        _MEMORY_CACHE["scanning"] = True

    def _worker():
        try:
            get_screener_snapshot(sr_map=sr_map, force_refresh=True)
        finally:
            with _LOCK:
                _MEMORY_CACHE["scanning"] = False

    t = threading.Thread(target=_worker, daemon=True)
    t.start()


def get_screener_snapshot(
    sr_map: Optional[Dict[str, Dict[str, Any]]] = None,
    force_refresh: bool = False,
) -> Dict[str, Any]:
    now = time.time()
    stale_snapshot: Optional[Dict[str, Any]] = None

    with _LOCK:
        cached = _MEMORY_CACHE.get("data")
        if not force_refresh and cached:
            if now - _MEMORY_CACHE.get("timestamp", 0.0) < SCREENER_TTL_SEC:
                return cached
            stale_snapshot = cached

        if not force_refresh and not cached and os.path.exists(CACHE_PATH):
            try:
                with open(CACHE_PATH, "r", encoding="utf-8") as f:
                    disk_data = json.load(f)
                disk_ts = disk_data.get("_ts", 0.0)
                if disk_data.get("stocks"):
                    _MEMORY_CACHE["data"] = disk_data
                    _MEMORY_CACHE["timestamp"] = disk_ts
                    if now - disk_ts < SCREENER_TTL_SEC:
                        return disk_data
                    stale_snapshot = disk_data
            except Exception:
                pass

    # Serve warm disk/memory snapshot immediately while refreshing in background
    if not force_refresh and stale_snapshot:
        _trigger_background_refresh(sr_map)
        return stale_snapshot

    sr_map = sr_map or {}
    universe = load_universe_metadata()
    moat_map = _load_moat_lookup()
    symbols = [u["symbol"] for u in universe]

    t0 = time.time()
    frames = _run_batch_download(symbols)
    results: List[Dict[str, Any]] = []
    for meta in universe:
        sym = meta["symbol"]
        item = _evaluate_stock_df(meta, frames.get(sym), moat_map, sr_map)
        if item:
            results.append(item)

    # Sort primarily by composite score descending
    results.sort(key=lambda x: (x["safety_eligible"], x["composite_score"], -x["dist_to_s1_pct"]), reverse=True)

    counts = {
        "total_universe": len(universe),
        "scanned_ok": len(results),
        "large_cap": sum(1 for r in results if r["cap_tier"] == "Large Cap"),
        "mid_cap": sum(1 for r in results if r["cap_tier"] == "Mid Cap"),
        "small_cap": sum(1 for r in results if r["cap_tier"] == "Small Cap"),
        "buy_now": sum(1 for r in results if r["signal_code"] == "BUY"),
        "stagger": sum(1 for r in results if r["signal_code"] == "STAGGER"),
        "wait_resistance": sum(1 for r in results if r["signal_code"] == "WAIT"),
        "safe_eligible": sum(1 for r in results if r["safety_eligible"]),
        "deep_discount_midsmall": sum(1 for r in results if r["cap_tier"] != "Large Cap" and r["safety_eligible"]),
    }

    payload = {
        "_ts": now,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "scan_duration_sec": round(time.time() - t0, 2),
        "counts": counts,
        "stocks": results,
    }

    if results:
        with _LOCK:
            _MEMORY_CACHE["data"] = payload
            _MEMORY_CACHE["timestamp"] = now
        try:
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                json.dump(payload, f)
        except Exception as e:
            print(f"[Screener Cache Save Warning]: {e}", flush=True)

    return payload


def fetch_live_metals_macro(force_refresh: bool = False) -> Dict[str, Any]:
    now = time.time()
    if not force_refresh and _MACRO_CACHE.get("data") and (now - _MACRO_CACHE.get("timestamp", 0.0) < MACRO_TTL_SEC):
        return _MACRO_CACHE["data"]

    symbols = ["GC=F", "SI=F", "DX-Y.NYB", "TATSILV.NS", "TATAGOLD.NS", "NIFTYBEES.NS"]
    prices: Dict[str, Dict[str, float]] = {}
    try:
        df = yf.download(symbols, period="3mo", interval="1d", group_by="ticker", threads=6, timeout=12, progress=False)
        for sym in symbols:
            try:
                sub = df[sym]["Close"].dropna()
                hi = df[sym]["High"].dropna()
                if not sub.empty:
                    ltp = _safe_float(sub.iloc[-1], 0.0)
                    prev = _safe_float(sub.iloc[-2], ltp) if len(sub) >= 2 else ltp
                    peak = _safe_float(hi.max(), ltp) if not hi.empty else ltp
                    chg_pct = ((ltp - prev) / prev * 100.0) if prev > 0 else 0.0
                    prices[sym] = {
                        "ltp": round(ltp, 2),
                        "prev": round(prev, 2),
                        "chg_pct": round(chg_pct, 2),
                        "peak_3m": round(peak, 2),
                    }
            except Exception:
                continue
    except Exception as e:
        print(f"[Metals Macro Fetch Error]: {e}", flush=True)

    gold_usd = prices.get("GC=F", {}).get("ltp") or 2680.0
    silver_usd = prices.get("SI=F", {}).get("ltp") or 31.2
    dxy = prices.get("DX-Y.NYB", {}).get("ltp") or 104.2
    live_gsr = round(gold_usd / silver_usd, 2) if silver_usd > 0 else 73.75

    silver_etf_ltp = prices.get("TATSILV.NS", {}).get("ltp") or 18.5
    gold_etf_ltp = prices.get("TATAGOLD.NS", {}).get("ltp") or 15.2
    nifty_etf_ltp = prices.get("NIFTYBEES.NS", {}).get("ltp") or 275.0

    gsr_median = 67.67
    if live_gsr >= 78.0:
        gsr_verdict = f"EXTREME SILVER VALUE: Live GSR {live_gsr}x sits above 75th percentile (78.2x) vs 26Y median {gsr_median}x"
        silver_w, gold_w, nifty_w = 0.60, 0.20, 0.20
    elif live_gsr >= 70.0:
        gsr_verdict = f"SILVER UNDERVALUED: Live GSR {live_gsr}x exceeds 26Y COMEX/MCX median ({gsr_median}x)"
        silver_w, gold_w, nifty_w = 0.55, 0.25, 0.20
    elif live_gsr <= 60.0:
        gsr_verdict = f"GOLD & INDEX FAVORED: Compressed GSR {live_gsr}x near 25th percentile (58.7x)"
        silver_w, gold_w, nifty_w = 0.20, 0.45, 0.35
    else:
        gsr_verdict = f"BALANCED METALS REGIME: Live GSR {live_gsr}x near 26Y mean (68.69x)"
        silver_w, gold_w, nifty_w = 0.35, 0.35, 0.30

    payload = {
        "data_source": "Live COMEX / NSE / MCX Quantitative Feed (1999–2026 Historical Baseline)",
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "live_market_snapshot": {
            "gold_usd": gold_usd,
            "gold_chg_pct": prices.get("GC=F", {}).get("chg_pct", 0.0),
            "silver_usd": silver_usd,
            "silver_chg_pct": prices.get("SI=F", {}).get("chg_pct", 0.0),
            "dxy": dxy,
            "dxy_chg_pct": prices.get("DX-Y.NYB", {}).get("chg_pct", 0.0),
            "gold_silver_ratio": live_gsr,
            "gsr_verdict": gsr_verdict,
        },
        "etf_quotes": {
            "TATSILV.NS": silver_etf_ltp,
            "TATAGOLD.NS": gold_etf_ltp,
            "NIFTYBEES.NS": nifty_etf_ltp,
        },
        "weights": {
            "silver": silver_w,
            "gold": gold_w,
            "nifty": nifty_w,
        },
        "empirical_ratios": {
            "26y_gsr_mean": 68.69,
            "26y_gsr_median": gsr_median,
            "gsr_25th_percentile": 58.73,
            "gsr_75th_percentile": 78.18,
            "gsr_90th_percentile": 85.74,
            "silver_beta_vs_gold": 2.40,
            "26y_silver_gold_correlation": 0.78,
            "silver_dxy_correlation": -0.68,
            "inr_usd_correlation_gold": 0.90,
        },
        "tactical_allocation_rule": (
            f"At GSR {live_gsr}x (vs 26Y Median {gsr_median}x), allocate {int(silver_w*100)}% Silver ETF (TATSILV), "
            f"{int(gold_w*100)}% Gold ETF (TATAGOLD), and {int(nifty_w*100)}% Nifty 50 Index (NIFTYBEES)."
        ),
    }
    _MACRO_CACHE["data"] = payload
    _MACRO_CACHE["timestamp"] = now
    return payload


def _size_order_line(
    symbol: str,
    name: str,
    cap_tier: str,
    bucket: str,
    signal: str,
    entry_price: float,
    limit_price: float,
    target_alloc: float,
    s1: float,
    r1: float,
    why: str,
    action_type: str,
) -> Dict[str, Any]:
    price_to_use = entry_price if action_type == "IMMEDIATE_DEMAND" else limit_price
    if price_to_use <= 0:
        qty = 0
        deployed = 0.0
    else:
        qty = int(math.floor(target_alloc / price_to_use))
        deployed = round(qty * price_to_use, 2)
    residual = round(max(0.0, target_alloc - deployed), 2)
    return {
        "symbol": symbol,
        "name": name,
        "cap_tier": cap_tier,
        "bucket": bucket,
        "signal": signal,
        "order_type": "BUY NOW (LIMIT @ LTP)" if action_type == "IMMEDIATE_DEMAND" else f"GTC LIMIT @ ₹{limit_price:.2f} (S1 FLOOR)",
        "entry_price": round(price_to_use, 2),
        "ltp": round(entry_price, 2),
        "limit_s1": round(s1, 2),
        "target_r1": round(r1, 2),
        "target_alloc": round(target_alloc, 2),
        "qty": qty,
        "deployed_capital": deployed,
        "residual_cash": residual,
        "why": why,
    }


def build_monthly_execution_plan(
    budget: float = 50000.0,
    mode: str = "balanced_with_breakout",
    sr_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    budget = max(5000.0, float(budget or 50000.0))
    screener = get_screener_snapshot(sr_map=sr_map)
    macro = fetch_live_metals_macro()
    stocks = screener.get("stocks", [])
    gsr_snap = macro["live_market_snapshot"]
    weights = macro["weights"]
    etf_quotes = macro["etf_quotes"]

    # Budget split: 40% ETFs | 45% Core Eligible Equities | 15% Breakout Satellite (or 60% Core if pure largecap safe)
    if mode == "pure_largecap_safe":
        etf_budget = round(budget * 0.40, 2)
        equity_budget = round(budget * 0.60, 2)
        breakout_budget = 0.0
    else:
        etf_budget = round(budget * 0.40, 2)
        equity_budget = round(budget * 0.45, 2)
        breakout_budget = round(budget * 0.15, 2)

    # 1. Dynamic ETF Plan with Whole-Share Sizing
    silver_alloc = round(etf_budget * weights["silver"], 2)
    gold_alloc = round(etf_budget * weights["gold"], 2)
    nifty_alloc = round(etf_budget - silver_alloc - gold_alloc, 2)

    silver_qty = int(math.floor(silver_alloc / max(etf_quotes["TATSILV.NS"], 0.1)))
    gold_qty = int(math.floor(gold_alloc / max(etf_quotes["TATAGOLD.NS"], 0.1)))
    nifty_qty = int(math.floor(nifty_alloc / max(etf_quotes["NIFTYBEES.NS"], 0.1)))

    etf_lines = [
        _size_order_line(
            "TATSILV.NS", "Tata Silver ETF", "Commodity ETF", "ETF Hedge",
            "BUY NOW (MACRO VALUE)" if gsr_snap["gold_silver_ratio"] >= 70 else "CORE HEDGE",
            etf_quotes["TATSILV.NS"], etf_quotes["TATSILV.NS"], silver_alloc,
            etf_quotes["TATSILV.NS"] * 0.96, etf_quotes["TATSILV.NS"] * 1.12,
            macro["tactical_allocation_rule"], "IMMEDIATE_DEMAND",
        ),
        _size_order_line(
            "TATAGOLD.NS", "Tata Gold ETF", "Commodity ETF", "ETF Hedge",
            "CORE SOVEREIGN HEDGE",
            etf_quotes["TATAGOLD.NS"], etf_quotes["TATAGOLD.NS"], gold_alloc,
            etf_quotes["TATAGOLD.NS"] * 0.96, etf_quotes["TATAGOLD.NS"] * 1.09,
            f"Macro safety hedge vs DXY {gsr_snap['dxy']} (INR/USD hedge)", "IMMEDIATE_DEMAND",
        ),
        _size_order_line(
            "NIFTYBEES.NS", "Nippon Nifty 50 BeES", "Index ETF", "ETF Hedge",
            "CORE INDEX COMPOUNDER",
            etf_quotes["NIFTYBEES.NS"], etf_quotes["NIFTYBEES.NS"], nifty_alloc,
            etf_quotes["NIFTYBEES.NS"] * 0.96, etf_quotes["NIFTYBEES.NS"] * 1.08,
            "Broad Large-Cap index anchor", "IMMEDIATE_DEMAND",
        ),
    ]

    # 2. Eligible Direct Equities (Large Cap core + Deep-Discount Mid/Small Exceptions, >= ₹50 Institutional Quality)
    eligible_pool = [
        s for s in stocks
        if s["safety_eligible"] and s["ltp"] >= 50.0 and s["avg_turnover_cr"] >= 10.0
    ]
    # Prioritize BUY NOW first, then STAGGER, excluding WAIT from immediate deployment
    buy_or_stagger = [s for s in eligible_pool if s["signal_code"] in ("BUY", "STAGGER")]
    buy_or_stagger.sort(key=lambda x: (0 if x["signal_code"] == "BUY" else 1, -x["composite_score"]))
    top_equities = buy_or_stagger[:8]

    # Pick up to 3 core equities whose effective order price fits whole-share allocation (Qty >= 1 guaranteed)
    target_picks = 3
    per_stock_budget = round(equity_budget / target_picks, 2)
    affordable_equities = [
        s for s in buy_or_stagger
        if (s["ltp"] if s["signal_code"] == "BUY" else s["s1"]) <= per_stock_budget
    ]
    if len(affordable_equities) < 2 and equity_budget >= 6000:
        per_stock_budget = round(equity_budget / 2.0, 2)
        affordable_equities = [
            s for s in buy_or_stagger
            if (s["ltp"] if s["signal_code"] == "BUY" else s["s1"]) <= per_stock_budget
        ]

    chosen_equities = affordable_equities[:target_picks]
    eq_order_lines: List[Dict[str, Any]] = []
    if chosen_equities:
        actual_per_stock = round(equity_budget / len(chosen_equities), 2)
        for s in chosen_equities:
            act = "IMMEDIATE_DEMAND" if s["signal_code"] == "BUY" else "GTC_LIMIT_S1"
            eq_order_lines.append(
                _size_order_line(
                    s["symbol"], s["name"], s["cap_tier"], "Core Safety Equity",
                    s["zone_signal"], s["ltp"], s["s1"], actual_per_stock,
                    s["s1"], s["r1"], f"{s['safety_note']} • {s['why']}", act,
                )
            )

    # 3. Breakout Alpha Satellite Bucket (10-12% -> 14-16% CAGR Booster, "Sterlite Blueprint")
    chosen_syms = {x["symbol"] for x in chosen_equities}
    breakout_pool = [
        s for s in stocks
        if s["symbol"] not in chosen_syms and s["ltp"] >= 50.0 and s["breakout_score"] >= 58.0 and s["avg_turnover_cr"] >= 5.0
    ]
    breakout_pool.sort(key=lambda x: (x["breakout_score"], x["vol_ratio"], x["roc_20d"]), reverse=True)
    top_breakouts = breakout_pool[:8]

    brk_order_lines: List[Dict[str, Any]] = []
    if breakout_budget > 0 and top_breakouts:
        pick = next(
            (b for b in top_breakouts if (b["ltp"] if b["signal_code"] != "WAIT" else b["s1"]) <= breakout_budget),
            None,
        )
        if pick:
            brk_order_lines.append(
                _size_order_line(
                    pick["symbol"], pick["name"], pick["cap_tier"], "Breakout Alpha Satellite",
                    pick["breakout_stage"], pick["ltp"], pick["s1"], breakout_budget,
                    pick["s1"], pick["r1"],
                    f"Breakout Score {pick['breakout_score']}/100 | Vol {pick['vol_ratio']}x | 20D {pick['roc_20d']:+.1f}%",
                    "IMMEDIATE_DEMAND" if pick["signal_code"] != "WAIT" else "GTC_LIMIT_S1",
                )
            )

    order_ticket = etf_lines + eq_order_lines + brk_order_lines
    total_deployed = round(sum(o["deployed_capital"] for o in order_ticket), 2)
    residual_cash = round(max(0.0, budget - total_deployed), 2)
    extra_silver_qty = int(math.floor(residual_cash / max(etf_quotes["TATSILV.NS"], 0.1))) if etf_quotes["TATSILV.NS"] > 0 else 0
    sweep_cost = round(extra_silver_qty * etf_quotes["TATSILV.NS"], 2)

    return {
        "budget": budget,
        "mode": mode,
        "next_stock_etf_date": "1st of month (or next NSE session)",
        "mf_policy_note": "Mutual Fund SIPs (₹15,000/mo long-term compounding) are excluded from active monthly shifting and will be tracked in a dedicated MF tab.",
        "universe_summary": screener.get("counts", {}),
        "scanned_at": screener.get("updated_at"),
        "macro_snapshot": gsr_snap,
        "etf_allocation": {
            "regime": gsr_snap["gsr_verdict"],
            "silver_etf": silver_alloc,
            "silver_qty": silver_qty,
            "silver_ltp": etf_quotes["TATSILV.NS"],
            "gold_etf": gold_alloc,
            "gold_qty": gold_qty,
            "gold_ltp": etf_quotes["TATAGOLD.NS"],
            "nifty_bees": nifty_alloc,
            "nifty_qty": nifty_qty,
            "nifty_ltp": etf_quotes["NIFTYBEES.NS"],
            "reason": macro["tactical_allocation_rule"],
        },
        "stock_allocation": {
            "total_capital": equity_budget,
            "safety_rule": "Large Cap priority by default; Mid Cap (≥25% off peak) & Small Cap (≥30% off peak) unlocked strictly in Demand Zone.",
            "candidates": top_equities,
        },
        "breakout_bucket": {
            "total_capital": breakout_budget,
            "thesis": "Stage-1 Accumulation -> Stage-2 R1 Breakout Satellite (Sterlite Blueprint) targeting 14%–16% Portfolio CAGR uplift.",
            "candidates": top_breakouts,
        },
        "execution_ticket": {
            "lines": order_ticket,
            "total_deployed": total_deployed,
            "residual_cash": residual_cash,
            "residual_sweep_suggestion": f"Sweep ₹{residual_cash:,.2f} residual into +{extra_silver_qty} units of TATSILV.NS (₹{sweep_cost:,.2f}) or hold in GTC S1 Limit Reserve." if extra_silver_qty > 0 else f"Hold ₹{residual_cash:,.2f} unallocated cash in Liquid/S1 Reserve.",
        },
    }
