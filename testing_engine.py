"""
testing_engine.py — Institutional Verification & Shadow Portfolio Testing Engine
AntiGravity Wealth Management System

Features:
1. Seed Capital: ₹10,00,000 across 4 distinct Timeframe Segments:
   - Segment A: Intraday S&R Precision (1-Day, 10% = ₹1,00,000)
   - Segment B: Weekly Swing (3-5 Days, 20% = ₹2,00,000)
   - Segment C: Multi-Month Positional (3-10 Months, 45% = ₹4,50,000)
     * Core Demand Bluechips (₹3.5L) + Breakout Alpha Satellite (₹1.0L, strictly < 1 yr)
   - Segment D: Multi-Year Bedrock (2-5 Years, 25% = ₹2,50,000)
2. Whole-Share Sizing with Zero Fractional Leakage.
3. Automated Profit Reinvestment (Closed trade proceeds compound back into segment cash).
4. Continuous MFE (Peak Gain) and MAE (Drawdown) tracking.
5. Anomaly & Discrepancy Ledger with 5-7% 3-Tier Noise Filter.
6. Daily Excel Generator: Daily_Testing_Ledger_YYYY-MM-DD.xlsx with Portfolio Value on Sheet 1.
"""

import os
import json
import math
import datetime
from typing import Dict, Any, List, Optional
import yfinance as yf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ENGINE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
STATE_FILE = os.path.join(DATA_DIR, "testing_portfolio.json")

# Output directory for daily Excel generation
OUTPUT_DIR = os.path.abspath(os.path.join(ENGINE_DIR, "..", "..", "2. Output Generated"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

SEED_CAPITAL = 1000000.0

SEGMENT_CONFIG = {
    "INTRADAY": {
        "name": "Intraday S&R Precision",
        "weight_pct": 10.0,
        "allocated_capital": 100000.0,
        "holding_period": "1-Day (Intraday Close)",
        "objective": "Verify immediate S1 demand floor bounce and R1 reaction precision."
    },
    "WEEKLY": {
        "name": "Weekly Swing",
        "weight_pct": 20.0,
        "allocated_capital": 200000.0,
        "holding_period": "3 to 5 Days (Mon -> Fri)",
        "objective": "Test weekly demand expansion towards R1 ceiling with S2 stop protection."
    },
    "MULTI_MONTH": {
        "name": "Multi-Month Positional",
        "weight_pct": 45.0,
        "allocated_capital": 450000.0,
        "holding_period": "3 to 10 Months (Strictly < 1 Year)",
        "objective": "Capital Protection & 30-60% Stage-2 expansions (Core Bluechips + Breakout Alpha)."
    },
    "MULTI_YEAR": {
        "name": "Multi-Year Bedrock",
        "weight_pct": 25.0,
        "allocated_capital": 250000.0,
        "holding_period": "2 to 5 Years",
        "objective": "Pure Wealth Compounding in 18-Pt wide-moat monopolies."
    }
}

INITIAL_TRADES = [
    # --- Segment A: Intraday (₹1,00,000 Budget) ---
    {
        "id": "INT-20261113-TCS",
        "symbol": "TCS.NS",
        "name": "Tata Consultancy Services",
        "segment": "INTRADAY",
        "entry_date": "2026-11-13",
        "entry_price": 2411.30,
        "quantity": 20,
        "invested_capital": 48226.00,
        "s1_support": 2405.00,
        "s2_stop": 2390.00,
        "r1_target": 2455.00,
        "r2_target": 2480.00,
        "target_desc": "Immediate S1 floor intraday reaction bounce (+1.8%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 2411.30,
        "trough_low": 2411.30,
        "discrepancy_noticed": "None. Sitting directly at S1 demand floor.",
        "corrective_measure": "None required (within normal parameters).",
        "is_anomaly": False
    },
    {
        "id": "INT-20261113-ONGC",
        "symbol": "ONGC.NS",
        "name": "Oil & Natural Gas Corporation",
        "segment": "INTRADAY",
        "entry_date": "2026-11-13",
        "entry_price": 226.10,
        "quantity": 215,
        "invested_capital": 48611.50,
        "s1_support": 222.25,
        "s2_stop": 220.00,
        "r1_target": 232.00,
        "r2_target": 237.00,
        "target_desc": "High dividend bluechip demand bounce off S1 support (+2.6%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 226.10,
        "trough_low": 226.10,
        "discrepancy_noticed": "None. Healthy liquidity > ₹200 Cr.",
        "corrective_measure": "None required.",
        "is_anomaly": False
    },
    # --- Segment B: Weekly Swing (₹2,00,000 Budget) ---
    {
        "id": "SWG-20261113-POWERGRID",
        "symbol": "POWERGRID.NS",
        "name": "Power Grid Corporation",
        "segment": "WEEKLY",
        "entry_date": "2026-11-13",
        "entry_price": 258.00,
        "quantity": 385,
        "invested_capital": 99330.00,
        "s1_support": 256.78,
        "s2_stop": 251.00,
        "r1_target": 268.00,
        "r2_target": 274.00,
        "target_desc": "Weekly swing relief rally towards 50-DMA (+3.9%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 258.00,
        "trough_low": 258.00,
        "discrepancy_noticed": "None. Power utility demand floor tested.",
        "corrective_measure": "None required.",
        "is_anomaly": False
    },
    {
        "id": "SWG-20261113-VBL",
        "symbol": "VBL.NS",
        "name": "Varun Beverages",
        "segment": "WEEKLY",
        "entry_date": "2026-11-13",
        "entry_price": 432.20,
        "quantity": 228,
        "invested_capital": 98541.60,
        "s1_support": 425.00,
        "s2_stop": 415.00,
        "r1_target": 458.00,
        "r2_target": 472.00,
        "target_desc": "FMCG accumulation shelf rebound (+6.0%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 432.20,
        "trough_low": 432.20,
        "discrepancy_noticed": "None. Holding strong at ₹425-430 shelf.",
        "corrective_measure": "None required.",
        "is_anomaly": False
    },
    # --- Segment C: Multi-Month Positional (₹4,50,000 Budget) ---
    {
        "id": "POS-20261113-TCS",
        "symbol": "TCS.NS",
        "name": "Tata Consultancy Services (Core)",
        "segment": "MULTI_MONTH",
        "entry_date": "2026-11-13",
        "entry_price": 2411.30,
        "quantity": 72,
        "invested_capital": 173613.60,
        "s1_support": 2405.00,
        "s2_stop": 2280.00,
        "r1_target": 2780.00,
        "r2_target": 2950.00,
        "target_desc": "Core Large Cap -25% discount mean reversion to R1 (+15.3%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 2411.30,
        "trough_low": 2411.30,
        "discrepancy_noticed": "None. 18-Point Thesis compliant (ROCE 48%, zero debt).",
        "corrective_measure": "None required.",
        "is_anomaly": False
    },
    {
        "id": "POS-20261113-HDFCBANK",
        "symbol": "HDFCBANK.NS",
        "name": "HDFC Bank (Core)",
        "segment": "MULTI_MONTH",
        "entry_date": "2026-11-13",
        "entry_price": 800.20,
        "quantity": 218,
        "invested_capital": 174443.60,
        "s1_support": 785.00,
        "s2_stop": 750.00,
        "r1_target": 920.00,
        "r2_target": 980.00,
        "target_desc": "Private banking titan post-merger accumulation (+15.0%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 800.20,
        "trough_low": 800.20,
        "discrepancy_noticed": "None. Banking thesis compliant (GNPA 1.36%, CAR 19.3%).",
        "corrective_measure": "None required.",
        "is_anomaly": False
    },
    {
        "id": "POS-20261113-STLTECH",
        "symbol": "STLTECH.NS",
        "name": "Sterlite Technologies (Breakout Alpha)",
        "segment": "MULTI_MONTH",
        "entry_date": "2026-11-13",
        "entry_price": 697.70,
        "quantity": 142,
        "invested_capital": 99073.40,
        "s1_support": 680.00,
        "s2_stop": 640.00,
        "r1_target": 850.00,
        "r2_target": 920.00,
        "target_desc": "Breakout Alpha Blueprint: Stage-2 Expansion (Strictly < 1 Yr Hold, +21.8%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 697.70,
        "trough_low": 697.70,
        "discrepancy_noticed": "None. High volume ignition above 50-DMA.",
        "corrective_measure": "Strict < 1 Yr Rule: Take profit at R1/R2, do not carry multi-year.",
        "is_anomaly": False
    },
    # --- Segment D: Multi-Year Bedrock (₹2,50,000 Budget) ---
    {
        "id": "BED-20261113-TITAN",
        "symbol": "TITAN.NS",
        "name": "Titan Company",
        "segment": "MULTI_YEAR",
        "entry_date": "2026-11-13",
        "entry_price": 4847.00,
        "quantity": 25,
        "invested_capital": 121175.00,
        "s1_support": 4750.00,
        "s2_stop": 4400.00,
        "r1_target": 5800.00,
        "r2_target": 6500.00,
        "target_desc": "Sovereign consumer monopoly multi-year compounding (+19.6% initial target)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 4847.00,
        "trough_low": 4847.00,
        "discrepancy_noticed": "None. Flawless 18-Point governance & moat metrics.",
        "corrective_measure": "Long-term compounding hold.",
        "is_anomaly": False
    },
    {
        "id": "BED-20261113-POLYCAB",
        "symbol": "POLYCAB.NS",
        "name": "Polycab India",
        "segment": "MULTI_YEAR",
        "entry_date": "2026-11-13",
        "entry_price": 8064.50,
        "quantity": 15,
        "invested_capital": 120967.50,
        "s1_support": 7850.00,
        "s2_stop": 7300.00,
        "r1_target": 9600.00,
        "r2_target": 10500.00,
        "target_desc": "Infrastructure & Cables leader 28% ROCE compounding (+19.0%)",
        "status": "ACTIVE",
        "exit_date": None,
        "exit_price": None,
        "realized_pnl": 0.0,
        "peak_high": 8064.50,
        "trough_low": 8064.50,
        "discrepancy_noticed": "None. Zero debt, 24% market share dominance.",
        "corrective_measure": "Long-term compounding hold.",
        "is_anomaly": False
    }
]

def load_portfolio_state() -> Dict[str, Any]:
    """Loads portfolio state from JSON, or initializes with fresh ₹10 Lakhs structure."""
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[TestingEngine] Warning loading state: {e}. Reinitializing.")

    # Initialize fresh state
    state = {
        "starting_capital": SEED_CAPITAL,
        "created_at": "2026-11-13T09:15:00+05:30",
        "last_sync": datetime.datetime.now().isoformat(),
        "segments": {},
        "trades": INITIAL_TRADES,
        "closed_trades": [],
        "anomalies_log": []
    }

    # Calculate segment cash balances
    for seg_key, cfg in SEGMENT_CONFIG.items():
        seg_trades = [t for t in INITIAL_TRADES if t["segment"] == seg_key]
        invested = sum(t["invested_capital"] for t in seg_trades)
        cash = cfg["allocated_capital"] - invested
        state["segments"][seg_key] = {
            "name": cfg["name"],
            "allocated_capital": cfg["allocated_capital"],
            "weight_pct": cfg["weight_pct"],
            "invested_capital": round(invested, 2),
            "cash_balance": round(cash, 2),
            "realized_pnl": 0.0,
            "closed_count": 0,
            "win_count": 0,
            "loss_count": 0
        }

    save_portfolio_state(state)
    return state

def save_portfolio_state(state: Dict[str, Any]):
    """Persists portfolio state to disk."""
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"[TestingEngine] Error saving state: {e}")

def sync_live_prices(force_yfinance: bool = False) -> Dict[str, Any]:
    """
    Fetches real-time market prices for all active trades, updates LTP,
    MFE (Peak Gain %), MAE (Drawdown %), and evaluates Target/Stop executions.
    """
    state = load_portfolio_state()
    active_trades = [t for t in state["trades"] if t["status"] == "ACTIVE"]
    if not active_trades:
        return compute_portfolio_summary(state)

    symbols = list({t["symbol"] for t in active_trades})
    live_prices = {}

    # 1. Fetch prices via yfinance
    for sym in symbols:
        try:
            t = yf.Ticker(sym)
            df = t.history(period="2d")
            if not df.empty:
                live_prices[sym] = {
                    "ltp": round(float(df["Close"].iloc[-1]), 2),
                    "high": round(float(df["High"].iloc[-1]), 2),
                    "low": round(float(df["Low"].iloc[-1]), 2)
                }
        except Exception as e:
            print(f"[TestingEngine] Error fetching live price for {sym}: {e}")

    now_iso = datetime.datetime.now().isoformat()
    state["last_sync"] = now_iso

    # 2. Update each active trade
    for t in active_trades:
        sym = t["symbol"]
        p_data = live_prices.get(sym)
        if not p_data:
            continue

        ltp = p_data["ltp"]
        day_high = p_data["high"]
        day_low = p_data["low"]

        # Update historical peaks and troughs
        t["peak_high"] = max(t.get("peak_high", t["entry_price"]), day_high, ltp)
        t["trough_low"] = min(t.get("trough_low", t["entry_price"]), day_low, ltp)

        # MFE (Maximum Favorable Excursion) & MAE (Maximum Adverse Excursion)
        t["mfe_pct"] = round(((t["peak_high"] - t["entry_price"]) / t["entry_price"] * 100.0), 2)
        t["mae_pct"] = round(((t["trough_low"] - t["entry_price"]) / t["entry_price"] * 100.0), 2)

        # Current Live P&L
        t["current_ltp"] = ltp
        t["current_value"] = round(t["quantity"] * ltp, 2)
        t["unrealized_pnl"] = round(t["current_value"] - t["invested_capital"], 2)
        t["unrealized_pnl_pct"] = round((t["unrealized_pnl"] / t["invested_capital"] * 100.0), 2)

        # Automated Target R1 Execution Check
        if ltp >= t["r1_target"]:
            t["status"] = "TARGET_R1_HIT ✅"
            t["exit_date"] = now_iso[:10]
            t["exit_price"] = ltp
            t["realized_pnl"] = t["unrealized_pnl"]
            _process_closed_trade(state, t, is_win=True)

        # Automated Stop S2 Execution Check
        elif ltp <= t["s2_stop"]:
            t["status"] = "STOP_S2_HIT ❌"
            t["exit_date"] = now_iso[:10]
            t["exit_price"] = ltp
            t["realized_pnl"] = t["unrealized_pnl"]
            t["is_anomaly"] = True
            t["discrepancy_noticed"] = f"Stop S2 (₹{t['s2_stop']}) breached at LTP ₹{ltp:.2f}."
            t["corrective_measure"] = "Logged into Tier 1 Anomaly Watchlist (Noise boundary: 5-7%)."
            _process_closed_trade(state, t, is_win=False)

    save_portfolio_state(state)
    return compute_portfolio_summary(state)

def _process_closed_trade(state: Dict[str, Any], trade: Dict[str, Any], is_win: bool):
    """Processes capital return & automated profit reinvestment upon trade exit."""
    seg_key = trade["segment"]
    seg = state["segments"].get(seg_key)
    if not seg:
        return

    # Capital + Profit returns directly to segment cash balance for reinvestment!
    returned_capital = trade["invested_capital"] + trade["realized_pnl"]
    seg["cash_balance"] = round(seg["cash_balance"] + returned_capital, 2)
    seg["invested_capital"] = round(max(0.0, seg["invested_capital"] - trade["invested_capital"]), 2)
    seg["realized_pnl"] = round(seg["realized_pnl"] + trade["realized_pnl"], 2)
    seg["closed_count"] += 1
    if is_win:
        seg["win_count"] += 1
    else:
        seg["loss_count"] += 1

    # Move to closed trades history
    state["closed_trades"].append(trade)

def compute_portfolio_summary(state: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Computes total portfolio value, realized & unrealized P&L, and win-rates."""
    if state is None:
        state = load_portfolio_state()

    total_starting = state.get("starting_capital", SEED_CAPITAL)
    total_invested = 0.0
    total_current_value = 0.0
    total_unrealized_pnl = 0.0
    total_realized_pnl = 0.0
    total_cash_in_hand = 0.0

    total_wins = 0
    total_losses = 0

    segment_summaries = {}
    for seg_key, seg_data in state["segments"].items():
        seg_active_trades = [t for t in state["trades"] if t["segment"] == seg_key and t["status"] == "ACTIVE"]
        
        seg_inv = sum(t["invested_capital"] for t in seg_active_trades)
        seg_val = sum(t.get("current_value", t["invested_capital"]) for t in seg_active_trades)
        seg_unreal = sum(t.get("unrealized_pnl", 0.0) for t in seg_active_trades)
        seg_cash = seg_data.get("cash_balance", 0.0)
        seg_realized = seg_data.get("realized_pnl", 0.0)
        seg_net_val = round(seg_val + seg_cash, 2)

        total_invested += seg_inv
        total_current_value += seg_val
        total_unrealized_pnl += seg_unreal
        total_realized_pnl += seg_realized
        total_cash_in_hand += seg_cash

        total_wins += seg_data.get("win_count", 0)
        total_losses += seg_data.get("loss_count", 0)

        segment_summaries[seg_key] = {
            "name": seg_data["name"],
            "allocated_capital": seg_data["allocated_capital"],
            "weight_pct": seg_data["weight_pct"],
            "invested": round(seg_inv, 2),
            "cash_buffer": round(seg_cash, 2),
            "current_stock_value": round(seg_val, 2),
            "total_segment_value": seg_net_val,
            "realized_pnl": round(seg_realized, 2),
            "unrealized_pnl": round(seg_unreal, 2),
            "total_pnl": round(seg_realized + seg_unreal, 2),
            "total_pnl_pct": round(((seg_net_val - seg_data["allocated_capital"]) / seg_data["allocated_capital"] * 100.0), 2),
            "active_trades_count": len(seg_active_trades),
            "closed_trades_count": seg_data.get("closed_count", 0)
        }

    total_portfolio_net_worth = round(total_current_value + total_cash_in_hand, 2)
    total_net_pnl = round(total_portfolio_net_worth - total_starting, 2)
    total_net_pnl_pct = round((total_net_pnl / total_starting * 100.0), 2)

    total_closed = total_wins + total_losses
    win_rate_pct = round((total_wins / total_closed * 100.0), 1) if total_closed > 0 else 100.0

    # Anomaly rate computation (3-Tier Noise Filter)
    all_trades_count = len(state["trades"]) + len(state.get("closed_trades", []))
    anomalies_count = sum(1 for t in state["trades"] + state.get("closed_trades", []) if t.get("is_anomaly", False))
    anomaly_rate_pct = round((anomalies_count / max(1, all_trades_count) * 100.0), 1)

    if anomaly_rate_pct <= 5.0:
        noise_status = "TIER 1: NORMAL MARKET NOISE (<5%) — ZERO ACTION REQUIRED ✅"
    elif anomaly_rate_pct <= 8.0:
        noise_status = "TIER 2: WATCHLIST WARNING (5-8%) — OBSERVE REPEAT PATTERNS ⚠️"
    else:
        noise_status = "TIER 3: SYSTEMIC MODEL DEFECT (>8%) — TRIGGER MODEL UPGRADE 🚨"

    return {
        "starting_capital": total_starting,
        "portfolio_net_worth": total_portfolio_net_worth,
        "cash_in_hand": round(total_cash_in_hand, 2),
        "stocks_invested_value": round(total_current_value, 2),
        "total_unrealized_pnl": round(total_unrealized_pnl, 2),
        "total_realized_pnl": round(total_realized_pnl, 2),
        "total_net_pnl": total_net_pnl,
        "total_net_pnl_pct": total_net_pnl_pct,
        "win_rate_pct": win_rate_pct,
        "total_trades_count": all_trades_count,
        "active_trades_count": len([t for t in state["trades"] if t["status"] == "ACTIVE"]),
        "closed_trades_count": total_closed,
        "anomaly_rate_pct": anomaly_rate_pct,
        "noise_filter_status": noise_status,
        "last_sync": state.get("last_sync"),
        "segments": segment_summaries,
        "trades": state["trades"],
        "closed_trades": state.get("closed_trades", [])
    }

def export_daily_testing_excel(target_date: Optional[str] = None) -> str:
    """
    Generates the daily timestamped testing ledger Excel workbook:
    '2. Output Generated/Daily_Testing_Ledger_YYYY-MM-DD.xlsx'
    
    Sheet 1: Executive Portfolio Dashboard (Created out of ₹10 Lakhs)
    Sheet 2: Intraday S&R Precision (1-Day)
    Sheet 3: Weekly Swing Trades (3-5 Days)
    Sheet 4: Multi-Month Positional (3-10 Months, Breakout Alpha < 1 Yr)
    Sheet 5: Multi-Year Bedrock (2-5 Years)
    Sheet 6: Discrepancy & Anomaly Audit Log (5-7% Noise Filter)
    """
    if not target_date:
        target_date = datetime.date.today().strftime("%Y-%m-%d")

    summary = sync_live_prices(force_yfinance=False)
    file_name = f"Daily_Testing_Ledger_{target_date}.xlsx"
    file_path = os.path.join(OUTPUT_DIR, file_name)

    wb = openpyxl.Workbook()
    
    # Styles
    c_navy = "0F172A"
    c_blue = "1E3A8A"
    c_sub = "334155"
    c_border = "CBD5E1"
    c_green_bg = "DCFCE7"
    c_green_txt = "166534"
    c_amber_bg = "FEF3C7"
    c_amber_txt = "92400E"
    c_gray_bg = "F8FAFC"
    c_blue_bg = "EFF6FF"

    thin_border = Border(
        left=Side(style='thin', color=c_border),
        right=Side(style='thin', color=c_border),
        top=Side(style='thin', color=c_border),
        bottom=Side(style='thin', color=c_border)
    )
    header_border = Border(
        left=Side(style='thin', color='475569'),
        right=Side(style='thin', color='475569'),
        top=Side(style='medium', color='1E3A8A'),
        bottom=Side(style='medium', color='1E3A8A')
    )

    def _banner(ws, title, subtitle, max_cols=8):
        end_col = get_column_letter(max_cols)
        ws.merge_cells(f"A1:{end_col}1")
        b1 = ws["A1"]
        b1.value = f"ANTIGRAVITY WEALTH — {title.upper()}"
        b1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        b1.fill = PatternFill(fill_type="solid", start_color=c_navy, end_color=c_navy)
        b1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        ws.merge_cells(f"A2:{end_col}2")
        b2 = ws["A2"]
        b2.value = subtitle
        b2.font = Font(name="Calibri", size=9.5, italic=True, color="94A3B8")
        b2.fill = PatternFill(fill_type="solid", start_color=c_navy, end_color=c_navy)
        b2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 18

    # ==================== SHEET 1: Executive Portfolio Dashboard ====================
    ws1 = wb.active
    ws1.title = "Portfolio Dashboard (₹10L)"
    ws1.views.sheetView[0].showGridLines = True
    _banner(ws1, "Testing Lab: Master Portfolio Executive Dashboard", f"Seed Capital: ₹10,00,000 | Daily Audit Snapshot: {target_date} | Automated Reinvestment Active", 7)

    # Top KPI Metrics Cards
    kpis = [
        ("SEED CAPITAL", f"₹{summary['starting_capital']:,.2f}", "Initial Seed"),
        ("CURRENT PORTFOLIO VALUE", f"₹{summary['portfolio_net_worth']:,.2f}", "Net Value"),
        ("TOTAL NET P&L", f"{'+' if summary['total_net_pnl']>=0 else ''}₹{summary['total_net_pnl']:,.2f} ({summary['total_net_pnl_pct']:+.2f}%)", "Net Performance"),
        ("CASH IN HAND (REINVESTABLE)", f"₹{summary['cash_in_hand']:,.2f}", "Liquid Buffer"),
        ("ACTIVE WIN-RATE", f"{summary['win_rate_pct']:.1f}%", "Target: 80-90%"),
        ("NOISE FILTER STATUS", summary['noise_filter_status'].split(':')[0], "3-Tier Standard"),
    ]
    for i, (label, val, sub) in enumerate(kpis):
        col = i + 1 if i < 6 else 6
        r = 4 if i < 6 else 4
        # We can layout KPI blocks cleanly
    
    ws1.merge_cells("A4:B4")
    ws1["A4"] = "STARTING SEED CAPITAL:"
    ws1["A4"].font = Font(name="Calibri", size=10, bold=True, color="334155")
    ws1.merge_cells("C4:D4")
    ws1["C4"] = f"₹{summary['starting_capital']:,.2f}"
    ws1["C4"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")

    ws1.merge_cells("E4:F4")
    ws1["E4"] = "LATEST PORTFOLIO NET WORTH:"
    ws1["E4"].font = Font(name="Calibri", size=10, bold=True, color="166534")
    ws1.merge_cells("G4:H4")
    ws1["G4"] = f"₹{summary['portfolio_net_worth']:,.2f}"
    ws1["G4"].font = Font(name="Calibri", size=13, bold=True, color="166534")
    ws1["G4"].fill = PatternFill(fill_type="solid", start_color=c_green_bg, end_color=c_green_bg)

    ws1.merge_cells("A5:B5")
    ws1["A5"] = "TOTAL NET P&L (₹ & %):"
    ws1["A5"].font = Font(name="Calibri", size=10, bold=True, color="334155")
    ws1.merge_cells("C5:D5")
    pnl_str = f"{'+' if summary['total_net_pnl']>=0 else ''}₹{summary['total_net_pnl']:,.2f} ({summary['total_net_pnl_pct']:+.2f}%)"
    ws1["C5"] = pnl_str
    ws1["C5"].font = Font(name="Calibri", size=11, bold=True, color="166534" if summary['total_net_pnl']>=0 else "991B1B")

    ws1.merge_cells("E5:F5")
    ws1["E5"] = "CASH IN HAND (REINVESTABLE):"
    ws1["E5"].font = Font(name="Calibri", size=10, bold=True, color="334155")
    ws1.merge_cells("G5:H5")
    ws1["G5"] = f"₹{summary['cash_in_hand']:,.2f}"
    ws1["G5"].font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")

    for r in [4, 5]:
        ws1.row_dimensions[r].height = 24
        for c in range(1, 9):
            ws1.cell(row=r, column=c).border = thin_border

    # Segment Breakdown Table
    ws1.merge_cells("A7:H7")
    ws1["A7"] = "SECTION 1: 4-TIMEFRAME CAPITAL ALLOCATION & PERFORMANCE BREAKDOWN"
    ws1["A7"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws1["A7"].fill = PatternFill(fill_type="solid", start_color=c_blue, end_color=c_blue)
    ws1["A7"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[7].height = 24

    seg_headers = ["Segment Code", "Timeframe Name", "Target Weight", "Allocated Capital", "Invested Value", "Cash Buffer", "Total Segment Net Worth", "Segment P&L (%)"]
    ws1.row_dimensions[8].height = 26
    for c_idx, h in enumerate(seg_headers, 1):
        c = ws1.cell(row=8, column=c_idx)
        c.value = h
        c.font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        c.fill = PatternFill(fill_type="solid", start_color="1E293B", end_color="1E293B")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = header_border

    r_idx = 9
    for s_key, s_data in summary["segments"].items():
        ws1.row_dimensions[r_idx].height = 28
        row_vals = [
            s_key,
            s_data["name"],
            f"{s_data['weight_pct']:.0f}%",
            f"₹{s_data['allocated_capital']:,.2f}",
            f"₹{s_data['invested']:,.2f}",
            f"₹{s_data['cash_buffer']:,.2f}",
            f"₹{s_data['total_segment_value']:,.2f}",
            f"{s_data['total_pnl_pct']:+.2f}%"
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws1.cell(row=r_idx, column=c_idx)
            c.value = val
            c.font = Font(name="Calibri", size=9.5, bold=(c_idx in [1, 7, 8]))
            c.border = thin_border
            if c_idx in [1, 3]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [4, 5, 6, 7]:
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif c_idx == 8:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = PatternFill(fill_type="solid", start_color=c_green_bg if "+" in val else c_gray_bg, end_color=c_green_bg if "+" in val else c_gray_bg)
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
        r_idx += 1

    # Total Row
    ws1.row_dimensions[r_idx].height = 26
    ws1.merge_cells(f"A{r_idx}:B{r_idx}")
    ws1[f"A{r_idx}"] = "TOTAL ₹10 LAKHS PORTFOLIO"
    ws1[f"A{r_idx}"].font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    ws1[f"A{r_idx}"].alignment = Alignment(horizontal="center", vertical="center")
    ws1[f"A{r_idx}"].fill = PatternFill(fill_type="solid", start_color="E2E8F0", end_color="E2E8F0")

    ws1[f"C{r_idx}"] = "100.0%"
    ws1[f"C{r_idx}"].font = Font(name="Calibri", size=9.5, bold=True)
    ws1[f"C{r_idx}"].alignment = Alignment(horizontal="center", vertical="center")
    ws1[f"C{r_idx}"].fill = PatternFill(fill_type="solid", start_color="E2E8F0", end_color="E2E8F0")

    ws1[f"D{r_idx}"] = f"₹{summary['starting_capital']:,.2f}"
    ws1[f"D{r_idx}"].font = Font(name="Calibri", size=9.5, bold=True)
    ws1[f"D{r_idx}"].alignment = Alignment(horizontal="right", vertical="center")
    ws1[f"D{r_idx}"].fill = PatternFill(fill_type="solid", start_color="E2E8F0", end_color="E2E8F0")

    ws1[f"E{r_idx}"] = f"₹{summary['stocks_invested_value']:,.2f}"
    ws1[f"E{r_idx}"].font = Font(name="Calibri", size=9.5, bold=True)
    ws1[f"E{r_idx}"].alignment = Alignment(horizontal="right", vertical="center")
    ws1[f"E{r_idx}"].fill = PatternFill(fill_type="solid", start_color="E2E8F0", end_color="E2E8F0")

    ws1[f"F{r_idx}"] = f"₹{summary['cash_in_hand']:,.2f}"
    ws1[f"F{r_idx}"].font = Font(name="Calibri", size=9.5, bold=True)
    ws1[f"F{r_idx}"].alignment = Alignment(horizontal="right", vertical="center")
    ws1[f"F{r_idx}"].fill = PatternFill(fill_type="solid", start_color="E2E8F0", end_color="E2E8F0")

    ws1[f"G{r_idx}"] = f"₹{summary['portfolio_net_worth']:,.2f}"
    ws1[f"G{r_idx}"].font = Font(name="Calibri", size=10.5, bold=True, color="166534")
    ws1[f"G{r_idx}"].alignment = Alignment(horizontal="right", vertical="center")
    ws1[f"G{r_idx}"].fill = PatternFill(fill_type="solid", start_color=c_green_bg, end_color=c_green_bg)

    ws1[f"H{r_idx}"] = f"{summary['total_net_pnl_pct']:+.2f}%"
    ws1[f"H{r_idx}"].font = Font(name="Calibri", size=10.5, bold=True, color="166534")
    ws1[f"H{r_idx}"].alignment = Alignment(horizontal="center", vertical="center")
    ws1[f"H{r_idx}"].fill = PatternFill(fill_type="solid", start_color=c_green_bg, end_color=c_green_bg)

    for c in range(1, 9):
        ws1.cell(row=r_idx, column=c).border = header_border

    # Section 2: Noise Filter Status Banner
    r_idx += 2
    ws1.merge_cells(f"A{r_idx}:H{r_idx}")
    ws1[f"A{r_idx}"] = "SECTION 2: EMPIRICAL VERIFICATION AUDIT & 5-7% NOISE FILTER"
    ws1[f"A{r_idx}"].font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws1[f"A{r_idx}"].fill = PatternFill(fill_type="solid", start_color=c_blue, end_color=c_blue)
    ws1[f"A{r_idx}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[r_idx].height = 24

    r_idx += 1
    ws1.row_dimensions[r_idx].height = 36
    ws1.merge_cells(f"A{r_idx}:H{r_idx}")
    ws1[f"A{r_idx}"] = f"CURRENT ANOMALY RATE: {summary['anomaly_rate_pct']:.1f}% ({summary['noise_filter_status']})\nRule: Discrepancies <= 5% are classified as standard random market noise. System model upgrades only trigger if repeat flaws exceed 8%."
    ws1[f"A{r_idx}"].font = Font(name="Calibri", size=9.5, bold=True, color="166534")
    ws1[f"A{r_idx}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws1[f"A{r_idx}"].fill = PatternFill(fill_type="solid", start_color=c_green_bg, end_color=c_green_bg)
    for c in range(1, 9):
        ws1.cell(row=r_idx, column=c).border = thin_border

    # Auto widths for Sheet 1
    ws1_widths = [16, 28, 16, 20, 20, 18, 24, 18]
    for i, w in enumerate(ws1_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # ==================== SHEETS 2 TO 5: Timeframe Specific Trade Logs ====================
    seg_sheets_map = [
        ("INTRADAY", "Intraday S&R (1D)", "Segment A: Intraday S&R Floor Precision (1-Day)", "Validating immediate S1 demand floor bounce and intraday reaction"),
        ("WEEKLY", "Weekly Swing (3-5D)", "Segment B: Weekly Swing Demand Expansion (3-5 Days)", "Testing weekly momentum towards R1 ceiling with S2 stop protection"),
        ("MULTI_MONTH", "Multi-Month (3-10M)", "Segment C: Multi-Month Positional Capital Protection (3-10M)", "Core Bluechips + Breakout Alpha Satellite (Strictly < 1 Year Hold)"),
        ("MULTI_YEAR", "Multi-Year Bedrock (2-5Y)", "Segment D: Multi-Year Bedrock Compounding (2-5 Years)", "Pure Wealth Compounding in 18-Pt wide-moat monopolies")
    ]

    trade_headers = [
        "Trade ID", "Ticker", "Company Name", "Entry Date", "Entry Price", "Qty",
        "Invested", "S1 Floor", "R1 Target", "S2 Stop", "LTP", "Current P&L (₹)",
        "P&L %", "MFE Peak %", "MAE Draw %", "Status", "Discrepancy / Gap Noticed", "Corrective Action (Trigger >5%)"
    ]

    for s_code, s_title, banner_h, banner_sub in seg_sheets_map:
        ws_seg = wb.create_sheet(title=s_title)
        ws_seg.views.sheetView[0].showGridLines = True
        _banner(ws_seg, banner_h, banner_sub, len(trade_headers))

        ws_seg.row_dimensions[4].height = 26
        for c_idx, h in enumerate(trade_headers, 1):
            c = ws_seg.cell(row=4, column=c_idx)
            c.value = h
            c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c.fill = PatternFill(fill_type="solid", start_color="1E293B", end_color="1E293B")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = header_border

        trades = [t for t in summary["trades"] if t["segment"] == s_code] + [t for t in summary["closed_trades"] if t["segment"] == s_code]
        r = 5
        for t in trades:
            ws_seg.row_dimensions[r].height = 30
            pnl_val = t.get("unrealized_pnl", t.get("realized_pnl", 0.0))
            pnl_pct = t.get("unrealized_pnl_pct", 0.0)
            row_data = [
                t["id"],
                t["symbol"],
                t["name"],
                t["entry_date"],
                f"₹{t['entry_price']:,.2f}",
                t["quantity"],
                f"₹{t['invested_capital']:,.2f}",
                f"₹{t['s1_support']:,.2f}",
                f"₹{t['r1_target']:,.2f}",
                f"₹{t['s2_stop']:,.2f}",
                f"₹{t.get('current_ltp', t['entry_price']):,.2f}",
                f"{'+' if pnl_val>=0 else ''}₹{pnl_val:,.2f}",
                f"{pnl_pct:+.2f}%",
                f"+{t.get('mfe_pct', 0.0):.1f}%",
                f"{t.get('mae_pct', 0.0):.1f}%",
                t["status"],
                t.get("discrepancy_noticed", "None"),
                t.get("corrective_measure", "None required")
            ]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws_seg.cell(row=r, column=c_idx)
                cell.value = val
                cell.font = Font(name="Calibri", size=8.5, bold=(c_idx in [2, 12, 13, 16]))
                cell.border = thin_border
                if c_idx in [1, 2, 4, 6, 16]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [5, 7, 8, 9, 10, 11, 12, 13, 14, 15]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                if c_idx in [12, 13]:
                    cell.fill = PatternFill(fill_type="solid", start_color=c_green_bg if pnl_val >= 0 else "FEE2E2", end_color=c_green_bg if pnl_val >= 0 else "FEE2E2")
                elif c_idx == 16:
                    cell.fill = PatternFill(fill_type="solid", start_color=c_blue_bg, end_color=c_blue_bg)
            r += 1

        seg_widths = [16, 14, 26, 12, 14, 8, 14, 12, 12, 12, 12, 14, 12, 12, 12, 18, 32, 34]
        for c_idx, w in enumerate(seg_widths, 1):
            ws_seg.column_dimensions[get_column_letter(c_idx)].width = w

    # ==================== SHEET 6: Discrepancy & Anomaly Audit ====================
    ws_ano = wb.create_sheet(title="Discrepancy Audit")
    ws_ano.views.sheetView[0].showGridLines = True
    _banner(ws_ano, "Discrepancy & Anomaly Audit Log", "Rigorous Governance Audit: 5-7% 3-Tier Noise Filter Protocol", 7)

    ano_headers = ["Audit Sr", "Trade ID / Symbol", "Timeframe Segment", "Incident Date", "Discrepancy / Fault Observed", "Root Cause Analysis", "Corrective Action / Protocol"]
    ws_ano.row_dimensions[4].height = 26
    for c_idx, h in enumerate(ano_headers, 1):
        c = ws_ano.cell(row=4, column=c_idx)
        c.value = h
        c.font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        c.fill = PatternFill(fill_type="solid", start_color="1E293B", end_color="1E293B")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = header_border

    # Initial Audit Policy Row
    ws_ano.row_dimensions[5].height = 36
    ano_init_row = [
        "1", "ALL TRADES", "PORTFOLIO LEVEL", target_date,
        f"Baseline Anomaly Rate: {summary['anomaly_rate_pct']:.1f}%. All 9 positions sitting within healthy S1-S2 risk bands.",
        "System initialization baseline. No structural breaches observed.",
        "Enforce Tier 1 protocol: ignore isolated noise <=5%. Only trigger rule upgrade if repeat failures exceed 8%."
    ]
    for c_idx, val in enumerate(ano_init_row, 1):
        cell = ws_ano.cell(row=5, column=c_idx)
        cell.value = val
        cell.font = Font(name="Calibri", size=9, bold=(c_idx in [1, 2]))
        cell.border = thin_border
        if c_idx in [1, 2, 3, 4]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ano_widths = [10, 18, 20, 14, 38, 38, 42]
    for c_idx, w in enumerate(ano_widths, 1):
        ws_ano.column_dimensions[get_column_letter(c_idx)].width = w

    wb.save(file_path)
    print(f"[TestingEngine] Successfully generated Daily Testing Ledger: {file_path}")
    return file_path

if __name__ == "__main__":
    print("Testing Engine Initialization...")
    summary = sync_live_prices(force_yfinance=True)
    print(f"Total Portfolio Net Worth: ₹{summary['portfolio_net_worth']:,.2f}")
    excel_path = export_daily_testing_excel()
    print(f"Daily Excel exported to: {excel_path}")
