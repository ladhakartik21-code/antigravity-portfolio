"""
AntiGravity Sovereign Wealth Engine - Autonomous Intraday S&R Execution Engine
Automated paper trading engine testing Support & Resistance bounces with ₹1,00,000 cash balance.
"""

import os
import json
import datetime
import time
import requests
from typing import Dict, Any, List

APP_DIR = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = os.path.join(APP_DIR, "intraday_trades.json")

# High-Conviction Multidisciplinary Intraday Universe
INTRADAY_UNIVERSE = {
    "POLYCAB.NS": {
        "name": "Polycab India",
        "sector": "Cables & Wires (28% ROCE)",
        "s1": 7680.0,
        "s2": 7610.0,
        "r1": 7840.0,
        "r2": 7920.0,
        "target_pct": 2.1,
        "sl_pct": 1.1,
        "thesis": "Testing Institutional Demand Zone (₹7,610 - ₹7,680) with 15-min volume absorption. Target R1 Supply shelf."
    },
    "VBL.NS": {
        "name": "Varun Beverages",
        "sector": "FMCG Moat (27% ROCE)",
        "s1": 412.0,
        "s2": 407.0,
        "r1": 424.0,
        "r2": 431.0,
        "target_pct": 2.9,
        "sl_pct": 1.2,
        "thesis": "Double-bottom at Monthly S1 Demand Zone (₹407 - ₹412). Intraday RSI oversold reversal."
    },
    "CDSL.NS": {
        "name": "Central Depository",
        "sector": "Demat Monopoly (Zero Debt)",
        "s1": 1210.0,
        "s2": 1195.0,
        "r1": 1255.0,
        "r2": 1278.0,
        "target_pct": 3.7,
        "sl_pct": 1.2,
        "thesis": "Monopoly duopoly resting on institutional liquidity shelf ₹1,195 - ₹1,210. Quick mean reversion to 20-EMA."
    },
    "TATAPOWER.NS": {
        "name": "Tata Power",
        "sector": "Energy Transition",
        "s1": 337.5,
        "s2": 333.0,
        "r1": 347.0,
        "r2": 354.0,
        "target_pct": 2.8,
        "sl_pct": 1.3,
        "thesis": "Tested lower channel demand at ₹333 - ₹337.5. Solar rooftop order execution supports high R:R intraday swing."
    },
    "HDFCBANK.NS": {
        "name": "HDFC Bank",
        "sector": "Private Banking Bluechip",
        "s1": 652.0,
        "s2": 646.0,
        "r1": 668.0,
        "r2": 675.0,
        "target_pct": 2.4,
        "sl_pct": 0.9,
        "thesis": "Heavyweight banking pivot at Monthly S2 (₹646 - ₹652). Institutional accumulation on 5-min delivery bar."
    },
    "TATSILV.NS": {
        "name": "Tata Silver ETF",
        "sector": "Precious & Solar Metals",
        "s1": 18.35,
        "s2": 18.15,
        "r1": 19.10,
        "r2": 19.45,
        "target_pct": 4.1,
        "sl_pct": 1.1,
        "thesis": "Bloomberg GSR 73.7x mean-reversion with global spot finding $54 floor. ETF demand shelf accumulation."
    }
}

def get_ist_now() -> datetime.datetime:
    """Returns current datetime in Indian Standard Time (UTC+5:30)"""
    utc = datetime.datetime.now(datetime.timezone.utc)
    ist = utc + datetime.timedelta(hours=5, minutes=30)
    return ist

def is_market_hours() -> bool:
    """Checks if current IST time is during NSE market hours (Mon-Fri 09:15 to 15:30)"""
    ist = get_ist_now()
    if ist.weekday() >= 5: # Saturday or Sunday
        return False
    t = ist.time()
    market_open = datetime.time(9, 15)
    market_close = datetime.time(15, 30)
    return market_open <= t <= market_close

def is_square_off_time() -> bool:
    """Checks if current IST time is at or past 15:15 EOD square-off time"""
    ist = get_ist_now()
    return is_market_hours() and ist.time() >= datetime.time(15, 15)

def get_initial_state() -> Dict[str, Any]:
    ist = get_ist_now()
    return {
        "account": {
            "initial_capital": 100000.0,
            "available_cash": 100000.0,
            "deployed_capital": 0.0,
            "realized_pnl": 0.0,
            "unrealized_pnl": 0.0,
            "total_value": 100000.0,
            "day_return_pct": 0.0
        },
        "config": {
            "auto_trading_enabled": True,
            "max_concurrent_positions": 4,
            "allocation_per_trade": 22000.0,
            "max_risk_pct": 1.2,
            "eod_square_off_time": "15:15 IST",
            "benchmark": "S&R Confluence Model"
        },
        "market_status": {
            "is_open": is_market_hours(),
            "current_time_ist": ist.strftime("%Y-%m-%d %H:%M:%S IST"),
            "next_bell": "09:15 AM IST (Next Session)",
            "message": "Market Open - Live Tick Monitoring" if is_market_hours() else "Market Closed - Engine Armed for 09:15 AM Bell"
        },
        "open_positions": [],
        "closed_trades": [],
        "performance_stats": {
            "total_trades": 0,
            "winning_trades": 0,
            "losing_trades": 0,
            "win_rate_pct": 0.0,
            "gross_profit": 0.0,
            "gross_loss": 0.0,
            "profit_factor": 0.0
        }
    }

def load_state() -> Dict[str, Any]:
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data
        except Exception:
            pass
    state = get_initial_state()
    save_state(state)
    return state

def save_state(state: Dict[str, Any]):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception as e:
        print(f"Error saving intraday state: {e}")

def fetch_live_ltp(symbol: str, fallback_price: float) -> float:
    """Fetches real-time price from yfinance or fallback"""
    try:
        import yfinance as yf
        ticker = yf.Ticker(symbol)
        data = ticker.history(period="1d", interval="1m")
        if not data.empty and len(data) > 0:
            return round(float(data["Close"].iloc[-1]), 2)
    except Exception:
        pass
    return round(float(fallback_price), 2)

def evaluate_intraday_engine(force_simulation: bool = False) -> Dict[str, Any]:
    """
    Core Intraday Execution Loop:
    1. Updates live LTP & PnL for open positions.
    2. Tests if Target or Stop-Loss is hit -> Closes trade.
    3. If within market hours, scans universe for S&R Demand Zone Touches -> Enters BUY.
    4. Auto squares off at 15:15 IST.
    """
    state = load_state()
    ist = get_ist_now()
    market_active = is_market_hours() or force_simulation
    state["market_status"]["is_open"] = is_market_hours()
    state["market_status"]["current_time_ist"] = ist.strftime("%Y-%m-%d %H:%M:%S IST")
    state["market_status"]["message"] = "Market Open - Live S&R Auto-Trading Active" if is_market_hours() else "Market Closed - Engine Armed for 09:15 AM Bell"

    open_pos = state.get("open_positions", [])
    closed = state.get("closed_trades", [])
    account = state["account"]
    config = state["config"]

    # 1. Update Open Positions & Check Target / Stop Loss
    remaining_open = []
    for pos in open_pos:
        sym = pos["symbol"]
        ltp = fetch_live_ltp(sym, pos["current_price"])
        pos["current_price"] = ltp
        pnl = round((ltp - pos["entry_price"]) * pos["quantity"], 2)
        pnl_pct = round((ltp - pos["entry_price"]) / pos["entry_price"] * 100, 2)
        pos["unrealized_pnl"] = pnl
        pos["unrealized_pnl_pct"] = pnl_pct

        is_target_hit = ltp >= pos["target_price"]
        is_sl_hit = ltp <= pos["stop_loss"]
        is_eod_square_off = is_square_off_time()

        if is_target_hit or is_sl_hit or is_eod_square_off:
            # Close Position
            exit_reason = "TARGET HIT (WIN)" if is_target_hit else ("STOP LOSS HIT (CUT)" if is_sl_hit else "EOD 15:15 AUTO SQUARE-OFF")
            closed_trade = {
                "trade_id": pos["trade_id"],
                "symbol": sym,
                "name": pos["name"],
                "sector": pos["sector"],
                "entry_time": pos["entry_time"],
                "exit_time": ist.strftime("%H:%M:%S"),
                "entry_price": pos["entry_price"],
                "exit_price": ltp,
                "quantity": pos["quantity"],
                "capital_deployed": pos["capital_deployed"],
                "target_price": pos["target_price"],
                "stop_loss": pos["stop_loss"],
                "realized_pnl": pnl,
                "realized_pnl_pct": pnl_pct,
                "status": "CLOSED",
                "outcome": exit_reason,
                "thesis_vindication": f"Level respected: Rebounded from Demand Zone ({pos['entry_price']}) to Target ({pos['target_price']})" if is_target_hit else "Strict Risk Defense: Exited immediately at Stop Loss to preserve capital."
            }
            closed.append(closed_trade)
            account["available_cash"] += pos["capital_deployed"] + pnl
            account["realized_pnl"] += pnl
        else:
            remaining_open.append(pos)

    state["open_positions"] = remaining_open
    state["closed_trades"] = closed

    # 2. Check for New Entries if Auto-Trading Enabled
    if config.get("auto_trading_enabled", True) and market_active:
        current_symbols = [p["symbol"] for p in state["open_positions"]]
        max_pos = config.get("max_concurrent_positions", 4)
        alloc_per_trade = config.get("allocation_per_trade", 22000.0)

        if len(state["open_positions"]) < max_pos and account["available_cash"] >= 15000:
            for sym, meta in INTRADAY_UNIVERSE.items():
                if sym in current_symbols:
                    continue
                if len(state["open_positions"]) >= max_pos:
                    break

                fallback_px = (meta["s1"] + meta["s2"]) / 2.0
                ltp = fetch_live_ltp(sym, fallback_px)
                
                # Demand Zone Test Condition: Price is resting between S2 and slightly above S1 (within 0.8%)
                in_demand_zone = (ltp >= meta["s2"] * 0.995) and (ltp <= meta["s1"] * 1.008)

                if in_demand_zone or force_simulation:
                    # Trigger Intraday BUY
                    buy_price = ltp
                    qty = int(alloc_per_trade / buy_price)
                    if qty <= 0:
                        qty = 1
                    actual_deployed = round(qty * buy_price, 2)

                    if account["available_cash"] >= actual_deployed:
                        account["available_cash"] -= actual_deployed
                        target_px = meta["r1"]
                        stop_loss = round(meta["s2"] * (1.0 - (meta["sl_pct"] / 100.0)), 2)
                        reward = target_px - buy_price
                        risk = max(buy_price - stop_loss, 0.1)
                        rr_ratio = f"1 : {round(reward / risk, 2)}"

                        new_trade = {
                            "trade_id": f"ITD-{int(time.time())}-{sym.split('.')[0]}",
                            "symbol": sym,
                            "name": meta["name"],
                            "sector": meta["sector"],
                            "entry_time": ist.strftime("%H:%M:%S"),
                            "entry_price": buy_price,
                            "current_price": buy_price,
                            "quantity": qty,
                            "capital_deployed": actual_deployed,
                            "target_price": target_px,
                            "stop_loss": stop_loss,
                            "upside_target_pct": meta["target_pct"],
                            "downside_risk_pct": meta["sl_pct"],
                            "risk_reward_ratio": rr_ratio,
                            "unrealized_pnl": 0.0,
                            "unrealized_pnl_pct": 0.0,
                            "thesis_reason": meta["thesis"],
                            "zone_confluence": f"Monthly S1/S2 Demand Zone (₹{meta['s2']} - ₹{meta['s1']})"
                        }
                        state["open_positions"].append(new_trade)
                        current_symbols.append(sym)

    # 3. Recalculate Totals
    total_deployed = sum(p["capital_deployed"] for p in state["open_positions"])
    total_unrealized = round(sum(p["unrealized_pnl"] for p in state["open_positions"]), 2)
    account["deployed_capital"] = total_deployed
    account["unrealized_pnl"] = total_unrealized
    account["total_value"] = round(account["available_cash"] + total_deployed + total_unrealized, 2)
    net_pnl = round(account["realized_pnl"] + total_unrealized, 2)
    account["day_return_pct"] = round(net_pnl / account["initial_capital"] * 100, 2)

    # 4. Update Performance Stats
    total_closed = len(state["closed_trades"])
    wins = len([t for t in state["closed_trades"] if t["realized_pnl"] > 0])
    losses = len([t for t in state["closed_trades"] if t["realized_pnl"] <= 0])
    win_rate = round((wins / total_closed * 100), 1) if total_closed > 0 else 0.0
    gross_p = sum(t["realized_pnl"] for t in state["closed_trades"] if t["realized_pnl"] > 0)
    gross_l = abs(sum(t["realized_pnl"] for t in state["closed_trades"] if t["realized_pnl"] < 0))
    pf = round(gross_p / gross_l, 2) if gross_l > 0 else (round(gross_p, 2) if gross_p > 0 else 0.0)

    state["performance_stats"] = {
        "total_trades": total_closed,
        "winning_trades": wins,
        "losing_trades": losses,
        "win_rate_pct": win_rate,
        "gross_profit": round(gross_p, 2),
        "gross_loss": round(gross_l, 2),
        "profit_factor": pf
    }

    save_state(state)
    return state

def square_off_all_positions() -> Dict[str, Any]:
    """Emergency square-off of all active intraday positions"""
    state = load_state()
    ist = get_ist_now()
    account = state["account"]
    closed = state.get("closed_trades", [])

    for pos in state.get("open_positions", []):
        sym = pos["symbol"]
        ltp = fetch_live_ltp(sym, pos["current_price"])
        pnl = round((ltp - pos["entry_price"]) * pos["quantity"], 2)
        pnl_pct = round((ltp - pos["entry_price"]) / pos["entry_price"] * 100, 2)

        closed_trade = {
            "trade_id": pos["trade_id"],
            "symbol": sym,
            "name": pos["name"],
            "sector": pos["sector"],
            "entry_time": pos["entry_time"],
            "exit_time": ist.strftime("%H:%M:%S"),
            "entry_price": pos["entry_price"],
            "exit_price": ltp,
            "quantity": pos["quantity"],
            "capital_deployed": pos["capital_deployed"],
            "target_price": pos["target_price"],
            "stop_loss": pos["stop_loss"],
            "realized_pnl": pnl,
            "realized_pnl_pct": pnl_pct,
            "status": "CLOSED",
            "outcome": "MANUAL MOBILE SQUARE-OFF",
            "thesis_vindication": "Manual override triggered from mobile interface."
        }
        closed.append(closed_trade)
        account["available_cash"] += pos["capital_deployed"] + pnl
        account["realized_pnl"] += pnl

    state["open_positions"] = []
    state["closed_trades"] = closed
    account["deployed_capital"] = 0.0
    account["unrealized_pnl"] = 0.0
    account["total_value"] = round(account["available_cash"], 2)
    account["day_return_pct"] = round(account["realized_pnl"] / account["initial_capital"] * 100, 2)
    save_state(state)
    return state

def toggle_auto_pilot() -> bool:
    """Toggles autonomous trading flag"""
    state = load_state()
    current = state["config"].get("auto_trading_enabled", True)
    state["config"]["auto_trading_enabled"] = not current
    save_state(state)
    return state["config"]["auto_trading_enabled"]

def reset_account() -> Dict[str, Any]:
    """Resets intraday account back to fresh ₹1,00,000"""
    state = get_initial_state()
    save_state(state)
    return state
