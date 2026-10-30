"""
Indices Engine: Live Real-Time Indian & Global Benchmark Monitor
Provides high-frequency, thread-pooled live price ingestion for:
- 16 Major Indian Indices (4x4 Grid)
- 16 Major Global Indices (4x4 Grid)
With in-memory caching (60s TTL), thread-pool fast_info fetching, and pre-seeded resilient fallbacks.
"""

import datetime
import os
import time
from concurrent.futures import ThreadPoolExecutor
import yfinance as yf

# Indian Standard Time Helper
def get_ist_now():
    utc_now = datetime.timezone.utc
    ist_tz = datetime.timezone(datetime.timedelta(hours=5, minutes=30))
    return datetime.datetime.now(ist_tz)

INDIAN_INDEX_DEFINITIONS = [
    {
        "id": "nifty_50",
        "name": "Nifty 50",
        "symbol": "^NSEI",
        "category": "Benchmark",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 23053.10,
        "fallback_prev": 22936.75,
        "fallback_high": 23118.10,
        "fallback_low": 22905.10
    },
    {
        "id": "bse_sensex",
        "name": "BSE Sensex",
        "symbol": "^BSESN",
        "category": "Benchmark",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 74459.53,
        "fallback_prev": 74151.04,
        "fallback_high": 74742.59,
        "fallback_low": 74082.03
    },
    {
        "id": "bank_nifty",
        "name": "Bank Nifty",
        "symbol": "^NSEBANK",
        "category": "Banking",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 54686.95,
        "fallback_prev": 54909.65,
        "fallback_high": 55068.50,
        "fallback_low": 54580.75
    },
    {
        "id": "nifty_it",
        "name": "Nifty IT",
        "symbol": "^CNXIT",
        "category": "Technology",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 32606.90,
        "fallback_prev": 31866.40,
        "fallback_high": 32705.45,
        "fallback_low": 31977.55
    },
    {
        "id": "nifty_financial",
        "name": "Nifty Financial",
        "symbol": "NIFTY_FIN_SERVICE.NS",
        "category": "Financial Services",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 24712.20,
        "fallback_prev": 24704.55,
        "fallback_high": 24840.90,
        "fallback_low": 24606.40
    },
    {
        "id": "nifty_auto",
        "name": "Nifty Auto",
        "symbol": "^CNXAUTO",
        "category": "Automobile",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 25422.70,
        "fallback_prev": 25645.60,
        "fallback_high": 25685.70,
        "fallback_low": 25422.70
    },
    {
        "id": "nifty_pharma",
        "name": "Nifty Pharma",
        "symbol": "^CNXPHARMA",
        "category": "Healthcare",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 27124.85,
        "fallback_prev": 26706.95,
        "fallback_high": 27245.10,
        "fallback_low": 26680.30
    },
    {
        "id": "nifty_fmcg",
        "name": "Nifty FMCG",
        "symbol": "^CNXFMCG",
        "category": "Consumer Goods",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 43609.20,
        "fallback_prev": 43882.70,
        "fallback_high": 44102.30,
        "fallback_low": 43510.80
    },
    {
        "id": "nifty_metal",
        "name": "Nifty Metal",
        "symbol": "^CNXMETAL",
        "category": "Metals & Mining",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 11784.90,
        "fallback_prev": 11849.25,
        "fallback_high": 11950.40,
        "fallback_low": 11720.10
    },
    {
        "id": "nifty_energy",
        "name": "Nifty Energy",
        "symbol": "^CNXENERGY",
        "category": "Power & Energy",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 34957.90,
        "fallback_prev": 35300.10,
        "fallback_high": 35450.00,
        "fallback_low": 34820.50
    },
    {
        "id": "nifty_infra",
        "name": "Nifty Infra",
        "symbol": "^CNXINFRA",
        "category": "Infrastructure",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 8799.00,
        "fallback_prev": 8792.50,
        "fallback_high": 8845.20,
        "fallback_low": 8750.60
    },
    {
        "id": "nifty_realty",
        "name": "Nifty Realty",
        "symbol": "^CNXREALTY",
        "category": "Real Estate",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 813.90,
        "fallback_prev": 816.50,
        "fallback_high": 822.40,
        "fallback_low": 809.10
    },
    {
        "id": "nifty_midcap",
        "name": "Nifty Midcap 100",
        "symbol": "NIFTY_MIDCAP_100.NS",
        "category": "Broader Midcap",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 57301.25,
        "fallback_prev": 57836.80,
        "fallback_high": 58020.10,
        "fallback_low": 57190.40
    },
    {
        "id": "nifty_smallcap",
        "name": "Nifty Smallcap 100",
        "symbol": "^CNXSC",
        "category": "Broader Smallcap",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 18463.60,
        "fallback_prev": 18807.30,
        "fallback_high": 18910.50,
        "fallback_low": 18410.20
    },
    {
        "id": "nifty_next_50",
        "name": "Nifty Next 50",
        "symbol": "^NSMIDCP",
        "category": "Large-Mid Alpha",
        "flag": "🇮🇳",
        "currency": "INR",
        "fallback_price": 67709.10,
        "fallback_prev": 68113.35,
        "fallback_high": 68450.00,
        "fallback_low": 67500.00
    },
    {
        "id": "india_vix",
        "name": "India VIX",
        "symbol": "^INDIAVIX",
        "category": "Market Volatility",
        "flag": "🇮🇳",
        "currency": "Points",
        "fallback_price": 15.30,
        "fallback_prev": 15.12,
        "fallback_high": 15.85,
        "fallback_low": 14.90
    }
]

GLOBAL_INDEX_DEFINITIONS = [
    {
        "id": "sp_500",
        "name": "S&P 500",
        "symbol": "^GSPC",
        "category": "US Benchmark",
        "flag": "🇺🇸",
        "currency": "USD",
        "fallback_price": 7594.89,
        "fallback_prev": 7721.02,
        "fallback_high": 7650.00,
        "fallback_low": 7550.00
    },
    {
        "id": "nasdaq_100",
        "name": "Nasdaq 100",
        "symbol": "^NDX",
        "category": "US Tech Leaders",
        "flag": "🇺🇸",
        "currency": "USD",
        "fallback_price": 28440.15,
        "fallback_prev": 29407.05,
        "fallback_high": 28900.00,
        "fallback_low": 28300.00
    },
    {
        "id": "dow_jones",
        "name": "Dow Jones",
        "symbol": "^DJI",
        "category": "US Bluechips",
        "flag": "🇺🇸",
        "currency": "USD",
        "fallback_price": 51959.39,
        "fallback_prev": 52343.34,
        "fallback_high": 52400.00,
        "fallback_low": 51800.00
    },
    {
        "id": "russell_2000",
        "name": "Russell 2000",
        "symbol": "^RUT",
        "category": "US Small Cap",
        "flag": "🇺🇸",
        "currency": "USD",
        "fallback_price": 2807.54,
        "fallback_prev": 2858.98,
        "fallback_high": 2865.00,
        "fallback_low": 2795.00
    },
    {
        "id": "nikkei_225",
        "name": "Nikkei 225",
        "symbol": "^N225",
        "category": "Japan Benchmark",
        "flag": "🇯🇵",
        "currency": "JPY",
        "fallback_price": 63290.11,
        "fallback_prev": 64580.64,
        "fallback_high": 64200.00,
        "fallback_low": 63100.00
    },
    {
        "id": "hang_seng",
        "name": "Hang Seng",
        "symbol": "^HSI",
        "category": "Hong Kong",
        "flag": "🇭🇰",
        "currency": "HKD",
        "fallback_price": 23294.42,
        "fallback_prev": 22958.65,
        "fallback_high": 23450.00,
        "fallback_low": 23050.00
    },
    {
        "id": "shanghai_comp",
        "name": "Shanghai Comp",
        "symbol": "000001.SS",
        "category": "China Mainland",
        "flag": "🇨🇳",
        "currency": "CNY",
        "fallback_price": 3599.02,
        "fallback_prev": 3645.44,
        "fallback_high": 3630.00,
        "fallback_low": 3585.00
    },
    {
        "id": "kospi",
        "name": "KOSPI",
        "symbol": "^KS11",
        "category": "South Korea",
        "flag": "🇰🇷",
        "currency": "KRW",
        "fallback_price": 5730.74,
        "fallback_prev": 5943.27,
        "fallback_high": 5850.00,
        "fallback_low": 5710.00
    },
    {
        "id": "taiwan_weighted",
        "name": "Taiwan Weighted",
        "symbol": "^TWII",
        "category": "Taiwan Tech/Semis",
        "flag": "🇹🇼",
        "currency": "TWD",
        "fallback_price": 45211.83,
        "fallback_prev": 46326.57,
        "fallback_high": 46100.00,
        "fallback_low": 45050.00
    },
    {
        "id": "ftse_100",
        "name": "FTSE 100",
        "symbol": "^FTSE",
        "category": "United Kingdom",
        "flag": "🇬🇧",
        "currency": "GBP",
        "fallback_price": 10720.47,
        "fallback_prev": 10725.66,
        "fallback_high": 10760.00,
        "fallback_low": 10690.00
    },
    {
        "id": "dax_40",
        "name": "DAX 40",
        "symbol": "^GDAXI",
        "category": "Germany",
        "flag": "🇩🇪",
        "currency": "EUR",
        "fallback_price": 25785.43,
        "fallback_prev": 26052.73,
        "fallback_high": 26100.00,
        "fallback_low": 25700.00
    },
    {
        "id": "cac_40",
        "name": "CAC 40",
        "symbol": "^FCHI",
        "category": "France Benchmark",
        "flag": "🇫🇷",
        "currency": "EUR",
        "fallback_price": 7953.65,
        "fallback_prev": 7975.82,
        "fallback_high": 8010.00,
        "fallback_low": 7930.00
    },
    {
        "id": "euro_stoxx_50",
        "name": "Euro Stoxx 50",
        "symbol": "^STOXX50E",
        "category": "Eurozone Top 50",
        "flag": "🇪🇺",
        "currency": "EUR",
        "fallback_price": 6124.63,
        "fallback_prev": 6207.33,
        "fallback_high": 6215.00,
        "fallback_low": 6110.00
    },
    {
        "id": "dxy",
        "name": "Dollar Index (DXY)",
        "symbol": "DX-Y.NYB",
        "category": "Greenback Gauge",
        "flag": "💵",
        "currency": "USD",
        "fallback_price": 101.84,
        "fallback_prev": 101.77,
        "fallback_high": 102.15,
        "fallback_low": 101.60
    },
    {
        "id": "cboe_vix",
        "name": "CBOE VIX",
        "symbol": "^VIX",
        "category": "Global Fear Gauge",
        "flag": "📉",
        "currency": "Points",
        "fallback_price": 20.22,
        "fallback_prev": 17.21,
        "fallback_high": 21.50,
        "fallback_low": 17.10
    },
    {
        "id": "brent_crude",
        "name": "Brent Crude",
        "symbol": "BZ=F",
        "category": "Energy Benchmark",
        "flag": "🛢️",
        "currency": "USD/bbl",
        "fallback_price": 113.89,
        "fallback_prev": 112.07,
        "fallback_high": 114.50,
        "fallback_low": 111.80
    }
]

# Global In-Memory Cache
_INDICES_CACHE = {
    "timestamp": None,
    "last_fetched_epoch": 0,
    "data": None
}

def _fetch_single_index(definition):
    """Fetches real-time fast_info quote for a single index symbol with graceful fallback."""
    item = dict(definition)
    sym = item["symbol"]
    try:
        t = yf.Ticker(sym)
        fi = t.fast_info
        p = float(fi["lastPrice"])
        prev = float(fi.get("previousClose", p) or p)
        chg = round(p - prev, 2)
        pchg = round((chg / prev) * 100.0, 2) if prev else 0.0
        high = float(fi.get("dayHigh", p) or p)
        low = float(fi.get("dayLow", p) or p)
        yr_high = float(fi.get("yearHigh", p) or p)
        yr_low = float(fi.get("yearLow", p) or p)

        item["price"] = round(p, 2)
        item["prev_close"] = round(prev, 2)
        item["change"] = chg
        item["pChange"] = pchg
        item["dayHigh"] = round(high, 2)
        item["dayLow"] = round(low, 2)
        item["yearHigh"] = round(yr_high, 2)
        item["yearLow"] = round(yr_low, 2)
        item["status"] = "LIVE"
    except Exception as e:
        # Resilient fallback to pre-seeded or last-known values
        fb_p = item.get("fallback_price", 0.0)
        fb_prev = item.get("fallback_prev", fb_p)
        chg = round(fb_p - fb_prev, 2)
        pchg = round((chg / fb_prev) * 100.0, 2) if fb_prev else 0.0
        item["price"] = fb_p
        item["prev_close"] = fb_prev
        item["change"] = chg
        item["pChange"] = pchg
        item["dayHigh"] = item.get("fallback_high", fb_p)
        item["dayLow"] = item.get("fallback_low", fb_p)
        item["yearHigh"] = fb_p * 1.15
        item["yearLow"] = fb_p * 0.85
        item["status"] = "FALLBACK"
        item["fetch_error"] = str(e)
    
    return item

def get_all_indices(force_refresh: bool = False, max_cache_age_seconds: int = 45):
    """
    Returns live quotes for 16 Indian & 16 Global Indices.
    Serves from memory cache if fresh, otherwise runs multithreaded refresh.
    """
    global _INDICES_CACHE
    now_epoch = time.time()

    if not force_refresh and _INDICES_CACHE["data"] is not None:
        age = now_epoch - _INDICES_CACHE["last_fetched_epoch"]
        if age < max_cache_age_seconds:
            return _INDICES_CACHE["data"]

    all_defs = INDIAN_INDEX_DEFINITIONS + GLOBAL_INDEX_DEFINITIONS
    
    try:
        with ThreadPoolExecutor(max_workers=16) as pool:
            results = list(pool.map(_fetch_single_index, all_defs))
        
        indian_results = results[:len(INDIAN_INDEX_DEFINITIONS)]
        global_results = results[len(INDIAN_INDEX_DEFINITIONS):]
    except Exception as e:
        # In case of catastrophic thread pool failure, build from fallbacks
        indian_results = [_fetch_single_index(d) for d in INDIAN_INDEX_DEFINITIONS]
        global_results = [_fetch_single_index(d) for d in GLOBAL_INDEX_DEFINITIONS]

    ist_now = get_ist_now()
    market_active = (ist_now.weekday() < 5) and (
        (ist_now.hour > 9 or (ist_now.hour == 9 and ist_now.minute >= 15)) and
        (ist_now.hour < 15 or (ist_now.hour == 15 and ist_now.minute <= 30))
    )

    data = {
        "timestamp": ist_now.strftime("%Y-%m-%d %H:%M:%S IST"),
        "indian_market_active": market_active,
        "counts": {
            "indian": len(indian_results),
            "global": len(global_results),
            "total": len(results) if 'results' in locals() else 32
        },
        "indian": indian_results,
        "global": global_results
    }

    _INDICES_CACHE["data"] = data
    _INDICES_CACHE["last_fetched_epoch"] = now_epoch
    _INDICES_CACHE["timestamp"] = ist_now.strftime("%Y-%m-%d %H:%M:%S IST")

    return data
