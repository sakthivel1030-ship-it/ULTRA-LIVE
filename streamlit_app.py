"""
OI ULTRA LIVE - Streamlit (mobile-friendly) version.

WHY THIS IS A SEPARATE FILE FROM OI_ULTRA_LIVE.py:
  The Excel/COM version needs a real, running, licensed copy of Microsoft
  Excel and Windows - neither of which exists on a phone. This version
  keeps the exact same data pipeline (RVOL gate -> HV bands -> ITM/OTM
  aggregates -> CE/PE Wall -> Max-OI-Change strikes -> SuperTrend/Vol
  Delta/Vol Surge -> Net GEX/Zero-Gamma Flip -> Squeeze Setup flag) but
  renders it as a live web page with Streamlit instead of writing to
  Excel - so you can open it in any phone browser.

HOW TO RUN IT:
  1. pip install streamlit requests pandas streamlit-autorefresh
     (streamlit-autorefresh is optional - the app works without it, just
     with a manual "Refresh now" button instead of a real auto-timer)
  2. Get today's Upstox access token the same way you already do for the
     other scripts (upstox_login.py) - you'll paste it into the sidebar
     when the app opens, or set it once via Streamlit secrets (see
     ACCESS TOKEN section below).
  3. streamlit run streamlit_app.py
     -> opens in your desktop browser at http://localhost:8501

HOW TO USE IT ON YOUR PHONE:
  Streamlit only serves a local address by itself, so "on your phone"
  means one of:
    a) Same Wi-Fi as your computer: run the command above, then find
       your computer's local IP (Windows: `ipconfig`, look for IPv4
       Address) and open http://<that-ip>:8501 in your phone's browser.
    b) From anywhere: deploy it to Streamlit Community Cloud (free) -
       push this file + requirements.txt to a GitHub repo, connect it at
       share.streamlit.io, and you get a permanent https:// URL you can
       open on your phone from any network. This is the way to go if
       you want it usable outside your home Wi-Fi.
  Either way, the page itself is responsive - tables scroll horizontally,
  and the "Squeeze Setups" section renders as stacked cards specifically
  so it's readable on a narrow phone screen without side-scrolling.

ACCESS TOKEN:
  Upstox access tokens are daily and Upstox's login flow needs a browser
  redirect, so this app does NOT do the OAuth dance itself - paste
  today's token (from your existing upstox_login.py) into the sidebar
  each day, or set it once as a Streamlit secret so you don't have to:
    .streamlit/secrets.toml  ->  UPSTOX_ACCESS_TOKEN = "eyJ0eXAi..."
  The sidebar field is pre-filled from that secret if present, and
  always overridable per-session.

WHAT'S DIFFERENT FROM THE EXCEL VERSION (besides no Excel):
  - No cross-process file-based rate limiter (msvcrt is Windows-only and
    unnecessary here) - a simple in-memory, thread-safe limiter is used
    instead, since Streamlit serves all sessions from one process.
  - No growing/resorted history sheet - each refresh shows the CURRENT
    qualifying stocks only (a live snapshot), since a phone screen isn't
    the place to scroll a whole day's accumulated history. If you want
    a permanent record, the EOD script (or a future CSV-logging add-on)
    is still the right tool for that.
  - SuperTrend/Vol Delta/Vol Surge are cached in Streamlit's
    session_state (refreshed every SUPERTREND_REFRESH_SEC, same as the
    Excel version) - this resets if you close the browser tab/session,
    which is fine since it's just a performance cache.
"""

import os
import time
import json
import gzip
import io
import math
import statistics
import threading
from datetime import datetime, date, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd
import streamlit as st

try:
    from streamlit_autorefresh import st_autorefresh
    HAS_AUTOREFRESH = True
except ImportError:
    HAS_AUTOREFRESH = False


# ---------------------------------------------------------------------------
# Settings - same defaults as OI_ULTRA_LIVE.py; the sidebar lets you
# override the ones worth tuning per-session without editing code.
# ---------------------------------------------------------------------------
RVOL_LOOKBACK_DAYS = 10
OI_CHANGE_MODE = "increase"

HV_THRESHOLD_PCT = 2.0
MIDVOL_HV_MIN = 1.5
HV_LOOKBACK_TRADING_DAYS = 20
HV_CALENDAR_BUFFER_DAYS = 40

SUPERTREND_CANDLE_UNIT = "minutes"
SUPERTREND_CANDLE_INTERVAL = "3"
SUPERTREND_ATR_PERIOD = 10
SUPERTREND_MULTIPLIER = 3.0
SUPERTREND_HISTORY_DAYS = 3
SUPERTREND_REFRESH_SEC = 180

VOL_SURGE_LOOKBACK = 10
VOL_SURGE_THRESHOLD = 2.0

GEX_RISK_FREE_RATE = 0.065
GEX_MIN_TIME_YEARS = 1 / 365
GEX_DEFAULT_LOT_SIZE = 1
GEX_DISPLAY_DIVISOR = 1e7
SQUEEZE_APPROACH_PCT_MAX = 3.0

FALLBACK_FNO_STOCKS = [
    "RELIANCE", "TCS", "HDFCBANK", "ICICIBANK", "INFY", "SBIN", "AXISBANK",
    "KOTAKBANK", "BAJFINANCE", "BHARTIARTL", "ITC", "LT", "HINDUNILVR",
    "MARUTI", "TATAMOTORS", "TATASTEEL", "ADANIENT", "ADANIPORTS",
    "SUNPHARMA", "TITAN", "ULTRACEMCO", "WIPRO", "HCLTECH", "ONGC",
    "NTPC", "POWERGRID", "M&M", "BAJAJFINSV", "ASIANPAINT", "DIVISLAB",
]

FETCH_WORKERS = 5  # a bit higher than the Excel version's default since
                    # there's no shared-with-other-scripts rate budget here


# ---------------------------------------------------------------------------
# Simple in-memory, thread-safe rate limiter (replaces the Excel version's
# cross-process file-locked one - not needed here since Streamlit serves
# every session from a single process/interpreter)
# ---------------------------------------------------------------------------
class RateLimiter:
    def __init__(self, limits):
        self.limits = limits
        self.max_period = max(period for _, period in limits)
        self.lock = threading.Lock()
        self.call_times = []
        self.penalty_until = 0.0

    def acquire(self):
        while True:
            with self.lock:
                now = time.time()
                if now < self.penalty_until:
                    wait = self.penalty_until - now
                else:
                    self.call_times = [t for t in self.call_times if now - t < self.max_period]
                    wait = 0.0
                    for max_calls, period in self.limits:
                        recent = [t for t in self.call_times if now - t < period]
                        if len(recent) >= max_calls:
                            wait = max(wait, period - (now - recent[0]))
                    if wait <= 0.0:
                        self.call_times.append(now)
                        return
            time.sleep(max(wait, 0.01))

    def report_429(self, cooldown):
        with self.lock:
            self.penalty_until = max(self.penalty_until, time.time() + cooldown)


RATE_LIMITER = RateLimiter(limits=[(3, 1.0), (120, 60.0)])
_SESSION = requests.Session()

INSTRUMENT_MASTER_URL = "https://assets.upstox.com/market-quote/instruments/exchange/complete.json.gz"
OPTION_CONTRACT_URL = "https://api.upstox.com/v2/option/contract"
OPTION_CHAIN_URL = "https://api.upstox.com/v2/option/chain"
MARKET_QUOTE_FULL_URL = "https://api.upstox.com/v2/market-quote/quotes"
MAX_QUOTE_KEYS_PER_CALL = 500
HISTORICAL_CANDLE_URL = "https://api.upstox.com/v2/historical-candle"
HISTORICAL_CANDLE_V3_URL = "https://api.upstox.com/v3/historical-candle"
INTRADAY_CANDLE_V3_URL = "https://api.upstox.com/v3/historical-candle/intraday"


def _get_with_retry(url, params, headers, max_retries=5):
    resp = None
    for attempt in range(max_retries):
        RATE_LIMITER.acquire()
        try:
            resp = _SESSION.get(url, params=params, headers=headers, timeout=30)
        except requests.exceptions.RequestException:
            time.sleep(1.0 * (attempt + 1))
            continue
        if resp.status_code == 429:
            cooldown = 2.0 * (2 ** attempt)
            RATE_LIMITER.report_429(cooldown)
            time.sleep(cooldown)
            continue
        return resp
    return resp


# ---------------------------------------------------------------------------
# Instrument universe / HV bands / lot size (all cached - expensive,
# once-a-day setup shared across every user session hitting this app)
# ---------------------------------------------------------------------------
def build_symbol_to_instrument_key():
    resp = _SESSION.get(INSTRUMENT_MASTER_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
    resp.raise_for_status()
    with gzip.GzipFile(fileobj=io.BytesIO(resp.content)) as gz:
        instruments = json.load(gz)
    lookup = {}
    for inst in instruments:
        if inst.get("segment") == "NSE_EQ" and inst.get("instrument_type") == "EQ":
            lookup[inst.get("trading_symbol")] = inst.get("instrument_key")
    return lookup, instruments


def discover_fno_underlying_keys(instruments):
    candidate_fields = ("underlying_key", "underlying_instrument_key", "asset_key")
    underlying_keys = set()
    for inst in instruments:
        if inst.get("segment") != "NSE_FO":
            continue
        for field in candidate_fields:
            key = inst.get(field)
            if isinstance(key, str) and key.startswith("NSE_EQ|"):
                underlying_keys.add(key)
                break
    return underlying_keys


def build_active_fno_stock_list(symbol_to_key, instruments):
    underlying_keys = discover_fno_underlying_keys(instruments)
    key_to_symbol = {v: k for k, v in symbol_to_key.items()}
    discovered = sorted({key_to_symbol[k] for k in underlying_keys if k in key_to_symbol})
    return discovered if len(discovered) >= 50 else FALLBACK_FNO_STOCKS


def build_lot_size_map(symbol_to_key, instruments):
    candidate_fields = ("underlying_key", "underlying_instrument_key", "asset_key")
    key_to_symbol = {v: k for k, v in symbol_to_key.items()}
    lot_size_map = {}
    for inst in instruments:
        if inst.get("segment") != "NSE_FO":
            continue
        lot_size = inst.get("lot_size")
        if not lot_size:
            continue
        for field in candidate_fields:
            uk = inst.get(field)
            if isinstance(uk, str) and uk.startswith("NSE_EQ|"):
                sym = key_to_symbol.get(uk)
                if sym and sym not in lot_size_map:
                    lot_size_map[sym] = lot_size
                break
    return lot_size_map


def fetch_daily_closes(instrument_key, headers, calendar_days=HV_CALENDAR_BUFFER_DAYS):
    to_date = date.today().strftime("%Y-%m-%d")
    from_date = (date.today() - timedelta(days=calendar_days)).strftime("%Y-%m-%d")
    url = f"{HISTORICAL_CANDLE_URL}/{instrument_key}/day/{to_date}/{from_date}"
    resp = _get_with_retry(url, {}, headers)
    if resp is None or resp.status_code != 200:
        return []
    candles = resp.json().get("data", {}).get("candles", [])
    return [c[4] for c in reversed(candles) if len(c) > 4 and c[4]]


def daily_historical_volatility_pct(closes, lookback_trading_days=HV_LOOKBACK_TRADING_DAYS):
    if len(closes) < 2:
        return None
    returns = [(closes[i] - closes[i - 1]) / closes[i - 1] for i in range(1, len(closes)) if closes[i - 1]]
    returns = returns[-lookback_trading_days:]
    if len(returns) < 2:
        return None
    return statistics.stdev(returns) * 100


def classify_stocks_by_hv(symbol_to_key, stocks, headers, progress_cb=None):
    high_vol_stocks, mid_vol_stocks = [], []
    prev_close_cache = {}
    for i, stock in enumerate(stocks, 1):
        instrument_key = symbol_to_key.get(stock)
        if not instrument_key:
            continue
        closes = fetch_daily_closes(instrument_key, headers)
        hv = daily_historical_volatility_pct(closes)
        if hv is None:
            continue
        if hv >= HV_THRESHOLD_PCT:
            high_vol_stocks.append(stock)
            band = "High-Vol"
        elif hv >= MIDVOL_HV_MIN:
            mid_vol_stocks.append(stock)
            band = "Mid-Vol"
        else:
            band = None
        if band and closes:
            prev_close_cache[stock] = closes[-1]
        if progress_cb:
            progress_cb(i, len(stocks))
    return high_vol_stocks, mid_vol_stocks, prev_close_cache


def get_nearest_expiry(instrument_key, headers):
    resp = _get_with_retry(OPTION_CONTRACT_URL, {"instrument_key": instrument_key}, headers)
    if resp is None or resp.status_code != 200:
        return None
    contracts = resp.json().get("data", [])
    if not contracts:
        return None
    today = date.today()
    expiries = sorted({c["expiry"] for c in contracts if c.get("expiry")})
    upcoming = [e for e in expiries if datetime.strptime(e, "%Y-%m-%d").date() >= today]
    return upcoming[0] if upcoming else (expiries[-1] if expiries else None)


def fetch_option_chain(instrument_key, expiry_date, headers):
    resp = _get_with_retry(OPTION_CHAIN_URL, {"instrument_key": instrument_key, "expiry_date": expiry_date}, headers)
    if resp is None or resp.status_code != 200:
        return None
    return resp.json().get("data", [])


# ---------------------------------------------------------------------------
# RVOL gate
# ---------------------------------------------------------------------------
def fetch_historical_days_candles(instrument_key, headers, lookback_days=RVOL_LOOKBACK_DAYS):
    calendar_days_needed = int(lookback_days * 1.6) + 5
    to_date = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    from_date = (date.today() - timedelta(days=calendar_days_needed)).strftime("%Y-%m-%d")
    url = f"{HISTORICAL_CANDLE_V3_URL}/{instrument_key}/minutes/1/{to_date}/{from_date}"
    resp = _get_with_retry(url, {}, headers)
    if resp is None or resp.status_code != 200:
        return {}
    candles = resp.json().get("data", {}).get("candles", [])
    by_date = {}
    for c in candles:
        if len(c) < 6 or not c[0]:
            continue
        by_date.setdefault(c[0][:10], []).append(c)
    for day_str in by_date:
        by_date[day_str].sort(key=lambda c: c[0])
    recent_days = sorted(by_date.keys(), reverse=True)[:lookback_days]
    return {d: by_date[d] for d in recent_days}


def _cumulative_volume_by_time(day_candles):
    result, running = [], 0
    for c in day_candles:
        running += (c[5] or 0)
        result.append((c[0][11:16], running))
    return result


def precompute_rvol_baselines(symbol_to_key, stocks, headers, progress_cb=None):
    baselines = {}
    for i, stock in enumerate(stocks, 1):
        instrument_key = symbol_to_key.get(stock)
        if not instrument_key:
            continue
        baselines[stock] = fetch_historical_days_candles(instrument_key, headers)
        if progress_cb:
            progress_cb(i, len(stocks))
    return baselines


def fetch_batch_volume(instrument_keys, headers):
    results = {}
    for i in range(0, len(instrument_keys), MAX_QUOTE_KEYS_PER_CALL):
        batch = instrument_keys[i:i + MAX_QUOTE_KEYS_PER_CALL]
        resp = _get_with_retry(MARKET_QUOTE_FULL_URL, {"instrument_key": ",".join(batch)}, headers)
        if resp is None or resp.status_code != 200:
            continue
        data = resp.json().get("data", {}) or {}
        for entry in data.values():
            key = entry.get("instrument_token")
            volume = entry.get("volume")
            if key and volume is not None:
                results[key] = volume
    return results


def compute_current_rvol(symbol_to_key, stocks, rvol_baselines, volume_data):
    now_time = datetime.now().strftime("%H:%M")
    rvol_by_stock = {}
    for stock in stocks:
        key = symbol_to_key.get(stock)
        baseline = rvol_baselines.get(stock)
        if not key or key not in volume_data or not baseline:
            continue
        today_volume = volume_data[key]
        historical_values = []
        for day_candles in baseline.values():
            matched = None
            for t, cum in _cumulative_volume_by_time(day_candles):
                if t <= now_time:
                    matched = cum
                else:
                    break
            if matched is not None:
                historical_values.append(matched)
        if not historical_values:
            continue
        avg_historical = sum(historical_values) / len(historical_values)
        if not avg_historical:
            continue
        rvol_by_stock[stock] = today_volume / avg_historical
    return rvol_by_stock


# ---------------------------------------------------------------------------
# SuperTrend / Vol Delta / Vol Surge
# ---------------------------------------------------------------------------
def fetch_intraday_candles(instrument_key, headers, unit=SUPERTREND_CANDLE_UNIT, interval=SUPERTREND_CANDLE_INTERVAL):
    url = f"{INTRADAY_CANDLE_V3_URL}/{instrument_key}/{unit}/{interval}"
    resp = _get_with_retry(url, {}, headers)
    if resp is None or resp.status_code != 200:
        return []
    return resp.json().get("data", {}).get("candles", [])


def fetch_recent_candles_for_supertrend(instrument_key, headers, unit=SUPERTREND_CANDLE_UNIT,
                                         interval=SUPERTREND_CANDLE_INTERVAL, history_days=SUPERTREND_HISTORY_DAYS):
    to_date = date.today().strftime("%Y-%m-%d")
    from_date = (date.today() - timedelta(days=history_days)).strftime("%Y-%m-%d")
    hist_url = f"{HISTORICAL_CANDLE_V3_URL}/{instrument_key}/{unit}/{interval}/{to_date}/{from_date}"
    hist_resp = _get_with_retry(hist_url, {}, headers)
    historical = []
    if hist_resp is not None and hist_resp.status_code == 200:
        historical = hist_resp.json().get("data", {}).get("candles", [])
    intraday = fetch_intraday_candles(instrument_key, headers, unit, interval)
    by_timestamp = {}
    for c in historical + intraday:
        if len(c) >= 6 and c[0]:
            by_timestamp[c[0]] = c
    return sorted(by_timestamp.values(), key=lambda c: c[0])


def compute_supertrend(candles, atr_period=SUPERTREND_ATR_PERIOD, multiplier=SUPERTREND_MULTIPLIER):
    n = len(candles)
    if n < atr_period + 2:
        return None, None, None
    highs = [c[2] for c in candles]
    lows = [c[3] for c in candles]
    closes = [c[4] for c in candles]

    true_ranges = []
    for i in range(n):
        if i == 0:
            true_ranges.append(highs[i] - lows[i])
        else:
            true_ranges.append(max(highs[i] - lows[i], abs(highs[i] - closes[i - 1]), abs(lows[i] - closes[i - 1])))

    atr = [None] * n
    atr[atr_period - 1] = sum(true_ranges[:atr_period]) / atr_period
    for i in range(atr_period, n):
        atr[i] = (atr[i - 1] * (atr_period - 1) + true_ranges[i]) / atr_period

    upper_band, lower_band, supertrend, direction = [None] * n, [None] * n, [None] * n, [None] * n
    start = atr_period - 1
    for i in range(start, n):
        hl2 = (highs[i] + lows[i]) / 2
        basic_upper = hl2 + multiplier * atr[i]
        basic_lower = hl2 - multiplier * atr[i]
        if i == start:
            upper_band[i], lower_band[i] = basic_upper, basic_lower
            direction[i] = -1 if closes[i] > basic_upper else 1
            supertrend[i] = lower_band[i] if direction[i] == -1 else upper_band[i]
            continue
        prev_upper, prev_lower = upper_band[i - 1], lower_band[i - 1]
        upper_band[i] = basic_upper if (basic_upper < prev_upper or closes[i - 1] > prev_upper) else prev_upper
        lower_band[i] = basic_lower if (basic_lower > prev_lower or closes[i - 1] < prev_lower) else prev_lower
        prev_supertrend = supertrend[i - 1]
        if prev_supertrend == prev_upper:
            direction[i] = -1 if closes[i] > upper_band[i] else 1
        else:
            direction[i] = 1 if closes[i] < lower_band[i] else -1
        supertrend[i] = lower_band[i] if direction[i] == -1 else upper_band[i]
    return closes[-1], supertrend[-1], direction[-1] == -1


def compute_latest_volume_delta(candles):
    if not candles:
        return None
    c = candles[-1]
    if len(c) < 6:
        return None
    open_, close_, volume = c[1], c[4], c[5]
    if close_ > open_:
        return volume
    if close_ < open_:
        return -volume
    return 0


def compute_volume_surge(candles, lookback=VOL_SURGE_LOOKBACK):
    if len(candles) < lookback + 1:
        return None
    latest_volume = candles[-1][5]
    baseline_candles = candles[-(lookback + 1):-1]
    baseline_volumes = [c[5] for c in baseline_candles if len(c) >= 6 and c[5] is not None]
    if not baseline_volumes:
        return None
    avg_volume = sum(baseline_volumes) / len(baseline_volumes)
    return (latest_volume / avg_volume) if avg_volume else None


def get_supertrend_vol_delta_surge(stock, instrument_key, headers, cache):
    entry = cache.get(stock)
    now = time.time()
    if entry and (now - entry["fetched_at"] < SUPERTREND_REFRESH_SEC):
        return entry["supertrend"], entry["vol_delta"], entry["vol_surge"]
    candles = fetch_recent_candles_for_supertrend(instrument_key, headers)
    _, supertrend_value, _ = compute_supertrend(candles)
    vol_delta_value = compute_latest_volume_delta(candles)
    vol_surge_value = compute_volume_surge(candles)
    cache[stock] = {"fetched_at": now, "supertrend": supertrend_value, "vol_delta": vol_delta_value, "vol_surge": vol_surge_value}
    return supertrend_value, vol_delta_value, vol_surge_value


# ---------------------------------------------------------------------------
# Net GEX / Zero-Gamma Flip / Squeeze Setup
# ---------------------------------------------------------------------------
def bs_gamma(spot, strike, iv_decimal, time_years, r=GEX_RISK_FREE_RATE):
    if not iv_decimal or iv_decimal <= 0 or not time_years or time_years <= 0 or not spot or spot <= 0 or not strike or strike <= 0:
        return None
    try:
        sqrt_t = math.sqrt(time_years)
        d1 = (math.log(spot / strike) + (r + 0.5 * iv_decimal ** 2) * time_years) / (iv_decimal * sqrt_t)
        pdf = math.exp(-0.5 * d1 * d1) / math.sqrt(2 * math.pi)
        return pdf / (spot * iv_decimal * sqrt_t)
    except (ValueError, ZeroDivisionError, OverflowError):
        return None


def _time_to_expiry_years(expiry_str, today=None):
    today = today or date.today()
    try:
        expiry_date = datetime.strptime(expiry_str, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return GEX_MIN_TIME_YEARS
    days = (expiry_date - today).days
    return max(days, 0) / 365.0 or GEX_MIN_TIME_YEARS


def compute_net_gex(chain_snapshot, spot, time_years, lot_size):
    total = 0.0
    any_computed = False
    for strike, ce_oi, ce_iv, pe_oi, pe_iv in chain_snapshot:
        if ce_oi and ce_iv:
            g = bs_gamma(spot, strike, ce_iv / 100.0, time_years)
            if g is not None:
                total += g * ce_oi * lot_size * spot * spot * 0.01
                any_computed = True
        if pe_oi and pe_iv:
            g = bs_gamma(spot, strike, pe_iv / 100.0, time_years)
            if g is not None:
                total -= g * pe_oi * lot_size * spot * spot * 0.01
                any_computed = True
    return total if any_computed else None


def compute_zero_gamma_flip(chain_snapshot, spot, time_years, lot_size):
    grid = sorted({s[0] for s in chain_snapshot})
    if len(grid) < 2:
        return None

    def net_gex_at(s_test):
        total = 0.0
        for strike, ce_oi, ce_iv, pe_oi, pe_iv in chain_snapshot:
            if ce_oi and ce_iv:
                g = bs_gamma(s_test, strike, ce_iv / 100.0, time_years)
                if g is not None:
                    total += g * ce_oi * lot_size * s_test * s_test * 0.01
            if pe_oi and pe_iv:
                g = bs_gamma(s_test, strike, pe_iv / 100.0, time_years)
                if g is not None:
                    total -= g * pe_oi * lot_size * s_test * s_test * 0.01
        return total

    values = [(s, net_gex_at(s)) for s in grid]
    crossings = []
    for (s0, v0), (s1, v1) in zip(values, values[1:]):
        if v0 == 0:
            crossings.append(s0)
        elif (v0 > 0) != (v1 > 0):
            frac = v0 / (v0 - v1) if (v0 - v1) else 0.5
            crossings.append(s0 + frac * (s1 - s0))
    if not crossings:
        return None
    return min(crossings, key=lambda c: abs(c - spot))


def evaluate_squeeze_setup(row):
    net_gex, gex_flip, spot = row.get("net_gex"), row.get("gex_flip"), row.get("spot")
    ce_strike, ce_chg = row.get("ce_strike"), row.get("ce_chg")
    vol_surge, vol_delta = row.get("vol_surge"), row.get("vol_delta")
    price_chg, above_supertrend = row.get("price_chg"), row.get("above_supertrend")

    if net_gex is None or net_gex >= 0:
        return False
    if gex_flip is None or spot is None or spot >= gex_flip:
        return False
    if (gex_flip - spot) / gex_flip * 100 > SQUEEZE_APPROACH_PCT_MAX:
        return False
    if ce_strike is None or ce_chg is None or ce_chg <= 0 or ce_strike <= spot:
        return False
    if vol_surge is None or vol_surge < VOL_SURGE_THRESHOLD:
        return False
    if vol_delta is None or vol_delta <= 0:
        return False
    if price_chg is None or price_chg <= 0:
        return False
    if above_supertrend is not True:
        return False
    return True


# ---------------------------------------------------------------------------
# Per-stock full fetch
# ---------------------------------------------------------------------------
def oi_change_pct(curr_oi, prev_oi):
    return ((curr_oi - prev_oi) / prev_oi) * 100 if prev_oi else None


def classify_moneyness(option_type, strike, spot, atm_strike):
    if strike == atm_strike:
        return "ATM"
    if option_type == "CE":
        return "ITM" if strike < spot else "OTM"
    return "ITM" if strike > spot else "OTM"


def process_one_stock(stock, symbol_to_key, expiry_cache, prev_close_cache, supertrend_cache,
                       rvol_value, band_label, lot_size, headers, oi_change_mode):
    instrument_key = symbol_to_key.get(stock)
    expiry = expiry_cache.get(stock)
    if not instrument_key or not expiry:
        return None

    chain = fetch_option_chain(instrument_key, expiry, headers)
    if not chain:
        return None

    spot = chain[0].get("underlying_spot_price", 0)
    strikes_only = [e for e in chain if e.get("strike_price") is not None]
    if not spot or not strikes_only:
        return None

    prev_close = prev_close_cache.get(stock)
    price_chg = ((spot - prev_close) / prev_close * 100) if prev_close else None
    atm_strike = min(strikes_only, key=lambda e: abs(e["strike_price"] - spot))["strike_price"]

    agg = {"CE_ITM": [0, 0], "PE_ITM": [0, 0], "CE_OTM": [0, 0], "PE_OTM": [0, 0]}
    iv_agg = {"CE_ITM": [0.0, 0], "PE_ITM": [0.0, 0], "CE_OTM": [0.0, 0], "PE_OTM": [0.0, 0]}
    ce_wall_strike, ce_wall_oi = None, -1
    pe_wall_strike, pe_wall_oi = None, -1
    ce_best = None
    pe_best = None
    chain_snapshot = []

    for entry in strikes_only:
        strike = entry["strike_price"]
        ce_md = (entry.get("call_options") or {}).get("market_data") or {}
        pe_md = (entry.get("put_options") or {}).get("market_data") or {}
        ce_greeks = (entry.get("call_options") or {}).get("option_greeks") or {}
        pe_greeks = (entry.get("put_options") or {}).get("option_greeks") or {}

        ce_oi_here = ce_md.get("oi", 0) or 0
        pe_oi_here = pe_md.get("oi", 0) or 0
        ce_prev_here = ce_md.get("prev_oi", 0) or 0
        pe_prev_here = pe_md.get("prev_oi", 0) or 0
        ce_iv_here = ce_greeks.get("iv")
        pe_iv_here = pe_greeks.get("iv")
        chain_snapshot.append((strike, ce_oi_here, ce_iv_here, pe_oi_here, pe_iv_here))

        if ce_oi_here > ce_wall_oi:
            ce_wall_oi, ce_wall_strike = ce_oi_here, strike
        if pe_oi_here > pe_wall_oi:
            pe_wall_oi, pe_wall_strike = pe_oi_here, strike

        ce_chg_here = ce_oi_here - ce_prev_here
        pe_chg_here = pe_oi_here - pe_prev_here
        if not (oi_change_mode == "increase" and ce_chg_here <= 0):
            ce_score = abs(ce_chg_here) if oi_change_mode == "abs" else ce_chg_here
            if ce_best is None or ce_score > ce_best[0]:
                ce_pct = ((ce_chg_here / ce_prev_here) * 100) if ce_prev_here else None
                ce_best = (ce_score, strike, ce_oi_here, ce_prev_here, ce_chg_here, ce_pct)
        if not (oi_change_mode == "increase" and pe_chg_here <= 0):
            pe_score = abs(pe_chg_here) if oi_change_mode == "abs" else pe_chg_here
            if pe_best is None or pe_score > pe_best[0]:
                pe_pct = ((pe_chg_here / pe_prev_here) * 100) if pe_prev_here else None
                pe_best = (pe_score, strike, pe_oi_here, pe_prev_here, pe_chg_here, pe_pct)

        for opt_type, md, greeks in (("CE", ce_md, ce_greeks), ("PE", pe_md, pe_greeks)):
            moneyness = classify_moneyness(opt_type, strike, spot, atm_strike)
            if moneyness == "ATM":
                continue
            oi = md.get("oi", 0) or 0
            prev_oi = md.get("prev_oi", 0) or 0
            bucket = agg[f"{opt_type}_{moneyness}"]
            bucket[0] += oi
            bucket[1] += prev_oi
            iv = greeks.get("iv")
            if iv is not None:
                iv_bucket = iv_agg[f"{opt_type}_{moneyness}"]
                iv_bucket[0] += iv
                iv_bucket[1] += 1

    def _summarize(bucket):
        oi, prev_oi = bucket
        return oi, prev_oi, oi_change_pct(oi, prev_oi)

    def _avg_iv(bucket):
        total, count = bucket
        return (total / count) if count else None

    def _combine_iv(a, b):
        vals = [v for v in (a, b) if v is not None]
        return (sum(vals) / len(vals)) if vals else None

    ce_itm_oi, ce_itm_prev, ce_itm_chg = _summarize(agg["CE_ITM"])
    pe_itm_oi, pe_itm_prev, pe_itm_chg = _summarize(agg["PE_ITM"])
    ce_otm_oi, ce_otm_prev, ce_otm_chg = _summarize(agg["CE_OTM"])
    pe_otm_oi, pe_otm_prev, pe_otm_chg = _summarize(agg["PE_OTM"])
    ce_iv = _combine_iv(_avg_iv(iv_agg["CE_ITM"]), _avg_iv(iv_agg["CE_OTM"]))
    pe_iv = _combine_iv(_avg_iv(iv_agg["PE_ITM"]), _avg_iv(iv_agg["PE_OTM"]))

    supertrend_value, vol_delta_value, vol_surge_value = get_supertrend_vol_delta_surge(
        stock, instrument_key, headers, supertrend_cache)
    above_supertrend = (spot > supertrend_value) if supertrend_value is not None else None

    time_years = _time_to_expiry_years(expiry)
    net_gex_raw = compute_net_gex(chain_snapshot, spot, time_years, lot_size)
    gex_flip = compute_zero_gamma_flip(chain_snapshot, spot, time_years, lot_size)
    net_gex_cr = (net_gex_raw / GEX_DISPLAY_DIVISOR) if net_gex_raw is not None else None

    if ce_best:
        _, ce_strike, ce_maxoi, ce_maxprev, ce_maxchg, ce_maxpct = ce_best
    else:
        ce_strike = ce_maxoi = ce_maxprev = ce_maxchg = ce_maxpct = None
    if pe_best:
        _, pe_strike, pe_maxoi, pe_maxprev, pe_maxchg, pe_maxpct = pe_best
    else:
        pe_strike = pe_maxoi = pe_maxprev = pe_maxchg = pe_maxpct = None

    row = {
        "band": band_label, "stock": stock, "rvol": rvol_value, "price_chg": price_chg,
        "spot": spot, "expiry": expiry,
        "ce_wall": ce_wall_strike, "pe_wall": pe_wall_strike,
        "ce_wall_oi": (ce_wall_oi if ce_wall_strike is not None else None),
        "pe_wall_oi": (pe_wall_oi if pe_wall_strike is not None else None),
        "ce_iv": ce_iv, "pe_iv": pe_iv,
        "itm_ce_oi": ce_itm_oi, "itm_ce_prev_oi": ce_itm_prev, "itm_ce_chg": ce_itm_chg,
        "itm_pe_oi": pe_itm_oi, "itm_pe_prev_oi": pe_itm_prev, "itm_pe_chg": pe_itm_chg,
        "otm_ce_oi": ce_otm_oi, "otm_ce_prev_oi": ce_otm_prev, "otm_ce_chg": ce_otm_chg,
        "otm_pe_oi": pe_otm_oi, "otm_pe_prev_oi": pe_otm_prev, "otm_pe_chg": pe_otm_chg,
        "ce_strike": ce_strike, "ce_oi": ce_maxoi, "ce_prev_oi": ce_maxprev,
        "ce_chg": ce_maxchg, "ce_chg_pct": ce_maxpct,
        "pe_strike": pe_strike, "pe_oi": pe_maxoi, "pe_prev_oi": pe_maxprev,
        "pe_chg": pe_maxchg, "pe_chg_pct": pe_maxpct,
        "supertrend": supertrend_value, "above_supertrend": above_supertrend,
        "vol_delta": vol_delta_value, "vol_surge": vol_surge_value,
        "net_gex": net_gex_cr, "gex_flip": gex_flip,
    }
    row["squeeze_setup"] = evaluate_squeeze_setup(row)
    return row


def run_cycle(qualifying_stocks, symbol_to_key, expiry_cache, prev_close_cache, supertrend_cache,
              rvol_by_stock, band_of_stock, lot_size_map, headers, oi_change_mode, fetch_workers):
    rows = []
    with ThreadPoolExecutor(max_workers=fetch_workers) as pool:
        futures = {
            pool.submit(process_one_stock, stock, symbol_to_key, expiry_cache, prev_close_cache,
                        supertrend_cache, rvol_by_stock[stock], band_of_stock.get(stock, "?"),
                        lot_size_map.get(stock, GEX_DEFAULT_LOT_SIZE), headers, oi_change_mode): stock
            for stock in qualifying_stocks
        }
        for future in as_completed(futures):
            row = future.result()
            if row:
                rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Cached one-time setup (shared across all sessions hitting this app; keyed
# by access token so a token change busts the cache automatically)
# ---------------------------------------------------------------------------
@st.cache_resource(ttl=6 * 3600, show_spinner=False)
def load_universe(access_token):
    headers = {"Accept": "application/json", "Authorization": f"Bearer {access_token}"}
    symbol_to_key, instruments = build_symbol_to_instrument_key()
    candidate_stocks = build_active_fno_stock_list(symbol_to_key, instruments)
    high_vol_stocks, mid_vol_stocks, prev_close_cache = classify_stocks_by_hv(symbol_to_key, candidate_stocks, headers)
    all_tracked_stocks = high_vol_stocks + mid_vol_stocks
    band_of_stock = {s: "High-Vol" for s in high_vol_stocks}
    band_of_stock.update({s: "Mid-Vol" for s in mid_vol_stocks})
    lot_size_map = build_lot_size_map(symbol_to_key, instruments)
    rvol_baselines = precompute_rvol_baselines(symbol_to_key, all_tracked_stocks, headers)
    expiry_cache = {}
    for stock in all_tracked_stocks:
        instrument_key = symbol_to_key.get(stock)
        if instrument_key:
            expiry = get_nearest_expiry(instrument_key, headers)
            if expiry:
                expiry_cache[stock] = expiry
    all_instrument_keys = [symbol_to_key[s] for s in all_tracked_stocks if s in symbol_to_key]
    return {
        "symbol_to_key": symbol_to_key, "all_tracked_stocks": all_tracked_stocks,
        "band_of_stock": band_of_stock, "prev_close_cache": prev_close_cache,
        "lot_size_map": lot_size_map, "rvol_baselines": rvol_baselines,
        "expiry_cache": expiry_cache, "all_instrument_keys": all_instrument_keys,
        "high_vol_count": len(high_vol_stocks), "mid_vol_count": len(mid_vol_stocks),
    }


def run_one_cycle(access_token, rvol_threshold, max_fetch_per_cycle, oi_change_mode, fetch_workers):
    headers = {"Accept": "application/json", "Authorization": f"Bearer {access_token}"}
    data = load_universe(access_token)

    volume_data = fetch_batch_volume(data["all_instrument_keys"], headers)
    rvol_by_stock = compute_current_rvol(data["symbol_to_key"], data["all_tracked_stocks"], data["rvol_baselines"], volume_data)
    qualifying_stocks = [s for s, v in rvol_by_stock.items() if v >= rvol_threshold]
    if max_fetch_per_cycle and len(qualifying_stocks) > max_fetch_per_cycle:
        qualifying_stocks.sort(key=lambda s: rvol_by_stock[s], reverse=True)
        qualifying_stocks = qualifying_stocks[:max_fetch_per_cycle]

    if "supertrend_cache" not in st.session_state:
        st.session_state.supertrend_cache = {}

    rows = run_cycle(qualifying_stocks, data["symbol_to_key"], data["expiry_cache"], data["prev_close_cache"],
                      st.session_state.supertrend_cache, rvol_by_stock, data["band_of_stock"],
                      data["lot_size_map"], headers, oi_change_mode, fetch_workers)
    return rows, data, len(qualifying_stocks), len(data["all_tracked_stocks"])


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="OI ULTRA LIVE", page_icon="\U0001F4C8", layout="wide")

st.markdown("""
<style>
/* Tighter padding + bigger tap targets for phone screens */
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
div[data-testid="stMetricValue"] {font-size: 1.3rem;}
</style>
""", unsafe_allow_html=True)

st.title("\U0001F4C8 OI ULTRA LIVE")

with st.sidebar:
    st.header("Settings")
    def _get_secret_token():
        try:
            return st.secrets.get("UPSTOX_ACCESS_TOKEN", "")
        except Exception:
            # No secrets.toml at all yet - that's fine, just means no
            # pre-filled default; the sidebar text box still works.
            return ""
    default_token = _get_secret_token()
    access_token = st.text_input("Upstox access token (today's)", value=default_token, type="password",
                                  help="Paste today's token from your upstox_login.py flow. "
                                       "Set UPSTOX_ACCESS_TOKEN in Streamlit secrets to pre-fill this.")
    rvol_threshold = st.slider("RVOL threshold", 1.0, 5.0, 1.25, 0.05)
    max_fetch_per_cycle = st.slider("Max stocks fetched per cycle", 5, 100, 40, 5,
                                     help="Caps API calls per refresh - lower = faster refresh on mobile data.")
    oi_change_mode = st.radio("OI change mode", ["increase", "abs"], index=0,
                               help="'increase': only OI build-ups count. 'abs': biggest move either way wins.")
    refresh_interval = st.slider("Auto-refresh interval (sec)", 30, 300, 60, 10)
    auto_refresh_on = st.checkbox("Auto-refresh", value=True)
    show_details = st.checkbox("Show Details table (raw OI numbers)", value=False)
    st.caption(f"Auto-refresh via streamlit-autorefresh: {'available' if HAS_AUTOREFRESH else 'NOT installed - use the manual button below'}")

if not access_token:
    st.warning("Paste today's Upstox access token in the sidebar to start.")
    st.stop()

if HAS_AUTOREFRESH and auto_refresh_on:
    st_autorefresh(interval=refresh_interval * 1000, key="oi_ultra_autorefresh")

col_a, col_b = st.columns([1, 1])
with col_a:
    manual_refresh = st.button("\U0001F504 Refresh now", use_container_width=True)
with col_b:
    st.caption(f"Last refreshed: {datetime.now().strftime('%H:%M:%S')}")

try:
    with st.spinner("Fetching live data..."):
        rows, data, n_qualifying, n_tracked = run_one_cycle(
            access_token, rvol_threshold, max_fetch_per_cycle, oi_change_mode, FETCH_WORKERS)
except requests.exceptions.HTTPError as e:
    st.error(f"Upstox API error - check your access token is current: {e}")
    st.stop()
except Exception as e:
    st.error(f"Something went wrong fetching data: {e}")
    st.stop()

squeeze_rows = [r for r in rows if r["squeeze_setup"]]

m1, m2, m3, m4 = st.columns(4)
m1.metric("High-Vol stocks", data["high_vol_count"])
m2.metric("Mid-Vol stocks", data["mid_vol_count"])
m3.metric("RVOL-qualifying", n_qualifying)
m4.metric("\U0001F680 Squeeze setups", len(squeeze_rows))

# --- Squeeze Setups: mobile-friendly stacked cards ---
if squeeze_rows:
    st.subheader("\U0001F680 Squeeze Setups")
    for r in sorted(squeeze_rows, key=lambda r: r["rvol"], reverse=True):
        with st.container(border=True):
            c1, c2, c3 = st.columns(3)
            c1.markdown(f"**{r['stock']}** ({r['band']})")
            c1.caption(f"RVOL {r['rvol']:.2f}")
            c2.markdown(f"Spot **{r['spot']}**")
            c2.caption(f"Price {r['price_chg']:+.1f}%" if r["price_chg"] is not None else "Price N/A")
            c3.markdown(f"GEX Flip **{r['gex_flip']:.1f}**" if r["gex_flip"] is not None else "GEX Flip N/A")
            c3.caption(f"Net GEX {r['net_gex']:.2f}Cr" if r["net_gex"] is not None else "Net GEX N/A")
            st.caption(f"CE Wall {r['ce_wall']} | PE Wall {r['pe_wall']} | "
                       f"Max ΔOI CE {r['ce_strike']} (+{r['ce_chg']:.0f}) | Expiry {r['expiry']}")

# --- Main table ---
st.subheader("All qualifying stocks")
if not rows:
    st.info("No stocks currently qualify at this RVOL threshold.")
else:
    df = pd.DataFrame(rows).sort_values("rvol", ascending=False)
    main_cols = {
        "stock": "Stock", "band": "Vol Band", "rvol": "RVOL", "vol_surge": "Vol Surge", "vol_delta": "Vol Delta",
        "price_chg": "Price Chg %", "spot": "Spot", "ce_wall": "CE Wall", "pe_wall": "PE Wall",
        "ce_iv": "CE IV", "pe_iv": "PE IV",
        "itm_ce_chg": "ITM CE Chg%", "otm_ce_chg": "OTM CE Chg%", "itm_pe_chg": "ITM PE Chg%", "otm_pe_chg": "OTM PE Chg%",
        "ce_strike": "CE Strike (MaxΔOI)", "ce_chg": "CE OI Chg", "ce_chg_pct": "CE OI Chg%",
        "pe_chg_pct": "PE OI Chg%", "pe_chg": "PE OI Chg", "pe_strike": "PE Strike (MaxΔOI)",
        "supertrend": "SuperTrend", "net_gex": "Net GEX (Cr)", "gex_flip": "GEX Flip",
        "squeeze_setup": "Squeeze",
    }
    main_df = df[list(main_cols.keys())].rename(columns=main_cols)
    main_df["Squeeze"] = main_df["Squeeze"].map({True: "\U0001F680", False: ""})

    def _color_signed(val):
        if not isinstance(val, (int, float)):
            return ""
        return "color: #1E7B34; font-weight: bold" if val > 0 else ("color: #C00000; font-weight: bold" if val < 0 else "")

    def _style_apply(styler, func, subset):
        # pandas >=2.1 renamed Styler.applymap -> Styler.map (and removed
        # applymap entirely in pandas 3.0) - support both.
        method = getattr(styler, "map", None) or styler.applymap
        return method(func, subset=subset)

    styled = main_df.style
    styled = _style_apply(styled, _color_signed,
                           ["ITM CE Chg%", "OTM CE Chg%", "ITM PE Chg%", "OTM PE Chg%", "Net GEX (Cr)", "PE OI Chg"])
    styled = _style_apply(styled, lambda v: "background-color: #F4B7B2", ["CE OI Chg"])
    styled = styled.format(precision=2)
    st.dataframe(styled, use_container_width=True, height=min(600, 60 + 35 * len(main_df)))

    if show_details:
        st.subheader("Details (raw OI numbers)")
        details_cols = {
            "stock": "Stock", "ce_wall_oi": "CE Wall OI", "pe_wall_oi": "PE Wall OI",
            "ce_oi": "CE OI", "ce_prev_oi": "CE Prev OI", "pe_oi": "PE OI", "pe_prev_oi": "PE Prev OI",
            "itm_ce_oi": "ITM CE OI", "itm_ce_prev_oi": "ITM CE Prev OI",
            "itm_pe_oi": "ITM PE OI", "itm_pe_prev_oi": "ITM PE Prev OI",
            "otm_ce_oi": "OTM CE OI", "otm_ce_prev_oi": "OTM CE Prev OI",
            "otm_pe_oi": "OTM PE OI", "otm_pe_prev_oi": "OTM PE Prev OI",
            "expiry": "Expiry",
        }
        details_df = df[list(details_cols.keys())].rename(columns=details_cols)
        st.dataframe(details_df, use_container_width=True)

st.caption("Not investment advice. RVOL/GEX/Squeeze Setup are probability tilts based on modeling "
           "assumptions (dealer sign convention, flat risk-free rate, static IV per strike) - see the "
           "comments in OI_ULTRA_LIVE.py for the full caveats.")
