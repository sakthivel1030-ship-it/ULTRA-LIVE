"""
OI ULTRA LIVE - merged NSE F&O options tracker, via the Upstox API -
LIVE EXCEL (COM) VERSION.

THIS IS A MERGE OF TWO EARLIER SCRIPTS INTO ONE:
  - UPSTOX_ITM_OTM_LIVE_TRACKER_EXCEL_COM_HVRVOL_MOD.py (the "HVROL
    dashboard"): auto-discovers the live F&O universe, classifies it into
    High-Vol / Mid-Vol bands by daily historical volatility, aggregates
    ITM/OTM CE/PE open interest and its %-change, tracks CE Wall/PE Wall
    (max outstanding OI strike), and computes SuperTrend, Vol Delta and
    Vol Surge per stock.
  - UPSTOX_RVOL_OI_CHANGE_LIVE_TRACKER.py (the "RVOL tracker"): gates
    which stocks get the expensive per-cycle option-chain fetch by RVOL
    (only stocks actually trading unusually heavily today), and surfaces
    the single CE strike and single PE strike building up the most OI
    that day.

WHAT ONE-FILE, MERGED CONCEPT MEANS HERE, EVERY REFRESH CYCLE (default
60 seconds):
  1. RVOL GATE (from the RVOL tracker): every stock in the auto-discovered
     F&O universe (already split into High-Vol/Mid-Vol bands by the HVROL
     dashboard's daily-HV classification) gets a cheap, whole-universe
     RVOL check. Only stocks at/above RVOL_MIN_THRESHOLD get the full,
     expensive per-cycle fetch below - this is what keeps the per-cycle
     API load bounded regardless of how big the tracked universe is.
  2. FULL FETCH (from the HVROL dashboard), for qualifying stocks only:
       - ITM/OTM CE/PE open interest and OI %-change (ATM excluded)
       - CE Wall / PE Wall - the single strike with the highest
         OUTSTANDING OI for each leg, across the WHOLE chain (ATM
         included)
       - SuperTrend, Vol Delta, Vol Surge (3-minute candles, cached and
         refreshed independently of the main cycle - see
         SUPERTREND_REFRESH_SEC)
  3. MAX-OI-CHANGE STRIKES (from the RVOL tracker): the single CE strike
     and single PE strike building up the most OI that day (current OI
     minus Upstox's prev_oi), across the whole chain.
  4. Everything lands in ONE live, visible Excel sheet via win32com - a
     GROWING, SELF-SORTING HISTORY (not a snapshot): every cycle's rows
     are added to whatever's already on the sheet today, then the whole
     sheet is re-sorted so that stocks are grouped together, the stock
     whose RVOL has been HIGHEST at any point today sits at the very
     top, and rows WITHIN each stock's own group run chronologically,
     oldest first.

WHY COM (a real, visible, running Excel) INSTEAD OF JUST WRITING AN
.xlsx FILE:
  - A live, auto-updating sheet needs an already-open Excel window to
    update in place - a plain file write can't do that.
  - Only works on WINDOWS, with a real licensed copy of Excel installed.
  - Requires the `pywin32` package (`pip install pywin32`).
  - Screen-flicker during each cycle's write is suppressed via
    ScreenUpdating/Calculation/EnableEvents (see _frozen_excel below).

SETUP:
  - Run upstox_login.py first (or otherwise refresh access_token.txt for
    today) in this same folder.
  - pip install requests pywin32
  - Must be run on Windows, with Microsoft Excel installed and licensed.

USAGE:
  python OI_ULTRA_LIVE.py
  -> launches (or reuses) a visible Excel window, opens/creates
     OI_ULTRA_LIVE_YYYY-MM-DD.xlsx, and refreshes it every
     REFRESH_INTERVAL_SEC seconds.
  -> by default only runs 9:15 AM - 3:15 PM IST, Mon-Fri; set
     MARKET_HOURS_ONLY = False below to run continuously instead.
  -> after every successful refresh cycle, a timestamped snapshot of the
     workbook is saved into backups/YYYY-MM-DD/ (SaveCopyAs, doesn't
     touch the live open workbook) - so a crash or corruption in the
     live file only ever costs one cycle's worth of data. At market
     close, one final end-of-day archive is saved and the intraday
     snapshots are cleaned up.
  -> Ctrl+C to stop.

ASSUMPTIONS MADE (change the constants below if any of these don't match
what you actually want):
  - RVOL_MIN_THRESHOLD = 1.25, SAME threshold applied to both the
    High-Vol and Mid-Vol bands (not two separate thresholds), computed
    the same way as both source scripts (today's cumulative volume vs.
    a 10-trading-day baseline at the same time of day).
  - MAX_FETCH_PER_CYCLE = 80 caps how many RVOL-qualifying stocks get
    the full per-cycle option-chain fetch, ranked by RVOL descending, so
    a day where hundreds of stocks all qualify at once can't blow out
    the per-cycle API budget. Set to None to remove the cap entirely.
  - "Highest change in OI" = the single CE strike and single PE strike
    per stock with the largest OI INCREASE (oi - prev_oi) for the day,
    across the whole chain (ATM included). Set OI_CHANGE_MODE = "abs" to
    let the biggest move either direction (including unwinds) win
    instead.
  - ONE FLAT sheet for both bands (not the HVROL dashboard's old
    side-by-side High-Vol/Mid-Vol column blocks) - a "Vol Band" column
    tells you which band a stock is in, which keeps one comprehensive,
    growing history simple to read rather than split across two areas.
  - The sheet keeps EVERY row seen today (grows across cycles) and
    re-sorts itself every cycle: stock groups ordered by that stock's
    highest RVOL today (descending), rows within a group chronological
    (oldest first) - same behavior you asked for on the RVOL tracker.
  - "CE OI Chg" (the max-OI-change column, CE side) keeps the fixed red
    highlight you asked for previously; "PE OI Chg" keeps the
    green-if-up/red-if-down conditional coloring.
"""

import os
import sys
import time
import json
import gzip
import io
import math
import statistics
from datetime import datetime, date, time as dtime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

if sys.platform != "win32":
    print("This script requires Windows (it drives Excel via COM automation).")
    sys.exit(1)

try:
    import win32com.client as win32
    import pywintypes
except ImportError:
    print("Missing dependency: pywin32.")
    print("Install it with:  pip install pywin32")
    sys.exit(1)

import msvcrt  # cross-process file locking for the shared rate limiter


# ---------------------------------------------------------------------------
# Settings - change these to match your setup
# ---------------------------------------------------------------------------
RVOL_MIN_THRESHOLD = 1.25       # only stocks at/above this RVOL get the full fetch
RVOL_LOOKBACK_DAYS = 10         # trading days used for the RVOL baseline
MAX_FETCH_PER_CYCLE = 80        # cap on RVOL-qualifying stocks fetched per cycle (None = no cap)
REFRESH_INTERVAL_SEC = 60       # 1 minute

# "increase" -> only positive OI changes count as a candidate max (a stock
# where every strike unwound OI shows blank CE/PE Max-Chg columns).
# "abs"      -> the strike with the single LARGEST OI move either way wins,
#               even if it's actually an unwind (a big negative number).
OI_CHANGE_MODE = "increase"

HV_THRESHOLD_PCT = 2.0     # daily HV >= this -> High-Vol band
MIDVOL_HV_MIN = 1.5        # daily HV in [MIDVOL_HV_MIN, HV_THRESHOLD_PCT) -> Mid-Vol band
HV_LOOKBACK_TRADING_DAYS = 20
HV_CALENDAR_BUFFER_DAYS = 40

SUPERTREND_CANDLE_UNIT = "minutes"
SUPERTREND_CANDLE_INTERVAL = "3"
SUPERTREND_ATR_PERIOD = 10
SUPERTREND_MULTIPLIER = 3.0
SUPERTREND_HISTORY_DAYS = 3
SUPERTREND_REFRESH_SEC = 180  # SuperTrend/Vol Delta/Vol Surge refresh cadence (independent of REFRESH_INTERVAL_SEC)

VOL_SURGE_LOOKBACK = 10
VOL_SURGE_THRESHOLD = 2.0

# ---------------------------------------------------------------------------
# Net GEX / Zero-Gamma Flip settings
# ---------------------------------------------------------------------------
# Gamma is computed OURSELVES via Black-Scholes from each strike's own IV,
# rather than trusting Upstox's own greeks.gamma field - this is what makes
# it possible to also re-price gamma at HYPOTHETICAL spot levels for the
# zero-gamma flip sweep below, using the exact same formula the "current
# spot" Net GEX number uses, so the two figures are internally consistent
# (one model, evaluated at different spot points) rather than mixing a
# real greek at today's spot with a model-implied one everywhere else.
#
# SIGN CONVENTION (this varies by GEX provider - there is no universal
# standard): dealers are assumed net LONG gamma from calls and net SHORT
# gamma from puts (the common "SqueezeMetrics-style" convention), so:
#     Net GEX = Σ(call gamma × call OI) - Σ(put gamma × put OI)
# scaled by lot size and spot^2. If your convention is the opposite,
# flip the sign in compute_net_gex() below.
GEX_RISK_FREE_RATE = 0.065     # flat annualized risk-free rate assumption (India ~repo rate)
GEX_MIN_TIME_YEARS = 1 / 365   # floor on time-to-expiry so expiry day doesn't divide by ~0
GEX_DEFAULT_LOT_SIZE = 1       # fallback if a stock's lot size can't be found in the instrument master
GEX_DISPLAY_DIVISOR = 1e7      # Net GEX is shown in ₹ Crore-equivalent units for readability

# "Squeeze Setup" flag - a probability tilt, not a signal to trade blind:
# flags stocks where negative Net GEX (dealers short gamma, so hedging
# AMPLIFIES moves rather than damping them) lines up with spot sitting
# just below the zero-gamma flip level (the "fuel zone" - once spot
# crosses above it, GEX typically flips positive and the move tends to
# get pinned/faded instead of continuing), fresh call OI building up at
# a strike ABOVE spot (a dealer short-call magnet), a volume surge with
# net buying, and price already trending up with SuperTrend confirming.
# GEX sign conventions vary by provider and squeezes can fail to
# trigger or reverse hard past the flip level - see evaluate_squeeze_setup().
SQUEEZE_APPROACH_PCT_MAX = 3.0  # spot must be within this % BELOW the flip level to count as "approaching"

MARKET_HOURS_ONLY = True
MARKET_OPEN = dtime(9, 15)
MARKET_CLOSE = dtime(15, 15)

FALLBACK_FNO_STOCKS = [
    "RELIANCE", "TCS", "HDFCBANK", "ICICIBANK", "INFY", "SBIN", "AXISBANK",
    "KOTAKBANK", "BAJFINANCE", "BHARTIARTL", "ITC", "LT", "HINDUNILVR",
    "MARUTI", "TATAMOTORS", "TATASTEEL", "ADANIENT", "ADANIPORTS",
    "SUNPHARMA", "TITAN", "ULTRACEMCO", "WIPRO", "HCLTECH", "ONGC",
    "NTPC", "POWERGRID", "M&M", "BAJAJFINSV", "ASIANPAINT", "DIVISLAB",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = os.path.join(SCRIPT_DIR, f"OI_ULTRA_LIVE_{date.today().strftime('%Y-%m-%d')}.xlsx")
MAIN_SHEET_NAME = "OI ULTRA LIVE"
DETAILS_SHEET_NAME = "OI ULTRA LIVE - Details"

BACKUP_DIR = os.path.join(SCRIPT_DIR, "backups", date.today().strftime("%Y-%m-%d"))
os.makedirs(BACKUP_DIR, exist_ok=True)
BACKUP_LABEL = "OI_ULTRA_LIVE"
BACKUP_RETENTION_COUNT = 200  # intraday snapshots kept per label; None = keep all

FETCH_WORKERS = 3  # how many stocks' option chains to fetch concurrently


# ---------------------------------------------------------------------------
# Shared cross-process rate limiter (same state file the other tracker
# scripts in this folder use, so running this alongside them still stays
# under Upstox's combined per-token limit)
# ---------------------------------------------------------------------------
class _FileLock:
    def __init__(self, path):
        self.path = path
        self.fh = None

    def __enter__(self):
        self.fh = open(self.path, "a+b")
        while True:
            try:
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
                return self
            except OSError:
                time.sleep(0.01)

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.fh.seek(0)
            msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
        finally:
            self.fh.close()


class RateLimiter:
    def __init__(self, limits, state_file):
        self.limits = limits
        self.max_period = max(period for _, period in limits)
        self.state_file = state_file
        self.lock_file = state_file + ".lock"
        if not os.path.exists(self.state_file):
            try:
                self._write_state_unlocked({"call_times": [], "penalty_until": 0.0})
            except FileExistsError:
                pass

    def _read_state_unlocked(self):
        try:
            with open(self.state_file, "r") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError, ValueError):
            return {"call_times": [], "penalty_until": 0.0}

    def _write_state_unlocked(self, state):
        tmp_path = f"{self.state_file}.tmp{os.getpid()}"
        with open(tmp_path, "w") as f:
            json.dump(state, f)
        os.replace(tmp_path, self.state_file)

    def acquire(self):
        while True:
            with _FileLock(self.lock_file):
                state = self._read_state_unlocked()
                now = time.time()
                penalty_until = state.get("penalty_until", 0.0)
                if now < penalty_until:
                    wait = penalty_until - now
                else:
                    call_times = [t for t in state.get("call_times", []) if now - t < self.max_period]
                    wait = 0.0
                    for max_calls, period in self.limits:
                        recent = [t for t in call_times if now - t < period]
                        if len(recent) >= max_calls:
                            wait = max(wait, period - (now - recent[0]))
                    if wait <= 0.0:
                        call_times.append(now)
                        state["call_times"] = call_times
                        self._write_state_unlocked(state)
                        return
            time.sleep(max(wait, 0.01))

    def report_429(self, cooldown):
        with _FileLock(self.lock_file):
            state = self._read_state_unlocked()
            state["penalty_until"] = max(state.get("penalty_until", 0.0), time.time() + cooldown)
            self._write_state_unlocked(state)


MAX_CALLS_PER_SEC = 3
MAX_CALLS_PER_MIN = 120
RATE_LIMIT_STATE_FILE = os.path.join(SCRIPT_DIR, ".upstox_shared_rate_limit.json")
RATE_LIMITER = RateLimiter(
    limits=[(MAX_CALLS_PER_SEC, 1.0), (MAX_CALLS_PER_MIN, 60.0)],
    state_file=RATE_LIMIT_STATE_FILE)

_SESSION = requests.Session()

with open(os.path.join(SCRIPT_DIR, "access_token.txt")) as f:
    ACCESS_TOKEN = f.read().strip()
HEADERS = {"Accept": "application/json", "Authorization": f"Bearer {ACCESS_TOKEN}"}

INSTRUMENT_MASTER_URL = "https://assets.upstox.com/market-quote/instruments/exchange/complete.json.gz"
OPTION_CONTRACT_URL = "https://api.upstox.com/v2/option/contract"
OPTION_CHAIN_URL = "https://api.upstox.com/v2/option/chain"
MARKET_QUOTE_FULL_URL = "https://api.upstox.com/v2/market-quote/quotes"
MAX_QUOTE_KEYS_PER_CALL = 500
HISTORICAL_CANDLE_URL = "https://api.upstox.com/v2/historical-candle"
HISTORICAL_CANDLE_V3_URL = "https://api.upstox.com/v3/historical-candle"
INTRADAY_CANDLE_V3_URL = "https://api.upstox.com/v3/historical-candle/intraday"


def _get_with_retry(url, params, max_retries=5):
    resp = None
    for attempt in range(max_retries):
        RATE_LIMITER.acquire()
        try:
            resp = _SESSION.get(url, params=params, headers=HEADERS, timeout=30)
        except requests.exceptions.RequestException as e:
            print(f"    [network] {type(e).__name__} on attempt {attempt + 1}/{max_retries} "
                  f"for {url.split('?')[0]} - retrying...")
            time.sleep(1.0 * (attempt + 1))
            continue
        if resp.status_code == 429:
            cooldown = 2.0 * (2 ** attempt)
            RATE_LIMITER.report_429(cooldown)
            print(f"    [429] rate-limited on attempt {attempt + 1}/{max_retries} for "
                  f"{url.split('?')[0]} - backing off {cooldown:.0f}s (all workers pause)...")
            time.sleep(cooldown)
            continue
        return resp
    return resp


# ---------------------------------------------------------------------------
# Instrument universe
# ---------------------------------------------------------------------------
def build_symbol_to_instrument_key(retry_delay_sec=30):
    print("Downloading Upstox instrument master...")
    attempt = 0
    while True:
        attempt += 1
        try:
            resp = _SESSION.get(INSTRUMENT_MASTER_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
            resp.raise_for_status()
            break
        except requests.exceptions.RequestException as e:
            print(f"  [network] {type(e).__name__} on attempt {attempt} - retrying in "
                  f"{retry_delay_sec}s... (Ctrl+C to give up)")
            time.sleep(retry_delay_sec)

    with gzip.GzipFile(fileobj=io.BytesIO(resp.content)) as gz:
        instruments = json.load(gz)

    lookup = {}
    for inst in instruments:
        if inst.get("segment") == "NSE_EQ" and inst.get("instrument_type") == "EQ":
            lookup[inst.get("trading_symbol")] = inst.get("instrument_key")

    print(f"Loaded {len(lookup)} NSE equity instruments.\n")
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


def build_active_fno_stock_list(symbol_to_key, instruments, fallback_list=FALLBACK_FNO_STOCKS):
    underlying_keys = discover_fno_underlying_keys(instruments)
    key_to_symbol = {v: k for k, v in symbol_to_key.items()}
    discovered = sorted({key_to_symbol[k] for k in underlying_keys if k in key_to_symbol})

    MIN_SANE_COUNT = 50
    if len(discovered) >= MIN_SANE_COUNT:
        print(f"Auto-discovered {len(discovered)} stocks with LIVE F&O contracts.\n")
        return discovered

    print(f"[warn] Could only auto-discover {len(discovered)} F&O stock(s) - "
          f"falling back to the built-in list ({len(fallback_list)} names).\n")
    return fallback_list


def build_lot_size_map(symbol_to_key, instruments):
    """Scans the already-downloaded instrument master for each F&O
    stock's option lot size (needed to convert gamma+OI into a rupee-ish
    Net GEX figure) - no extra API call, since `instruments` is the same
    payload build_symbol_to_instrument_key already fetched. Lot size is
    effectively constant across strikes/expiries for a given underlying,
    so the first NSE_FO contract found for that underlying is used."""
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


def get_nearest_expiry(instrument_key):
    resp = _get_with_retry(OPTION_CONTRACT_URL, {"instrument_key": instrument_key})
    if resp is None or resp.status_code != 200:
        return None
    contracts = resp.json().get("data", [])
    if not contracts:
        return None
    today = date.today()
    expiries = sorted({c["expiry"] for c in contracts if c.get("expiry")})
    upcoming = [e for e in expiries if datetime.strptime(e, "%Y-%m-%d").date() >= today]
    return upcoming[0] if upcoming else (expiries[-1] if expiries else None)


def fetch_option_chain(instrument_key, expiry_date):
    resp = _get_with_retry(OPTION_CHAIN_URL, {"instrument_key": instrument_key, "expiry_date": expiry_date})
    if resp is None:
        return None, "no response after retries (network failure)"
    if resp.status_code != 200:
        snippet = resp.text[:150].replace("\n", " ").strip()
        return None, f"HTTP {resp.status_code}: {snippet}"
    return resp.json().get("data", []), None


def fetch_daily_closes(instrument_key, calendar_days=HV_CALENDAR_BUFFER_DAYS):
    to_date = date.today().strftime("%Y-%m-%d")
    from_date = (date.today() - timedelta(days=calendar_days)).strftime("%Y-%m-%d")
    url = f"{HISTORICAL_CANDLE_URL}/{instrument_key}/day/{to_date}/{from_date}"
    resp = _get_with_retry(url, {})
    if resp is None or resp.status_code != 200:
        return []
    candles = resp.json().get("data", {}).get("candles", [])
    return [c[4] for c in reversed(candles) if len(c) > 4 and c[4]]


# ---------------------------------------------------------------------------
# HV band classification (High-Vol / Mid-Vol) - from the HVROL dashboard
# ---------------------------------------------------------------------------
def daily_historical_volatility_pct(closes, lookback_trading_days=HV_LOOKBACK_TRADING_DAYS):
    if len(closes) < 2:
        return None
    returns = [
        (closes[i] - closes[i - 1]) / closes[i - 1]
        for i in range(1, len(closes)) if closes[i - 1]
    ]
    returns = returns[-lookback_trading_days:]
    if len(returns) < 2:
        return None
    return statistics.stdev(returns) * 100


def classify_stocks_by_hv(symbol_to_key, stocks, high_vol_min=HV_THRESHOLD_PCT, mid_vol_min=MIDVOL_HV_MIN):
    print(f"Computing daily historical volatility for {len(stocks)} stocks "
          f"(High-Vol band: HV >= {high_vol_min}%, Mid-Vol band: {mid_vol_min}%-{high_vol_min}%)...")
    high_vol_stocks, mid_vol_stocks = [], []
    prev_close_cache = {}
    for i, stock in enumerate(stocks, 1):
        instrument_key = symbol_to_key.get(stock)
        if not instrument_key:
            continue
        closes = fetch_daily_closes(instrument_key)
        hv = daily_historical_volatility_pct(closes)
        if hv is None:
            continue
        if hv >= high_vol_min:
            high_vol_stocks.append(stock)
            band = "High-Vol"
        elif hv >= mid_vol_min:
            mid_vol_stocks.append(stock)
            band = "Mid-Vol"
        else:
            band = None
        if band and closes:
            prev_close_cache[stock] = closes[-1]
        if i % 25 == 0 or i == len(stocks):
            print(f"  [{i}/{len(stocks)}] HV computed so far...")
    print(f"\n{len(high_vol_stocks)} High-Vol, {len(mid_vol_stocks)} Mid-Vol stocks classified.\n")
    return high_vol_stocks, mid_vol_stocks, prev_close_cache


# ---------------------------------------------------------------------------
# RVOL gate - from the RVOL tracker (cheap, whole-universe, every cycle)
# ---------------------------------------------------------------------------
def fetch_historical_days_candles(instrument_key, lookback_days=RVOL_LOOKBACK_DAYS):
    calendar_days_needed = int(lookback_days * 1.6) + 5
    to_date = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
    from_date = (date.today() - timedelta(days=calendar_days_needed)).strftime("%Y-%m-%d")
    url = f"{HISTORICAL_CANDLE_V3_URL}/{instrument_key}/minutes/1/{to_date}/{from_date}"
    resp = _get_with_retry(url, {})
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


def precompute_rvol_baselines(symbol_to_key, stocks):
    print(f"Precomputing RVOL baselines for {len(stocks)} stocks (once, at startup)...")
    baselines = {}
    for i, stock in enumerate(stocks, 1):
        instrument_key = symbol_to_key.get(stock)
        if not instrument_key:
            continue
        baselines[stock] = fetch_historical_days_candles(instrument_key)
        if i % 25 == 0 or i == len(stocks):
            print(f"  [{i}/{len(stocks)}] RVOL baselines fetched so far...")
    print()
    return baselines


def fetch_batch_volume(instrument_keys):
    results = {}
    for i in range(0, len(instrument_keys), MAX_QUOTE_KEYS_PER_CALL):
        batch = instrument_keys[i:i + MAX_QUOTE_KEYS_PER_CALL]
        resp = _get_with_retry(MARKET_QUOTE_FULL_URL, {"instrument_key": ",".join(batch)})
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
# SuperTrend / Vol Delta / Vol Surge - from the HVROL dashboard, refreshed
# independently of the main RVOL/OI cycle (SUPERTREND_REFRESH_SEC)
# ---------------------------------------------------------------------------
def fetch_intraday_candles(instrument_key, unit=SUPERTREND_CANDLE_UNIT, interval=SUPERTREND_CANDLE_INTERVAL):
    url = f"{INTRADAY_CANDLE_V3_URL}/{instrument_key}/{unit}/{interval}"
    resp = _get_with_retry(url, {})
    if resp is None or resp.status_code != 200:
        return []
    return resp.json().get("data", {}).get("candles", [])


def fetch_recent_candles_for_supertrend(instrument_key, unit=SUPERTREND_CANDLE_UNIT,
                                         interval=SUPERTREND_CANDLE_INTERVAL,
                                         history_days=SUPERTREND_HISTORY_DAYS):
    to_date = date.today().strftime("%Y-%m-%d")
    from_date = (date.today() - timedelta(days=history_days)).strftime("%Y-%m-%d")
    hist_url = f"{HISTORICAL_CANDLE_V3_URL}/{instrument_key}/{unit}/{interval}/{to_date}/{from_date}"
    hist_resp = _get_with_retry(hist_url, {})
    historical = []
    if hist_resp is not None and hist_resp.status_code == 200:
        historical = hist_resp.json().get("data", {}).get("candles", [])
    intraday = fetch_intraday_candles(instrument_key, unit, interval)
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
            true_ranges.append(max(
                highs[i] - lows[i],
                abs(highs[i] - closes[i - 1]),
                abs(lows[i] - closes[i - 1]),
            ))

    atr = [None] * n
    atr[atr_period - 1] = sum(true_ranges[:atr_period]) / atr_period
    for i in range(atr_period, n):
        atr[i] = (atr[i - 1] * (atr_period - 1) + true_ranges[i]) / atr_period

    upper_band = [None] * n
    lower_band = [None] * n
    supertrend = [None] * n
    direction = [None] * n

    start = atr_period - 1
    for i in range(start, n):
        hl2 = (highs[i] + lows[i]) / 2
        basic_upper = hl2 + multiplier * atr[i]
        basic_lower = hl2 - multiplier * atr[i]

        if i == start:
            upper_band[i] = basic_upper
            lower_band[i] = basic_lower
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
    if not avg_volume:
        return None
    return latest_volume / avg_volume


def get_supertrend_vol_delta_surge(stock, instrument_key, cache):
    """Cached, refreshed every SUPERTREND_REFRESH_SEC - none of these can
    change faster than a new 3-minute candle closes anyway. RVOL is NOT
    computed here (unlike the original HVROL dashboard) - it's already
    computed once per cycle for the WHOLE universe via the cheap
    compute_current_rvol() batch path above, and that same value is
    reused for display, avoiding a second, more expensive per-stock RVOL
    calculation."""
    entry = cache.get(stock)
    now = time.time()
    if entry and (now - entry["fetched_at"] < SUPERTREND_REFRESH_SEC):
        return entry["supertrend"], entry["vol_delta"], entry["vol_surge"]
    candles = fetch_recent_candles_for_supertrend(instrument_key)
    _, supertrend_value, _ = compute_supertrend(candles)
    vol_delta_value = compute_latest_volume_delta(candles)
    vol_surge_value = compute_volume_surge(candles)
    cache[stock] = {"fetched_at": now, "supertrend": supertrend_value,
                     "vol_delta": vol_delta_value, "vol_surge": vol_surge_value}
    return supertrend_value, vol_delta_value, vol_surge_value


# ---------------------------------------------------------------------------
# Per-stock full fetch - ITM/OTM aggregates, walls, max-OI-change strikes
# ---------------------------------------------------------------------------
def oi_change_pct(curr_oi, prev_oi):
    if not prev_oi:
        return None
    return ((curr_oi - prev_oi) / prev_oi) * 100


def classify_moneyness(option_type, strike, spot, atm_strike):
    if strike == atm_strike:
        return "ATM"
    if option_type == "CE":
        return "ITM" if strike < spot else "OTM"
    return "ITM" if strike > spot else "OTM"


# ---------------------------------------------------------------------------
# Net GEX / Zero-Gamma Flip
# ---------------------------------------------------------------------------
def bs_gamma(spot, strike, iv_decimal, time_years, r=GEX_RISK_FREE_RATE):
    """Black-Scholes gamma - same value for a call and a put at the same
    strike/IV, so this one function covers both legs. Returns None for any
    input that would make the formula blow up or is nonsensical (zero/None
    IV, zero/negative spot or strike, non-positive time)."""
    if not iv_decimal or iv_decimal <= 0 or not time_years or time_years <= 0 \
            or not spot or spot <= 0 or not strike or strike <= 0:
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
    """chain_snapshot: list of (strike, ce_oi, ce_iv_pct, pe_oi, pe_iv_pct).
    Net GEX = Σ(call gamma x call OI) - Σ(put gamma x put OI), scaled by
    lot size and spot^2 (the standard "dollar gamma per 1% move" scaling) -
    see the sign-convention note above GEX_RISK_FREE_RATE. Returns the raw
    (unscaled-for-display) figure, or None if nothing could be computed."""
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
    """Sweeps Net GEX across every strike present in the chain, treating
    each strike in turn as a HYPOTHETICAL spot price (holding each
    option's own IV fixed - the standard simplifying assumption; a full
    vol-smile-aware model would let IV drift with hypothetical spot too),
    then finds where the sign flips and linearly interpolates between the
    two bracketing strikes for a precise crossing level. If several
    crossings exist, returns the one nearest today's ACTUAL spot (the
    most immediately relevant one for trading purposes). Returns None if
    there's no sign change across the observed strike range."""
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
    """True if this stock currently matches the 'gamma squeeze candidate'
    pattern: negative Net GEX + spot just below (within
    SQUEEZE_APPROACH_PCT_MAX%) the zero-gamma flip + fresh call OI
    building above spot + a volume surge with net buying + price already
    trending up with SuperTrend confirming. A probability tilt based on a
    widely-referenced options-flow pattern, not a guarantee - see the
    SQUEEZE_APPROACH_PCT_MAX comment above for the caveats."""
    net_gex = row.get("net_gex")
    gex_flip = row.get("gex_flip")
    spot = row.get("spot")
    ce_strike = row.get("ce_strike")
    ce_chg = row.get("ce_chg")
    vol_surge = row.get("vol_surge")
    vol_delta = row.get("vol_delta")
    price_chg = row.get("price_chg")
    above_supertrend = row.get("above_supertrend")

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


def process_one_stock(stock, symbol_to_key, expiry_cache, prev_close_cache, supertrend_cache,
                       rvol_value, band_label, lot_size):
    instrument_key = symbol_to_key.get(stock)
    expiry = expiry_cache.get(stock)
    if not instrument_key or not expiry:
        return None

    chain, error_detail = fetch_option_chain(instrument_key, expiry)
    if not chain:
        print(f"  {stock}: option chain fetch failed ({error_detail}), skipping")
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
    ce_best = None  # (score, strike, oi, prev_oi, oi_chg, oi_chg_pct)
    pe_best = None
    chain_snapshot = []  # (strike, ce_oi, ce_iv_pct, pe_oi, pe_iv_pct) - feeds Net GEX / flip below

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
        if not (OI_CHANGE_MODE == "increase" and ce_chg_here <= 0):
            ce_score = abs(ce_chg_here) if OI_CHANGE_MODE == "abs" else ce_chg_here
            if ce_best is None or ce_score > ce_best[0]:
                ce_pct = ((ce_chg_here / ce_prev_here) * 100) if ce_prev_here else None
                ce_best = (ce_score, strike, ce_oi_here, ce_prev_here, ce_chg_here, ce_pct)
        if not (OI_CHANGE_MODE == "increase" and pe_chg_here <= 0):
            pe_score = abs(pe_chg_here) if OI_CHANGE_MODE == "abs" else pe_chg_here
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
        stock, instrument_key, supertrend_cache)
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

    def _fmt(v, suffix="%"):
        return "N/A" if v is None else f"{v:.1f}{suffix}"

    gex_txt = f"{net_gex_cr:.2f}Cr" if net_gex_cr is not None else "N/A"
    flip_txt = f"{gex_flip:.1f}" if gex_flip is not None else "N/A"
    squeeze_txt = " | \U0001F680 SQUEEZE SETUP" if row["squeeze_setup"] else ""
    print(f"  [{band_label}] {stock}: RVOL {rvol_value:.2f} | price {_fmt(price_chg)} | "
          f"CE Wall {ce_wall_strike} | PE Wall {pe_wall_strike} | "
          f"Max{{Delta}}OI CE {ce_strike}({ce_maxchg if ce_maxchg is not None else 'N/A'}) | "
          f"Max{{Delta}}OI PE {pe_strike}({pe_maxchg if pe_maxchg is not None else 'N/A'}) | "
          f"ITM CE {_fmt(ce_itm_chg)} vs PE {_fmt(pe_itm_chg)} | "
          f"Net GEX {gex_txt} | GEX Flip {flip_txt}{squeeze_txt}")

    return row


def run_cycle(qualifying_stocks, symbol_to_key, expiry_cache, prev_close_cache, supertrend_cache,
              rvol_by_stock, band_of_stock, lot_size_map):
    rows = []
    with ThreadPoolExecutor(max_workers=FETCH_WORKERS) as pool:
        futures = {
            pool.submit(process_one_stock, stock, symbol_to_key, expiry_cache, prev_close_cache,
                        supertrend_cache, rvol_by_stock[stock], band_of_stock.get(stock, "?"),
                        lot_size_map.get(stock, GEX_DEFAULT_LOT_SIZE)): stock
            for stock in qualifying_stocks
        }
        for future in as_completed(futures):
            row = future.result()
            if row:
                rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Excel (COM) - TWO flat, growing, self-sorting, flicker-free live sheets:
# MAIN (the frequently-watched columns) and DETAILS (the bulkier raw OI /
# expiry columns, moved off the main sheet to keep it scannable). Both
# sheets are kept in the SAME row order every cycle (grouped/sorted by
# _order_by_rvol_then_time), so row N on one sheet is always the same
# (Last Updated, Stock) as row N on the other.
# ---------------------------------------------------------------------------
MAIN_HEADERS_ROW = [
    "Last Updated", "Vol Band", "Stock", "Spot", "Price Chg %",
    "RVOL", "Vol Surge", "Vol Delta",
    "CE Wall (Max OI)", "PE Wall (Max OI)",
    "CE IV", "PE IV",
    "ITM CE OI Chg %", "OTM CE OI Chg %", "ITM PE OI Chg %", "OTM PE OI Chg %",
    "CE Strike (Max OI Chg)", "CE OI Chg", "CE OI Chg %",
    "PE OI Chg %", "PE OI Chg", "PE Strike (Max OI Chg)",
    "SuperTrend", "Net GEX (Cr)", "GEX Flip", "Squeeze Setup",
]
MAIN_COL_WIDTHS = [
    18, 10, 14, 10, 10,
    8, 10, 10,
    16, 16,
    9, 9,
    12, 12, 12, 12,
    20, 10, 10,
    10, 10, 20,
    11, 12, 10, 14,
]

DETAILS_HEADERS_ROW = [
    "Last Updated", "Stock",
    "CE Wall OI", "PE Wall OI",
    "CE OI", "CE Prev OI", "PE OI", "PE Prev OI",
    "ITM CE OI", "ITM CE Prev OI", "ITM PE OI", "ITM PE Prev OI",
    "OTM CE OI", "OTM CE Prev OI", "OTM PE OI", "OTM PE Prev OI",
    "Expiry",
]
DETAILS_COL_WIDTHS = [18, 14, 12, 12, 10, 10, 10, 10, 10, 12, 10, 12, 10, 12, 10, 12, 12]

MAIN_HEADER_TO_KEY = dict(zip(MAIN_HEADERS_ROW, [
    "last_updated", "band", "stock", "spot", "price_chg",
    "rvol", "vol_surge", "vol_delta",
    "ce_wall", "pe_wall",
    "ce_iv", "pe_iv",
    "itm_ce_chg", "otm_ce_chg", "itm_pe_chg", "otm_pe_chg",
    "ce_strike", "ce_chg", "ce_chg_pct",
    "pe_chg_pct", "pe_chg", "pe_strike",
    "supertrend", "net_gex", "gex_flip", "squeeze_setup",
]))

DETAILS_HEADER_TO_KEY = dict(zip(DETAILS_HEADERS_ROW, [
    "last_updated", "stock",
    "ce_wall_oi", "pe_wall_oi",
    "ce_oi", "ce_prev_oi", "pe_oi", "pe_prev_oi",
    "itm_ce_oi", "itm_ce_prev_oi", "itm_pe_oi", "itm_pe_prev_oi",
    "otm_ce_oi", "otm_ce_prev_oi", "otm_pe_oi", "otm_pe_prev_oi",
    "expiry",
]))

# Column numbers (1-based, on the MAIN sheet) of interest for formatting -
# kept as named constants so a future column reorder only needs updating
# here.
COL_ITM_CE_CHG, COL_OTM_CE_CHG, COL_ITM_PE_CHG, COL_OTM_PE_CHG = 13, 14, 15, 16
COL_CE_OI_CHG = 18
COL_PE_OI_CHG = 21
COL_NET_GEX = 24
COL_SQUEEZE_SETUP = 26


def get_excel_app():
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_UP = -4162
XL_TO_LEFT = -4159


class _frozen_excel:
    """Suspends ScreenUpdating/Calculation/EnableEvents for a block of COM
    writes so Excel repaints once at the end instead of after every single
    Range.Value/NumberFormat/Interior.Color call - eliminates the visible
    flicker of a live-updating sheet. Fails safe: a COM error while
    toggling these never blocks the actual write, only skips the
    flicker-suppression for that one cycle."""
    def __init__(self, excel):
        self.excel = excel
        self._prev_screen_updating = True
        self._prev_calculation = XL_CALCULATION_AUTOMATIC
        self._prev_events = True

    def __enter__(self):
        try:
            self._prev_screen_updating = self.excel.ScreenUpdating
            self._prev_calculation = self.excel.Calculation
            self._prev_events = self.excel.EnableEvents
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False
            self.excel.Calculation = XL_CALCULATION_MANUAL
        except pywintypes.com_error as e:
            print(f"[warn] Could not suspend Excel screen updating this cycle ({e}).")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            self.excel.Calculation = self._prev_calculation
            if self._prev_calculation == XL_CALCULATION_AUTOMATIC:
                self.excel.Calculate()
            self.excel.ScreenUpdating = self._prev_screen_updating
            self.excel.EnableEvents = self._prev_events
        except pywintypes.com_error as e:
            print(f"[warn] Could not restore Excel screen updating after this cycle's write ({e}).")
        return False


def _com_retry(func, max_retries=8, initial_delay=0.3, max_delay=10.0):
    last_exc = None
    delay = initial_delay
    for _ in range(max_retries):
        try:
            return func()
        except pywintypes.com_error as e:
            last_exc = e
            time.sleep(delay)
            delay = min(delay * 2, max_delay)
    raise last_exc


def get_or_open_workbook_com(excel, filepath):
    filename = os.path.basename(filepath)
    for wb in excel.Workbooks:
        if wb.Name == filename:
            return wb
    if os.path.exists(filepath):
        return excel.Workbooks.Open(filepath)
    wb = excel.Workbooks.Add()
    while wb.Sheets.Count > 1:
        wb.Sheets(wb.Sheets.Count).Delete()
    wb.SaveAs(filepath)
    return wb


def get_or_add_sheet(wb, name):
    for sht in wb.Sheets:
        if sht.Name == name:
            return sht
    if wb.Sheets.Count == 1 and wb.Sheets(1).Cells(1, 1).Value is None \
            and wb.Sheets(1).Name.startswith("Sheet"):
        sht = wb.Sheets(1)
        sht.Name = name
        return sht
    sht = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    sht.Name = name
    return sht


def ensure_header(sht, headers_row, col_widths):
    if sht.Cells(1, 1).Value == headers_row[0]:
        return
    header_range = sht.Range(sht.Cells(1, 1), sht.Cells(1, len(headers_row)))
    header_range.Value = [headers_row]
    header_range.Font.Bold = True
    header_range.Interior.Color = _hex_to_bgr("D9E1F2")
    for i, width in enumerate(col_widths, 1):
        sht.Columns(i).ColumnWidth = width
    try:
        header_range.AutoFilter()
    except pywintypes.com_error:
        pass
    try:
        sht.Application.ActiveWindow.SplitRow = 1
        sht.Application.ActiveWindow.FreezePanes = True
    except Exception:
        pass


def _hex_to_bgr(hex_color):
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (b << 16) + (g << 8) + r


COLOR_OI_UP = _hex_to_bgr("1E7B34")    # green - build-up
COLOR_OI_DOWN = _hex_to_bgr("C00000")  # red - unwind
COLOR_SQUEEZE = _hex_to_bgr("FF8C00")  # dark orange - squeeze setup flag


def _apply_signed_coloring(sht, first_row, last_row, col):
    """Box (cell background) coloring: green fill if positive, red fill
    if negative."""
    for row in range(first_row, last_row + 1):
        cell = sht.Cells(row, col)
        try:
            val = cell.Value
            if isinstance(val, (int, float)):
                cell.Interior.Color = COLOR_OI_UP if val > 0 else COLOR_OI_DOWN
            else:
                cell.Interior.ColorIndex = -4142  # xlNone
        except pywintypes.com_error:
            continue


def _apply_signed_font_coloring(sht, first_row, last_row, col):
    """Font-only coloring: bold, green text if positive, bold red text if
    negative - no cell background fill (explicitly cleared in case a
    previous cycle's box coloring is still sitting there)."""
    for row in range(first_row, last_row + 1):
        cell = sht.Cells(row, col)
        try:
            cell.Interior.ColorIndex = -4142  # xlNone - no box/background fill
            val = cell.Value
            if isinstance(val, (int, float)):
                cell.Font.Bold = True
                cell.Font.Color = COLOR_OI_UP if val > 0 else COLOR_OI_DOWN
            else:
                cell.Font.Bold = False
                cell.Font.ColorIndex = -4105  # xlAutomatic
        except pywintypes.com_error:
            continue


def _apply_squeeze_coloring(sht, first_row, last_row, col):
    """Bold orange text for rows flagged 'YES' by evaluate_squeeze_setup;
    plain/unstyled otherwise. No box fill, matching the font-only style
    used for the other flag/derived columns."""
    for row in range(first_row, last_row + 1):
        cell = sht.Cells(row, col)
        try:
            cell.Interior.ColorIndex = -4142  # xlNone
            if cell.Value == "YES":
                cell.Font.Bold = True
                cell.Font.Color = COLOR_SQUEEZE
            else:
                cell.Font.Bold = False
                cell.Font.ColorIndex = -4105  # xlAutomatic
        except pywintypes.com_error:
            continue


def _normalize_last_updated(v):
    if isinstance(v, float) and 0 <= v <= 1:
        total_seconds = int(v * 86400)
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    elif v is not None:
        return str(v)
    return v


def _read_sheet_rows_generic(sht, header_to_key, header_row=1):
    """Reads a sheet's OWN current header row (whatever columns it holds)
    to build a {internal_key: column} map, then reads every existing data
    row using THAT map - works for either the MAIN or the DETAILS sheet,
    whatever width each currently has, via 2 bulk Range.Value reads."""
    last_header_col = sht.Cells(header_row, sht.Columns.Count).End(XL_TO_LEFT).Column
    if last_header_col < 1:
        return []
    header_range = sht.Range(sht.Cells(header_row, 1), sht.Cells(header_row, last_header_col))
    header_vals = header_range.Value
    if header_vals and isinstance(header_vals[0], (list, tuple)):
        header_vals = header_vals[0]

    header_map = {}
    for idx, h in enumerate(header_vals, start=1):
        if h:
            key = header_to_key.get(h)
            if key:
                header_map[key] = idx

    stock_col = header_map.get("stock")
    if stock_col is None:
        return []

    last_row = sht.Cells(sht.Rows.Count, 1).End(XL_UP).Row
    if last_row < header_row + 1:
        return []

    max_col = max(header_map.values())
    data_range = sht.Range(sht.Cells(header_row + 1, 1), sht.Cells(last_row, max_col))
    data_rows = data_range.Value
    if last_row == header_row + 1 and data_rows and not isinstance(data_rows[0], (list, tuple)):
        data_rows = (data_rows,)

    results = []
    for row_tuple in data_rows:
        if row_tuple[stock_col - 1] is None:
            continue
        d = {key: row_tuple[c - 1] for key, c in header_map.items()}
        if "last_updated" in d:
            d["last_updated"] = _normalize_last_updated(d["last_updated"])
        results.append(d)
    return results


def _order_by_rvol_then_time(row_dicts):
    """Groups rows by stock. Stock GROUPS are ordered by that stock's
    single highest RVOL seen across ANY of its rows so far today,
    descending. WITHIN each stock's own group, rows are ordered
    chronologically ascending (earliest time on top)."""
    grouped = {}
    for d in row_dicts:
        grouped.setdefault(d.get("stock"), []).append(d)

    def _stock_max_rvol(stock_rows):
        best = 0
        for d in stock_rows:
            v = d.get("rvol")
            if isinstance(v, (int, float)):
                best = max(best, v)
        return best

    ordered_stocks = sorted(grouped.keys(), key=lambda s: _stock_max_rvol(grouped[s]), reverse=True)

    combined = []
    for stock in ordered_stocks:
        stock_rows = grouped[stock]
        stock_rows.sort(key=lambda d: str(d.get("last_updated") or ""))
        combined.extend(stock_rows)
    return combined


def _disp(v):
    return "N/A" if v is None else v


def com_resort_sheets(main_sht, details_sht, rows, refreshed_at):
    """Appends this cycle's rows to whatever's already on EITHER sheet
    (read back by header name from both, then merged by (stock, last
    updated) key so a row's full field set survives even though it's
    split across two sheets), then rewrites BOTH sheets' WHOLE data areas
    re-sorted the SAME way: stock groups ordered by that stock's highest
    RVOL today (descending), rows WITHIN each stock's group ordered
    chronologically (oldest first) - a growing, self-sorting history on
    each sheet, with row N always the same (Last Updated, Stock) on both."""
    ts = refreshed_at.strftime("%Y-%m-%d %H:%M:%S")

    existing_main = _read_sheet_rows_generic(main_sht, MAIN_HEADER_TO_KEY)
    existing_details = _read_sheet_rows_generic(details_sht, DETAILS_HEADER_TO_KEY)
    merged = {}
    for d in existing_main + existing_details:
        key = (d.get("stock"), d.get("last_updated"))
        merged.setdefault(key, {}).update(d)
    existing_combined = list(merged.values())

    new_rows = []
    for r in rows:
        d = dict(r)
        d["last_updated"] = ts
        d["rvol"] = round(r["rvol"], 2)
        for k in ("price_chg", "ce_iv", "pe_iv", "itm_ce_chg", "itm_pe_chg", "otm_ce_chg",
                  "otm_pe_chg", "ce_chg_pct", "pe_chg_pct"):
            if d.get(k) is not None:
                d[k] = round(d[k], 1)
        if d.get("supertrend") is not None:
            d["supertrend"] = round(d["supertrend"], 1)
        if d.get("vol_surge") is not None:
            d["vol_surge"] = round(d["vol_surge"], 2)
        if d.get("net_gex") is not None:
            d["net_gex"] = round(d["net_gex"], 2)
        if d.get("gex_flip") is not None:
            d["gex_flip"] = round(d["gex_flip"], 1)
        d["squeeze_setup"] = "YES" if d.get("squeeze_setup") else ""
        new_rows.append(d)

    combined = _order_by_rvol_then_time(existing_combined + new_rows)

    # --- MAIN sheet ---
    main_last_col = len(MAIN_HEADERS_ROW)
    used_last_row = main_sht.Cells(main_sht.Rows.Count, 1).End(XL_UP).Row
    if used_last_row > 1:
        _com_retry(lambda: main_sht.Range(main_sht.Cells(2, 1), main_sht.Cells(used_last_row, main_last_col)).ClearContents())

    main_block = []
    for d in combined:
        main_block.append([
            d.get("last_updated"), d.get("band"), d.get("stock"), d.get("spot"), _disp(d.get("price_chg")),
            d.get("rvol"), _disp(d.get("vol_surge")), d.get("vol_delta"),
            _disp(d.get("ce_wall")), _disp(d.get("pe_wall")),
            _disp(d.get("ce_iv")), _disp(d.get("pe_iv")),
            _disp(d.get("itm_ce_chg")), _disp(d.get("otm_ce_chg")),
            _disp(d.get("itm_pe_chg")), _disp(d.get("otm_pe_chg")),
            _disp(d.get("ce_strike")), d.get("ce_chg"), _disp(d.get("ce_chg_pct")),
            _disp(d.get("pe_chg_pct")), d.get("pe_chg"), _disp(d.get("pe_strike")),
            _disp(d.get("supertrend")), _disp(d.get("net_gex")), _disp(d.get("gex_flip")),
            d.get("squeeze_setup", ""),
        ])

    if main_block:
        first_row, last_row = 2, 1 + len(main_block)
        _com_retry(lambda: setattr(main_sht.Range(main_sht.Cells(first_row, 1), main_sht.Cells(last_row, 1)),
                                    "NumberFormat", "@"))
        target = main_sht.Range(main_sht.Cells(first_row, 1), main_sht.Cells(last_row, main_last_col))
        _com_retry(lambda: setattr(target, "Value", main_block))

        # ITM/OTM OI Chg % columns (grouped together): bold font color
        # only (green positive / red negative) - no cell background fill
        for col in (COL_ITM_CE_CHG, COL_OTM_CE_CHG, COL_ITM_PE_CHG, COL_OTM_PE_CHG):
            _apply_signed_font_coloring(main_sht, first_row, last_row, col)

        # CE OI Chg (max-OI-change column, CE side): fixed red highlight
        ce_chg_range = main_sht.Range(main_sht.Cells(first_row, COL_CE_OI_CHG), main_sht.Cells(last_row, COL_CE_OI_CHG))
        _com_retry(lambda: setattr(ce_chg_range, "NumberFormat", "0"))
        _com_retry(lambda: setattr(ce_chg_range.Interior, "Color", COLOR_OI_DOWN))

        # PE OI Chg (max-OI-change column, PE side): fixed GREEN fill,
        # no conditional sign-based coloring - mirrors CE OI Chg's fixed
        # red fill on the other side. Font explicitly reset to plain/
        # automatic (no bold, no color override) in case a previous
        # cycle's font styling is still sitting on these cells.
        pe_chg_range = main_sht.Range(main_sht.Cells(first_row, COL_PE_OI_CHG), main_sht.Cells(last_row, COL_PE_OI_CHG))
        _com_retry(lambda: setattr(pe_chg_range, "NumberFormat", "0"))
        _com_retry(lambda: setattr(pe_chg_range.Interior, "Color", COLOR_OI_UP))
        _com_retry(lambda: setattr(pe_chg_range.Font, "Bold", False))
        _com_retry(lambda: setattr(pe_chg_range.Font, "ColorIndex", -4105))

        # Net GEX: bold font color only (green positive / red negative) -
        # sign matters far more than magnitude here, so this mirrors the
        # ITM/OTM %-chg columns' styling rather than the box-fill columns.
        _apply_signed_font_coloring(main_sht, first_row, last_row, COL_NET_GEX)

        # Squeeze Setup: bold orange text when flagged "YES"
        _apply_squeeze_coloring(main_sht, first_row, last_row, COL_SQUEEZE_SETUP)

    # --- DETAILS sheet ---
    details_last_col = len(DETAILS_HEADERS_ROW)
    used_last_row_d = details_sht.Cells(details_sht.Rows.Count, 1).End(XL_UP).Row
    if used_last_row_d > 1:
        _com_retry(lambda: details_sht.Range(details_sht.Cells(2, 1),
                                              details_sht.Cells(used_last_row_d, details_last_col)).ClearContents())

    details_block = []
    for d in combined:
        details_block.append([
            d.get("last_updated"), d.get("stock"),
            d.get("ce_wall_oi"), d.get("pe_wall_oi"),
            d.get("ce_oi"), d.get("ce_prev_oi"), d.get("pe_oi"), d.get("pe_prev_oi"),
            d.get("itm_ce_oi"), d.get("itm_ce_prev_oi"),
            d.get("itm_pe_oi"), d.get("itm_pe_prev_oi"),
            d.get("otm_ce_oi"), d.get("otm_ce_prev_oi"),
            d.get("otm_pe_oi"), d.get("otm_pe_prev_oi"),
            d.get("expiry"),
        ])

    if details_block:
        first_row, last_row = 2, 1 + len(details_block)
        _com_retry(lambda: setattr(details_sht.Range(details_sht.Cells(first_row, 1), details_sht.Cells(last_row, 1)),
                                    "NumberFormat", "@"))
        target = details_sht.Range(details_sht.Cells(first_row, 1), details_sht.Cells(last_row, details_last_col))
        _com_retry(lambda: setattr(target, "Value", details_block))


def backup_workbook_copy(wb, label=BACKUP_LABEL):
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_path = os.path.join(BACKUP_DIR, f"{label}_{ts}.xlsx")
    try:
        wb.SaveCopyAs(backup_path)
        _prune_old_backups(label)
        return backup_path
    except Exception as e:
        print(f"[warn] Backup failed for {label}: {e}")
        return None


def _prune_old_backups(label=BACKUP_LABEL):
    if BACKUP_RETENTION_COUNT is None:
        return
    try:
        matches = sorted(
            (f for f in os.listdir(BACKUP_DIR)
             if f.startswith(label + "_") and f.endswith(".xlsx") and "_EOD_" not in f),
            key=lambda f: os.path.getmtime(os.path.join(BACKUP_DIR, f)),
        )
        excess = len(matches) - BACKUP_RETENTION_COUNT
        for old_file in matches[:max(0, excess)]:
            os.remove(os.path.join(BACKUP_DIR, old_file))
    except Exception as e:
        print(f"[warn] Backup pruning skipped: {e}")


def archive_end_of_day_and_cleanup(wb, label=BACKUP_LABEL):
    today_str = date.today().strftime("%Y-%m-%d")
    print("\nMarket closed - saving end-of-day archive and cleaning up intraday backups...")
    eod_path = os.path.join(BACKUP_DIR, f"{label}_EOD_{today_str}.xlsx")
    try:
        wb.SaveCopyAs(eod_path)
        print(f"  Saved end-of-day archive: {eod_path}")
    except Exception as e:
        print(f"  [warn] Could not save end-of-day archive: {e}")
        print("  Keeping all intraday backups since the end-of-day archive failed to save.")
        return
    deleted = 0
    for fname in os.listdir(BACKUP_DIR):
        if fname.startswith(label + "_") and fname.endswith(".xlsx") and "_EOD_" not in fname:
            try:
                os.remove(os.path.join(BACKUP_DIR, fname))
                deleted += 1
            except Exception as e:
                print(f"  [warn] Could not delete {fname}: {e}")
    print(f"  Deleted {deleted} intraday backup file(s). End-of-day archive kept.")


def market_is_open(now=None):
    now = now or datetime.now()
    if now.weekday() >= 5:
        return False
    return MARKET_OPEN <= now.time() <= MARKET_CLOSE


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    symbol_to_key, instruments = build_symbol_to_instrument_key()
    candidate_stocks = build_active_fno_stock_list(symbol_to_key, instruments)

    high_vol_stocks, mid_vol_stocks, prev_close_cache = classify_stocks_by_hv(symbol_to_key, candidate_stocks)
    if not high_vol_stocks and not mid_vol_stocks:
        print("No stocks fell in either HV band - nothing to track. Exiting.")
        return
    all_tracked_stocks = high_vol_stocks + mid_vol_stocks
    band_of_stock = {s: "High-Vol" for s in high_vol_stocks}
    band_of_stock.update({s: "Mid-Vol" for s in mid_vol_stocks})

    lot_size_map = build_lot_size_map(symbol_to_key, instruments)
    missing_lot_size = [s for s in all_tracked_stocks if s not in lot_size_map]
    if missing_lot_size:
        print(f"[warn] No lot size found for {len(missing_lot_size)} stock(s) - Net GEX/GEX Flip "
              f"for those will use a fallback lot size of {GEX_DEFAULT_LOT_SIZE} (likely inaccurate).")

    rvol_baselines = precompute_rvol_baselines(symbol_to_key, all_tracked_stocks)

    print(f"Resolving nearest expiry for {len(all_tracked_stocks)} stocks (once)...")
    expiry_cache = {}
    for i, stock in enumerate(all_tracked_stocks, 1):
        instrument_key = symbol_to_key.get(stock)
        if not instrument_key:
            continue
        expiry = get_nearest_expiry(instrument_key)
        if expiry:
            expiry_cache[stock] = expiry
        time.sleep(0.1)
    print(f"Resolved expiry for {len(expiry_cache)}/{len(all_tracked_stocks)} stocks.\n")

    all_instrument_keys = [symbol_to_key[s] for s in all_tracked_stocks if s in symbol_to_key]
    supertrend_cache = {}

    print("Launching Excel (this window will stay open and update live)...")
    excel = get_excel_app()
    wb = get_or_open_workbook_com(excel, OUT_FILE)
    main_sht = get_or_add_sheet(wb, MAIN_SHEET_NAME)
    details_sht = get_or_add_sheet(wb, DETAILS_SHEET_NAME)
    with _frozen_excel(excel):
        ensure_header(main_sht, MAIN_HEADERS_ROW, MAIN_COL_WIDTHS)
        ensure_header(details_sht, DETAILS_HEADERS_ROW, DETAILS_COL_WIDTHS)
    wb.Save()

    print(f"OI ULTRA LIVE tracking {len(high_vol_stocks)} High-Vol + {len(mid_vol_stocks)} Mid-Vol "
          f"stocks, gated by RVOL >= {RVOL_MIN_THRESHOLD}, refreshing every {REFRESH_INTERVAL_SEC}s "
          f"-> {OUT_FILE}")
    if MARKET_HOURS_ONLY:
        print(f"Restricted to NSE market hours ({MARKET_OPEN}-{MARKET_CLOSE} IST, Mon-Fri).")
    print("Leave the Excel window open (don't close it) for live updates to keep working.")
    print("Press Ctrl+C to stop.\n")

    try:
        while True:
            now = datetime.now()
            if MARKET_HOURS_ONLY:
                if now.weekday() >= 5 or now.time() > MARKET_CLOSE:
                    print(f"[{now.strftime('%H:%M:%S')}] Market session over for today - exiting.")
                    archive_end_of_day_and_cleanup(wb)
                    break
                if now.time() < MARKET_OPEN:
                    print(f"[{now.strftime('%H:%M:%S')}] Waiting for market to open at {MARKET_OPEN}...")
                    time.sleep(30)
                    continue

            cycle_start = time.time()
            print(f"\n=== Refresh cycle started at {now.strftime('%Y-%m-%d %H:%M:%S')} ===")

            volume_data = fetch_batch_volume(all_instrument_keys)
            rvol_by_stock = compute_current_rvol(symbol_to_key, all_tracked_stocks, rvol_baselines, volume_data)
            qualifying_stocks = [s for s, v in rvol_by_stock.items() if v >= RVOL_MIN_THRESHOLD]
            if MAX_FETCH_PER_CYCLE and len(qualifying_stocks) > MAX_FETCH_PER_CYCLE:
                qualifying_stocks.sort(key=lambda s: rvol_by_stock[s], reverse=True)
                qualifying_stocks = qualifying_stocks[:MAX_FETCH_PER_CYCLE]
            print(f"RVOL >= {RVOL_MIN_THRESHOLD}: {len(qualifying_stocks)}/{len(all_tracked_stocks)} "
                  f"stocks qualify this cycle (fetch cap {MAX_FETCH_PER_CYCLE}).")

            rows = run_cycle(qualifying_stocks, symbol_to_key, expiry_cache, prev_close_cache,
                              supertrend_cache, rvol_by_stock, band_of_stock, lot_size_map)

            try:
                _ = wb.FullName
            except pywintypes.com_error:
                print("[warn] Excel or the workbook was closed - relaunching/reopening...")
                try:
                    excel = get_excel_app()
                except pywintypes.com_error:
                    excel = win32.gencache.EnsureDispatch("Excel.Application")
                    excel.Visible = True
                    excel.DisplayAlerts = False
                wb = get_or_open_workbook_com(excel, OUT_FILE)
                main_sht = get_or_add_sheet(wb, MAIN_SHEET_NAME)
                details_sht = get_or_add_sheet(wb, DETAILS_SHEET_NAME)
                with _frozen_excel(excel):
                    ensure_header(main_sht, MAIN_HEADERS_ROW, MAIN_COL_WIDTHS)
                    ensure_header(details_sht, DETAILS_HEADERS_ROW, DETAILS_COL_WIDTHS)

            try:
                with _frozen_excel(excel):
                    com_resort_sheets(main_sht, details_sht, rows, now)
            except pywintypes.com_error as e:
                print(f"[warn] Excel write failed this cycle ({e}). Skipping; will try again next cycle.")
                time.sleep(max(1, REFRESH_INTERVAL_SEC - (time.time() - cycle_start)))
                continue

            try:
                wb.Save()
            except Exception as e:
                print(f"[warn] Save deferred (Excel busy/editing mode): {e}")

            backup_path = backup_workbook_copy(wb)
            if backup_path:
                print(f"Backup snapshot saved to {BACKUP_DIR}\\")

            elapsed = time.time() - cycle_start
            print(f"{len(rows)} qualifying stock row(s) written this cycle. (cycle took {elapsed:.1f}s)")
            time.sleep(max(1, REFRESH_INTERVAL_SEC - elapsed))
    except KeyboardInterrupt:
        print("\nStopped by user. (Excel window is left open - close it manually if done.)")


if __name__ == "__main__":
    main()
