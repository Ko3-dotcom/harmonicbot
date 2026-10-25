"""
HarmonicBot Pro — Backend Server
Local:  python backend.py
Cloud:  gunicorn backend:app --bind 0.0.0.0:$PORT
Deps:   pip install flask flask-cors requests numpy scipy openpyxl gunicorn
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import hmac, hashlib, time, requests, math, io, datetime, os
import numpy as np
from scipy.signal import argrelextrema
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})

BINANCE_BASE = "https://api.binance.com"
TESTNET_BASE = "https://testnet.binance.vision"

# ── Password protection for cloud deployment ──────────────────────────────
# Set BOT_PASSWORD env variable on Railway to protect your bot
# Leave empty to disable (for local use)
BOT_PASSWORD = os.environ.get("BOT_PASSWORD", "")

@app.before_request
def check_password():
    """Block unauthorized requests when BOT_PASSWORD is set."""
    if not BOT_PASSWORD:
        return  # no password set → open access (local use)
    if request.method == "OPTIONS":
        return  # allow CORS preflight
    if request.path == "/api/ping":
        return  # allow health check
    pwd = request.headers.get("X-Bot-Password","") or \
          (request.json or {}).get("bot_password","")
    if pwd != BOT_PASSWORD:
        return jsonify({"ok": False, "error": "Unauthorized"}), 401

@app.route("/api/ping", methods=["GET"])
def ping():
    """Health check endpoint for Railway."""
    return jsonify({"ok": True, "status": "HarmonicBot running",
                    "time": datetime.datetime.now().isoformat()})

# ─── SIGNING ───────────────────────────────────────────────────────────────

def sign(secret, params):
    query = "&".join(f"{k}={v}" for k, v in params.items())
    return hmac.new(secret.encode(), query.encode(), hashlib.sha256).hexdigest()

def binance_signed(api_key, api_secret, method, path, params=None, testnet=False):
    base = TESTNET_BASE if testnet else BINANCE_BASE
    if params is None: params = {}
    params["timestamp"]  = int(time.time() * 1000)
    params["recvWindow"] = 5000
    query     = "&".join(f"{k}={v}" for k, v in params.items())
    signature = hmac.new(api_secret.encode(), query.encode(), hashlib.sha256).hexdigest()
    full_url  = f"{base}{path}?{query}&signature={signature}"
    headers   = {"X-MBX-APIKEY": api_key}
    if   method == "GET":    r = requests.get(full_url,    headers=headers, timeout=10)
    elif method == "POST":   r = requests.post(full_url,   headers=headers, timeout=10)
    elif method == "DELETE": r = requests.delete(full_url, headers=headers, timeout=10)
    r.raise_for_status()
    return r.json()

# ─── MARKET LIST ───────────────────────────────────────────────────────────

@app.route("/api/top_coins", methods=["GET"])
def top_coins():
    n       = int(request.args.get("n", 10))
    testnet = request.args.get("testnet","false") == "true"
    base    = TESTNET_BASE if testnet else BINANCE_BASE
    try:
        r = requests.get(f"{base}/api/v3/ticker/24hr", timeout=15)
        tickers = r.json()
        usdt = [t for t in tickers if t["symbol"].endswith("USDT") and float(t["quoteVolume"]) > 0]
        usdt.sort(key=lambda x: float(x["quoteVolume"]), reverse=True)
        result = [{"symbol": t["symbol"],
                   "price":  float(t["lastPrice"]),
                   "change": float(t["priceChangePercent"]),
                   "volume": float(t["quoteVolume"])} for t in usdt[:n]]
        return jsonify({"ok": True, "coins": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

@app.route("/api/all_coins", methods=["GET"])
def all_coins():
    testnet = request.args.get("testnet","false") == "true"
    base    = TESTNET_BASE if testnet else BINANCE_BASE
    try:
        r = requests.get(f"{base}/api/v3/ticker/24hr", timeout=15)
        tickers = r.json()
        usdt = [{"symbol": t["symbol"], "price": float(t["lastPrice"]),
                 "volume": float(t["quoteVolume"])}
                for t in tickers if t["symbol"].endswith("USDT") and float(t["quoteVolume"]) > 1000]
        usdt.sort(key=lambda x: x["volume"], reverse=True)
        return jsonify({"ok": True, "coins": usdt})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

# Known bStock symbols on Binance (tokenized stocks)
BSTOCKS_KNOWN = [
    {"symbol":"NVDABUSDT",  "name":"NVIDIA",       "sector":"Tech"},
    {"symbol":"TSLABUSDT",  "name":"Tesla",         "sector":"EV"},
    {"symbol":"CRCLBUSDT",  "name":"Circle",        "sector":"Fintech"},
    {"symbol":"MUBUSDT",    "name":"Micron",        "sector":"Semiconductors"},
    {"symbol":"SNDKBUSDT",  "name":"Sandisk",       "sector":"Storage"},
    {"symbol":"AAPLBUSDT",  "name":"Apple",         "sector":"Tech"},
    {"symbol":"MSFTBUSDT",  "name":"Microsoft",     "sector":"Tech"},
    {"symbol":"GOOGBUSDT",  "name":"Alphabet",      "sector":"Tech"},
    {"symbol":"AMZNBUSDT",  "name":"Amazon",        "sector":"Commerce"},
    {"symbol":"METABUSDT",  "name":"Meta",          "sector":"Social"},
    {"symbol":"COINBUSDT",  "name":"Coinbase",      "sector":"Crypto"},
    {"symbol":"SPYBUSDT",   "name":"S&P500 ETF",    "sector":"Index"},
    {"symbol":"QQBBUSDT",   "name":"Nasdaq ETF",    "sector":"Index"},
    {"symbol":"BRKBBUSDT",  "name":"Berkshire B",   "sector":"Finance"},
    {"symbol":"JPMBBUSDT",  "name":"JPMorgan",      "sector":"Finance"},
]

@app.route("/api/bstocks", methods=["GET"])
def get_bstocks():
    """
    Return available bStock symbols with live prices from Binance.
    Falls back gracefully if symbol not listed on user's account.
    """
    testnet = request.args.get("testnet","false") == "true"
    base    = TESTNET_BASE if testnet else BINANCE_BASE
    try:
        r       = requests.get(f"{base}/api/v3/ticker/24hr", timeout=15)
        tickers = {t["symbol"]: t for t in r.json()}
        result  = []
        for s in BSTOCKS_KNOWN:
            sym = s["symbol"]
            if sym in tickers:
                t = tickers[sym]
                result.append({
                    "symbol":  sym,
                    "name":    s["name"],
                    "sector":  s["sector"],
                    "price":   float(t["lastPrice"]),
                    "change":  float(t["priceChangePercent"]),
                    "volume":  float(t["quoteVolume"]),
                    "type":    "bstock",
                })
        # Known crypto that happen to end in BUSDT pattern — exclude them
        CRYPTO_EXCEPTIONS = {"BNBUSDT","BTCUSDT","ETHUSDT","LTCUSDT","XRPUSDT"}
        # Auto-detect any other BUSDT tokens not in our known list
        # Must have 4+ chars before BUSDT (e.g. NVDABUSDT=9 chars, BNBUSDT=7 but excluded)
        known_syms = {s["symbol"] for s in BSTOCKS_KNOWN}
        for sym, t in tickers.items():
            if sym in CRYPTO_EXCEPTIONS: continue
            if sym in known_syms: continue
            if sym.endswith("BUSDT") and len(sym) >= 8 and float(t["quoteVolume"]) > 0:
                result.append({
                    "symbol": sym,
                    "name":   sym.replace("BUSDT",""),
                    "sector": "TradFi",
                    "price":  float(t["lastPrice"]),
                    "change": float(t["priceChangePercent"]),
                    "volume": float(t["quoteVolume"]),
                    "type":   "bstock",
                })
        result.sort(key=lambda x: x["volume"], reverse=True)
        return jsonify({"ok": True, "bstocks": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

# ─── ACCOUNT ───────────────────────────────────────────────────────────────

@app.route("/api/account", methods=["POST"])
def get_account():
    d = request.json
    try:
        data     = binance_signed(d["api_key"], d["api_secret"], "GET", "/api/v3/account", testnet=d.get("testnet",False))
        balances = {b["asset"]: float(b["free"]) for b in data["balances"] if float(b["free"]) > 0}
        return jsonify({"ok": True, "balances": balances})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

@app.route("/api/open_orders", methods=["POST"])
def open_orders():
    d = request.json
    try:
        data = binance_signed(d["api_key"], d["api_secret"], "GET", "/api/v3/openOrders",
                              {}, testnet=d.get("testnet",False))
        return jsonify({"ok": True, "orders": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

@app.route("/api/trade_history", methods=["POST"])
def trade_history():
    d       = request.json
    testnet = d.get("testnet", False)
    symbols = d.get("symbols") or ["BTCUSDT","ETHUSDT","BNBUSDT","SOLUSDT","XRPUSDT","ADAUSDT"]
    all_trades = []
    for sym in symbols[:10]:
        try:
            trades = binance_signed(d["api_key"], d["api_secret"], "GET", "/api/v3/myTrades",
                                    {"symbol": sym, "limit": 20}, testnet=testnet)
            for t in trades: t["symbol"] = sym
            all_trades.extend(trades)
            time.sleep(0.1)
        except: pass
    all_trades.sort(key=lambda x: x.get("time", 0), reverse=True)
    return jsonify({"ok": True, "trades": all_trades[:50]})

# ─── LOT SIZE ──────────────────────────────────────────────────────────────

def get_lot_size(symbol, testnet=False):
    base = TESTNET_BASE if testnet else BINANCE_BASE
    try:
        r    = requests.get(f"{base}/api/v3/exchangeInfo", params={"symbol": symbol}, timeout=10)
        info = r.json()
        for s in info.get("symbols", []):
            if s["symbol"] == symbol:
                for f in s["filters"]:
                    if f["filterType"] == "LOT_SIZE":
                        step      = float(f["stepSize"])
                        precision = len(f["stepSize"].rstrip("0").split(".")[-1]) if "." in f["stepSize"] else 0
                        return step, precision
    except: pass
    return 0.01, 2

def round_qty(qty, step, precision):
    qty = math.floor(qty / step) * step
    return round(qty, precision)

# ─── ORDERS ────────────────────────────────────────────────────────────────

@app.route("/api/order", methods=["POST"])
def place_order():
    d       = request.json
    testnet = d.get("testnet", False)
    try:
        step, precision = get_lot_size(d["symbol"], testnet)
        quantity        = round_qty(float(d["quantity"]), step, precision)
        if quantity <= 0:
            return jsonify({"ok": False, "error": f"Quantity too small after rounding"}), 400
        order_type = d.get("type", "LIMIT")
        params = {"symbol": d["symbol"], "side": d["side"], "type": order_type, "quantity": str(quantity)}
        if order_type == "LIMIT":
            params["timeInForce"] = "GTC"
            params["price"]       = d["price"]
        result = binance_signed(d["api_key"], d["api_secret"], "POST", "/api/v3/order", params, testnet=testnet)
        oco = {"note": "Set SL/TP manually"}
        if order_type == "LIMIT":
            try:
                oco = binance_signed(d["api_key"], d["api_secret"], "POST", "/api/v3/order/oco", {
                    "symbol": d["symbol"], "side": "SELL" if d["side"]=="BUY" else "BUY",
                    "quantity": str(quantity), "price": d["tp"],
                    "stopPrice": d["sl"], "stopLimitPrice": d["sl"], "stopLimitTimeInForce": "GTC",
                }, testnet=testnet)
            except Exception as e2:
                oco = {"error": str(e2)}
        return jsonify({"ok": True, "order": result, "oco": oco})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

@app.route("/api/cancel_order", methods=["POST"])
def cancel_order():
    d = request.json
    try:
        result = binance_signed(d["api_key"], d["api_secret"], "DELETE", "/api/v3/order",
                                {"symbol": d["symbol"], "orderId": d["orderId"]}, testnet=d.get("testnet",False))
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

@app.route("/api/close_position", methods=["POST"])
def close_position():
    """Market sell entire balance of a coin to close position instantly."""
    d       = request.json
    testnet = d.get("testnet", False)
    symbol  = d["symbol"]  # e.g. BTCUSDT
    asset   = symbol.replace("USDT", "")
    try:
        # Get current balance of the asset
        account  = binance_signed(d["api_key"], d["api_secret"], "GET", "/api/v3/account", testnet=testnet)
        balances = {b["asset"]: float(b["free"]) for b in account["balances"]}
        qty      = balances.get(asset, 0)
        if qty <= 0:
            return jsonify({"ok": False, "error": f"No {asset} balance to sell"}), 400
        # Round to correct precision
        step, precision = get_lot_size(symbol, testnet)
        qty = round_qty(qty, step, precision)
        if qty <= 0:
            return jsonify({"ok": False, "error": f"Quantity too small after rounding"}), 400
        # Place market sell
        result = binance_signed(d["api_key"], d["api_secret"], "POST", "/api/v3/order", {
            "symbol":   symbol,
            "side":     "SELL",
            "type":     "MARKET",
            "quantity": str(qty),
        }, testnet=testnet)
        return jsonify({"ok": True, "result": result, "qty": qty})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400

# In-memory TSL monitor store
_tsl_positions = {}  # key: trade_id

@app.route("/api/tsl/start", methods=["POST"])
def tsl_start():
    """Register a position for backend TSL monitoring."""
    d = request.json
    trade_id = str(d["trade_id"])
    _tsl_positions[trade_id] = {
        "symbol":       d["symbol"],
        "side":         d["side"],        # BUY or SELL
        "entry":        float(d["entry"]),
        "hard_sl":      float(d["hard_sl"]),
        "hard_tp":      float(d["hard_tp"]),
        "tsl_callback": float(d["tsl_callback"]) / 100,  # store as decimal
        "tsl_activation":float(d["tsl_activation"]),
        "peak_price":   float(d["entry"]),  # tracks highest price seen
        "current_sl":   float(d["hard_sl"]),
        "activated":    False,
        "status":       "OPEN",
        "testnet":      d.get("testnet", False),
        "api_key":      d.get("api_key",""),
        "api_secret":   d.get("api_secret",""),
        "quantity":     d.get("quantity", "0"),
    }
    return jsonify({"ok": True, "trade_id": trade_id})

@app.route("/api/tsl/status", methods=["POST"])
def tsl_status():
    """Check + update all active TSL positions. Call this every 30s from frontend."""
    d        = request.json
    testnet  = d.get("testnet", False)
    base     = TESTNET_BASE if testnet else BINANCE_BASE
    updates  = []

    for trade_id, pos in list(_tsl_positions.items()):
        if pos["status"] != "OPEN": continue
        try:
            # fetch current price
            r     = requests.get(f"{base}/api/v3/ticker/price",
                                 params={"symbol": pos["symbol"]}, timeout=5)
            price = float(r.json()["price"])
            bull  = pos["side"] == "BUY"
            fired = False
            reason = ""

            # update peak price
            if bull and price > pos["peak_price"]:
                pos["peak_price"] = price
            elif not bull and price < pos["peak_price"]:
                pos["peak_price"] = price

            # check if TSL activated
            if not pos["activated"]:
                if bull  and price >= pos["tsl_activation"]: pos["activated"] = True
                if not bull and price <= pos["tsl_activation"]: pos["activated"] = True

            # update trailing SL if activated
            if pos["activated"]:
                if bull:
                    new_sl = round(pos["peak_price"] * (1 - pos["tsl_callback"]), 8)
                    if new_sl > pos["current_sl"]:
                        pos["current_sl"] = new_sl
                else:
                    new_sl = round(pos["peak_price"] * (1 + pos["tsl_callback"]), 8)
                    if new_sl < pos["current_sl"]:
                        pos["current_sl"] = new_sl

            # check exit conditions
            if bull:
                if price <= pos["current_sl"]:  fired=True; reason="TSL" if pos["activated"] else "HARD_SL"
                if price >= pos["hard_tp"]:      fired=True; reason="TP"
            else:
                if price >= pos["current_sl"]:  fired=True; reason="TSL" if pos["activated"] else "HARD_SL"
                if price <= pos["hard_tp"]:      fired=True; reason="TP"

            if fired:
                # auto market sell
                if pos["api_key"] and pos["quantity"] != "0":
                    try:
                        step, precision = get_lot_size(pos["symbol"], testnet)
                        qty = round_qty(float(pos["quantity"]), step, precision)
                        binance_signed(pos["api_key"], pos["api_secret"], "POST", "/api/v3/order", {
                            "symbol":   pos["symbol"],
                            "side":     "SELL" if bull else "BUY",
                            "type":     "MARKET",
                            "quantity": str(qty),
                        }, testnet=pos["testnet"])
                    except: pass
                pos["status"] = reason
                pos["exit_price"] = price

            updates.append({
                "trade_id":    trade_id,
                "symbol":      pos["symbol"],
                "price":       price,
                "peak_price":  pos["peak_price"],
                "current_sl":  pos["current_sl"],
                "activated":   pos["activated"],
                "status":      pos["status"],
                "exit_price":  pos.get("exit_price"),
                "reason":      reason if fired else "",
            })
        except: pass

    return jsonify({"ok": True, "positions": updates})

@app.route("/api/tsl/stop", methods=["POST"])
def tsl_stop():
    """Remove a position from TSL monitoring."""
    trade_id = str(request.json.get("trade_id",""))
    _tsl_positions.pop(trade_id, None)
    return jsonify({"ok": True})

@app.route("/api/tsl/list", methods=["GET"])
def tsl_list():
    """List all active TSL positions."""
    return jsonify({"ok": True, "positions": list(_tsl_positions.values())})

# ─── HELPERS ───────────────────────────────────────────────────────────────

def rsi(closes, period=14):
    closes = np.array(closes, dtype=float)
    delta  = np.diff(closes)
    gain   = np.where(delta > 0, delta, 0)
    loss   = np.where(delta < 0, -delta, 0)
    ag = np.convolve(gain, np.ones(period)/period, mode='valid')
    al = np.convolve(loss, np.ones(period)/period, mode='valid')
    rs = np.divide(ag, al, out=np.zeros_like(ag), where=al!=0)
    return 100 - (100 / (1 + rs))

def ema(closes, period):
    closes = np.array(closes, dtype=float)
    k = 2/(period+1); result=[closes[0]]
    for c in closes[1:]: result.append(c*k + result[-1]*(1-k))
    return np.array(result)

def vwap_calc(candles):
    tp  = np.array([(c['h']+c['l']+c['c'])/3 for c in candles])
    vol = np.array([c['v'] for c in candles])
    return np.cumsum(tp*vol)/np.cumsum(vol)

def atr(candles, period=14):
    """Average True Range — measures average candle movement."""
    highs  = np.array([c['h'] for c in candles])
    lows   = np.array([c['l'] for c in candles])
    closes = np.array([c['c'] for c in candles])
    tr = np.maximum(highs[1:]-lows[1:],
         np.maximum(abs(highs[1:]-closes[:-1]),
                    abs(lows[1:]-closes[:-1])))
    if len(tr) < period: return float(np.mean(tr)) if len(tr)>0 else 0
    return float(np.mean(tr[-period:]))

def geometric_target(cp, direction):
    """
    Pip target from original notebook: py[-4]*0.6 + py[-3]*0.4
    cp = [X, A, B, C, D] pattern points
    Returns the geometric price target derived from pattern geometry.
    """
    if len(cp) < 5: return None
    target = cp[-4]*0.6 + cp[-3]*0.4
    return round(float(target), 8)

def peak_detect(price, order=10):
    max_idx = list(argrelextrema(price, np.greater, order=order)[0])
    min_idx = list(argrelextrema(price, np.less,    order=order)[0])
    idx = sorted(set(max_idx+min_idx+[len(price)-1]))
    ci  = idx[-5:] if len(idx)>=5 else idx
    return ci, price[ci]

PATTERNS = {
    "Gartley":   {"AB":(0.618,0.618),"BC":(0.382,0.886),"CD":(1.27,1.618)},
    "Butterfly": {"AB":(0.786,0.786),"BC":(0.382,0.886),"CD":(1.618,2.618)},
    "Bat":       {"AB":(0.382,0.500),"BC":(0.382,0.886),"CD":(1.618,2.618)},
    "Crab":      {"AB":(0.382,0.618),"BC":(0.382,0.886),"CD":(2.24,3.618)},
    # Shark: notebook skips XA->AB, checks BC vs AB and CD vs BC
    "Shark":     {"BC":(1.13,1.618), "CD":(1.618,2.24)},
    # Trio: notebook uses moves[1],[2],[3] — skips first move (XA)
    # checks AB vs XA(=moves[1]) and BC vs AB(=moves[2])
    "Trio":      {"AB":(0.618,0.786),"BC":(1.27,1.618)},
    "Cyph":      {"AB":(0.382,0.618),"BC":(1.13,1.414), "CD":(1.272,2.0)},
}

def check_ratios(moves, ratios, err):
    """Check Fibonacci ratios — faithful to original notebook."""
    XA,AB,BC,CD = moves
    def ok(v,lo,hi,ref): return (lo-err)*abs(ref)<=abs(v)<=(hi+err)*abs(ref)
    # Standard patterns (Gartley, Butterfly, Bat, Crab, Cyph)
    if "AB" in ratios and "CD" in ratios:
        return ok(AB,*ratios["AB"],XA) and ok(BC,*ratios["BC"],AB) and ok(CD,*ratios["CD"],BC)
    # Shark: no XA->AB check, BC vs AB, CD vs BC (as in notebook)
    if "BC" in ratios and "CD" in ratios and "AB" not in ratios:
        return ok(BC,*ratios["BC"],AB) and ok(CD,*ratios["CD"],BC)
    # Trio: AB vs XA, BC vs AB, no CD (as in notebook)
    if "AB" in ratios and "CD" not in ratios:
        return ok(AB,*ratios["AB"],XA) and ok(BC,*ratios["BC"],AB)
    return False

def get_direction(XA,AB,BC,CD,pat_name):
    """Direction check faithful to notebook — Shark and Trio have no CD requirement."""
    no_cd_pats = {"Shark","Trio"}
    if pat_name in no_cd_pats:
        if XA>0 and AB<0 and BC>0: return 1
        if XA<0 and AB>0 and BC<0: return -1
        return None
    # Standard 4-leg patterns
    if XA>0 and AB<0 and BC>0 and CD<0: return 1
    if XA<0 and AB>0 and BC<0 and CD>0: return -1
    return None

def pattern_score(rsi_now, trend_up, direction, atr_val, entry,
                  tp_geo, candles=None):
    """
    Score a signal 0-100 based on confirmation quality.
    Pure addition — does not affect detection logic.
    """
    score   = 0
    details = {}

    # 1. RSI in ideal zone (40-60) → +20pts
    rsi_ideal = 40 <= rsi_now <= 60
    pts = 20 if rsi_ideal else (10 if 30 <= rsi_now <= 70 else 0)
    score += pts; details["RSI"] = pts

    # 2. Trend confirmed → +20pts
    trend_ok = (direction==1 and trend_up) or (direction==-1 and not trend_up)
    pts = 20 if trend_ok else 0
    score += pts; details["Trend"] = pts

    # 3. ATR ratio — SL not too tight relative to entry → +15pts
    atr_pct = (atr_val / entry * 100) if entry > 0 else 0
    pts = 15 if atr_pct >= 0.5 else (8 if atr_pct >= 0.2 else 0)
    score += pts; details["ATR"] = pts

    # 4. Volume spike (needs candles) → +15pts
    vol_pts = 0
    if candles and len(candles) >= 20:
        vols    = np.array([c['v'] for c in candles])
        avg_vol = float(np.mean(vols[-20:]))
        cur_vol = float(vols[-1])
        if cur_vol > avg_vol * 2.0:   vol_pts = 15
        elif cur_vol > avg_vol * 1.5: vol_pts = 10
        elif cur_vol > avg_vol * 1.0: vol_pts = 5
    score += vol_pts; details["Volume"] = vol_pts

    # 5. Geometric TP available → +10pts
    pts = 10 if tp_geo else 0
    score += pts; details["GeoTP"] = pts

    # 6. Confluence placeholder → +20pts added by frontend when detected
    # (we give 0 here; frontend adds 20 if confluence found)
    details["Confluence"] = 0

    return min(score, 80), details  # max 80 here, frontend adds up to 20 for confluence

def detect_patterns(closes, err=0.25, rsi_filter=True, trend_filter=True,
                    rsi_ob=70, rsi_os=30, candles=None, atr_mult=1.5):
    price = np.array(closes)
    if len(price)<60: return []
    rsi_vals = rsi(price,14)
    rsi_now  = float(rsi_vals[-1]) if len(rsi_vals)>0 else 50
    ema50    = float(ema(price,50)[-1])
    ema200   = float(ema(price,min(200,len(price)-1))[-1])
    trend_up = ema50>ema200
    ci,cp    = peak_detect(price,10)
    if len(cp)<5: return []
    XA,AB,BC,CD = cp[1]-cp[0],cp[2]-cp[1],cp[3]-cp[2],cp[4]-cp[3]

    # ATR-based SL/TP
    atr_val = atr(candles, 14) if candles else float(np.std(price[-20:]))
    entry   = float(price[-1])
    atr_sl  = round(atr_val * atr_mult, 8)  # SL distance in price units
    atr_tp  = round(atr_val * atr_mult * 2, 8)  # TP = 2× SL (R:R 2)

    # Geometric target from notebook: py[-4]*0.6 + py[-3]*0.4
    geo_target = geometric_target(cp, None)

    results=[]
    for name,rat in PATTERNS.items():
        if not check_ratios([XA,AB,BC,CD],rat,err): continue
        d = get_direction(XA,AB,BC,CD,name)
        if d is None: continue
        if rsi_filter:
            if d==1  and rsi_now>rsi_ob: continue
            if d==-1 and rsi_now<rsi_os: continue
        if trend_filter:
            if d==1  and not trend_up: continue
            if d==-1 and trend_up:     continue

        # ATR SL/TP levels
        sl_atr = round(entry - atr_sl, 8) if d==1 else round(entry + atr_sl, 8)
        tp_atr = round(entry + atr_tp, 8) if d==1 else round(entry - atr_tp, 8)
        # Use geometric target as TP if it makes sense directionally
        tp_geo = geo_target
        if tp_geo and d==1  and tp_geo <= entry: tp_geo = None
        if tp_geo and d==-1 and tp_geo >= entry: tp_geo = None

        # ── Trailing SL parameters ──
        # Callback rate = ATR / entry * 100 (as %)
        # Minimum 0.5%, maximum 5% — clamped for safety
        tsl_callback = round(min(max((atr_val / entry * 100) * 1.5, 0.5), 5.0), 2)
        # Activation price = entry + 1×ATR (bull) or entry - 1×ATR (bear)
        # TSL only starts once you're already winning by 1 ATR
        tsl_activation = round(entry + atr_val, 8) if d==1 else round(entry - atr_val, 8)
        # Hard SL = entry - 1.5×ATR (immediate protection if price drops right away)
        hard_sl = round(entry - atr_val * 1.5, 8) if d==1 else round(entry + atr_val * 1.5, 8)
        # Hard TP = geometric target if available, else 3×ATR
        hard_tp = tp_geo if tp_geo else (round(entry + atr_val * 3, 8) if d==1 else round(entry - atr_val * 3, 8))

        # Pattern strength score (pure addition, no effect on detection)
        score, score_details = pattern_score(
            rsi_now, trend_up, d, atr_val, entry, tp_geo, candles)

        results.append({
            "pattern":        name,
            "direction":      d,
            "rsi":            round(rsi_now, 2),
            "trend_up":       bool(trend_up),
            "ema50":          round(ema50, 8),
            "ema200":         round(ema200, 8),
            "points":         [round(float(p),8) for p in cp],
            "idx":            [int(i) for i in ci],
            "atr":            round(atr_val, 8),
            "sl_atr":         sl_atr,
            "tp_atr":         tp_atr,
            "tp_geo":         tp_geo,
            "entry":          round(entry, 8),
            # TSL fields
            "tsl_callback":   tsl_callback,
            "tsl_activation": round(tsl_activation, 8),
            "hard_sl":        round(hard_sl, 8),
            "hard_tp":        round(hard_tp, 8),
            # Score
            "score":          score,
            "score_details":  score_details,
        })
    return results

def detect_scalp(candles, err=0.20):
    if len(candles)<50: return []
    closes  = np.array([c['c'] for c in candles])
    volumes = np.array([c['v'] for c in candles])
    rsi_vals= rsi(closes,14)
    rsi_now = float(rsi_vals[-1]) if len(rsi_vals)>0 else 50
    if not (40<=rsi_now<=60): return []
    vwap_now  = float(vwap_calc(candles)[-1])
    avg_vol   = float(np.mean(volumes[-20:]))
    if float(volumes[-1]) <= avg_vol*1.5: return []
    ci,cp = peak_detect(closes,3)
    if len(cp)<5: return []
    XA,AB,BC,CD = cp[1]-cp[0],cp[2]-cp[1],cp[3]-cp[2],cp[4]-cp[3]
    results=[]
    for name,rat in PATTERNS.items():
        if not check_ratios([XA,AB,BC,CD],rat,err): continue
        d = get_direction(XA,AB,BC,CD,name)
        if d is None: continue
        price_now = float(closes[-1])
        if d==1  and price_now<vwap_now: continue
        if d==-1 and price_now>vwap_now: continue
        results.append({"pattern":name,"direction":d,"rsi":round(rsi_now,2),
                         "vwap":round(vwap_now,8),"vol_ratio":round(float(volumes[-1])/avg_vol,2),
                         "points":[round(float(p),8) for p in cp],"scalp":True})
    return results

def fetch_closes(symbol, interval, limit, testnet=False):
    base = TESTNET_BASE if testnet else BINANCE_BASE
    r = requests.get(f"{base}/api/v3/klines",
                     params={"symbol":symbol,"interval":interval,"limit":limit}, timeout=12)
    r.raise_for_status()
    raw = r.json()
    closes  = [float(k[4]) for k in raw]
    candles = [{"t":k[0],"o":float(k[1]),"h":float(k[2]),"l":float(k[3]),"c":float(k[4]),"v":float(k[5])} for k in raw]
    return closes, candles, raw

# ─── SCAN (multi-coin) ─────────────────────────────────────────────────────

@app.route("/api/scan", methods=["POST"])
def scan():
    d       = request.json
    symbols = d.get("symbols", [d.get("symbol","BTCUSDT")])
    if isinstance(symbols, str): symbols = [symbols]
    interval= d.get("interval","1h")
    err     = d.get("err_allowed",0.25)
    rsi_f   = d.get("rsi_filter",True)
    trend_f = d.get("trend_filter",True)
    rsi_ob  = d.get("rsi_ob",70)
    rsi_os  = d.get("rsi_os",30)
    testnet = d.get("testnet",False)
    scalp   = d.get("scalp",False)
    all_signals = []
    for symbol in symbols:
        try:
            closes, candles, _ = fetch_closes(symbol, interval, 500, testnet)
            sigs = detect_scalp(candles,err) if scalp else detect_patterns(closes,err,rsi_f,trend_f,rsi_ob,rsi_os,candles=candles)
            for s in sigs:
                s["symbol"] = symbol
            all_signals.extend(sigs)
            time.sleep(0.1)
        except Exception as e:
            all_signals.append({"symbol":symbol,"error":str(e)})
    return jsonify({"ok":True,"signals":all_signals})

# ─── BACKTEST (multi-coin) ─────────────────────────────────────────────────

def run_backtest_one(symbol, interval, limit, sl_pct, tp_pct, err, rsi_f, trend_f, patterns, testnet):
    try:
        closes, candles, klines = fetch_closes(symbol, interval, limit, testnet)
        closes = np.array(closes)
        highs  = np.array([float(k[2]) for k in klines])
        lows   = np.array([float(k[3]) for k in klines])
        times  = [k[0] for k in klines]
        trades = []
        for i in range(60, len(closes)-5):
            for sig in detect_patterns(closes[:i], err, rsi_f, trend_f,
                                       candles=candles[:i] if i<=len(candles) else candles):
                if sig["pattern"] not in patterns: continue
                entry=closes[i]
                # Use ATR-based SL/TP if available, else fall back to fixed %
                sl=sig.get("hard_sl") or (entry*(1-sl_pct) if sig["direction"]==1 else entry*(1+sl_pct))
                tp=sig.get("hard_tp") or (entry*(1+tp_pct) if sig["direction"]==1 else entry*(1-tp_pct))
                outcome="open"; ep=closes[min(i+20,len(closes)-1)]; eb=i+20
                for j in range(i+1,min(i+21,len(closes))):
                    if sig["direction"]==1:
                        if lows[j]<=sl:  outcome="SL";ep=sl;eb=j;break
                        if highs[j]>=tp: outcome="TP";ep=tp;eb=j;break
                    else:
                        if highs[j]>=sl: outcome="SL";ep=sl;eb=j;break
                        if lows[j]<=tp:  outcome="TP";ep=tp;eb=j;break
                pnl=((ep-entry)/entry*100)*sig["direction"]
                trades.append({"pattern":sig["pattern"],"direction":sig["direction"],
                                "entry":round(float(entry),8),"exit":round(float(ep),8),
                                "sl":round(float(sl),8),"tp":round(float(tp),8),
                                "outcome":outcome,"pnl_pct":round(float(pnl),4),
                                "rsi":sig["rsi"],"entry_time":times[i]})
        if not trades:
            return {"symbol":symbol,"trades":[],"stats":{"total_trades":0,"win_rate":0,"total_pnl_pct":0,"sharpe":0,"wins":0,"losses":0,"avg_win_pct":0,"avg_loss_pct":0,"by_pattern":{}}}
        wins=[t for t in trades if t["outcome"]=="TP"]
        losses=[t for t in trades if t["outcome"]=="SL"]
        total=len(trades)
        wr=len(wins)/total*100 if total else 0
        aw=float(np.mean([t["pnl_pct"] for t in wins]))   if wins   else 0
        al=float(np.mean([t["pnl_pct"] for t in losses])) if losses else 0
        tp2=sum(t["pnl_pct"] for t in trades)
        sh=float(np.mean([t["pnl_pct"] for t in trades])/(np.std([t["pnl_pct"] for t in trades])+1e-9)*np.sqrt(252))
        by_p={}
        for t in trades:
            p=t["pattern"]
            if p not in by_p: by_p[p]={"total":0,"wins":0,"pnl":0}
            by_p[p]["total"]+=1
            if t["outcome"]=="TP": by_p[p]["wins"]+=1
            by_p[p]["pnl"]+=t["pnl_pct"]
        for p in by_p:
            by_p[p]["win_rate"]=round(by_p[p]["wins"]/by_p[p]["total"]*100,1)
            by_p[p]["pnl"]=round(by_p[p]["pnl"],2)
        return {"symbol":symbol,"trades":trades[-50:],"stats":{
            "total_trades":total,"wins":len(wins),"losses":len(losses),
            "win_rate":round(wr,1),"avg_win_pct":round(aw,3),"avg_loss_pct":round(al,3),
            "total_pnl_pct":round(float(tp2),3),"sharpe":round(sh,2),"by_pattern":by_p}}
    except Exception as e:
        return {"symbol":symbol,"error":str(e),"trades":[],"stats":{}}

@app.route("/api/backtest", methods=["POST"])
def backtest():
    d        = request.json
    symbols  = d.get("symbols", [d.get("symbol","BTCUSDT")])
    if isinstance(symbols, str): symbols = [symbols]
    interval = d.get("interval","1h")
    limit    = int(d.get("limit",500))
    sl_pct   = float(d.get("sl_pct",2.0))/100
    tp_pct   = float(d.get("tp_pct",4.0))/100
    err      = float(d.get("err_allowed",0.25))
    rsi_f    = d.get("rsi_filter",True)
    trend_f  = d.get("trend_filter",True)
    patterns = d.get("patterns",list(PATTERNS.keys()))
    testnet  = d.get("testnet",False)
    results  = []
    for symbol in symbols:
        res = run_backtest_one(symbol,interval,limit,sl_pct,tp_pct,err,rsi_f,trend_f,patterns,testnet)
        results.append(res)
        time.sleep(0.2)
    return jsonify({"ok":True,"results":results})

# ─── EXCEL EXPORT ──────────────────────────────────────────────────────────

@app.route("/api/export_excel", methods=["POST"])
def export_excel():
    d       = request.json
    results = d.get("results", [])
    wb      = openpyxl.Workbook()

    GOLD  = "FFE8A838"; DARK  = "FF0D0F14"; GREEN = "FF2DD4A0"
    RED   = "FFE8504A"; BLUE  = "FF4A9DE8"; WHITE = "FFD8DCE8"
    GREY  = "FF12151C"; GREY2 = "FF171B25"

    def hf(c): return PatternFill("solid", fgColor=c)
    def bd():
        s = Side(style='thin', color="FF1E2433")
        return Border(left=s,right=s,top=s,bottom=s)
    def hdr(ws,row,cols,bg=GREY2):
        for col,h in enumerate(cols,1):
            c=ws.cell(row=row,column=col,value=h)
            c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
            c.fill=hf(bg); c.alignment=Alignment(horizontal="center"); c.border=bd()

    # ── Summary sheet ──
    ws = wb.active; ws.title = "Summary"
    ws.merge_cells("A1:H1")
    ws["A1"]="HarmonicBot Pro — Multi-Coin Backtest Report"
    ws["A1"].font=Font(name="Arial",bold=True,size=15,color=GOLD)
    ws["A1"].fill=hf(DARK); ws["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=36
    ws.merge_cells("A2:H2")
    ws["A2"]=f"Generated: {time.strftime('%Y-%m-%d %H:%M')}  |  Coins: {len(results)}"
    ws["A2"].font=Font(name="Arial",size=9,color="FF6B7590",italic=True)
    ws["A2"].fill=hf(DARK); ws["A2"].alignment=Alignment(horizontal="center")
    ws.row_dimensions[2].height=18; ws.row_dimensions[3].height=8

    hdr(ws,4,["Symbol","Trades","Wins","Losses","Win Rate","Total P&L","Sharpe","Best Pattern"])
    ws.row_dimensions[4].height=20

    for i,res in enumerate(results,1):
        r=4+i; s=res.get("stats",{}); sym=res.get("symbol","?")
        fill=hf(DARK) if i%2==0 else hf(GREY)
        for col in range(1,9):
            ws.cell(row=r,column=col).fill=fill; ws.cell(row=r,column=col).border=bd()
        wr=s.get("win_rate",0); pnl=s.get("total_pnl_pct",0); sh=s.get("sharpe",0)
        by_p=s.get("by_pattern",{})
        best_pat=max(by_p,key=lambda x:by_p[x].get("win_rate",0)) if by_p else "—"
        wr_col=GREEN if wr>=60 else GOLD if wr>=45 else RED
        pnl_col=GREEN if pnl>=0 else RED
        ws.cell(row=r,column=1,value=sym).font=Font(name="Arial",bold=True,size=10,color=GOLD)
        ws.cell(row=r,column=2,value=s.get("total_trades",0)).font=Font(name="Arial",size=10,color=WHITE)
        ws.cell(row=r,column=3,value=s.get("wins",0)).font=Font(name="Arial",size=10,color=GREEN)
        ws.cell(row=r,column=4,value=s.get("losses",0)).font=Font(name="Arial",size=10,color=RED)
        ws.cell(row=r,column=5,value=f"{wr}%").font=Font(name="Arial",bold=True,size=11,color=wr_col)
        ws.cell(row=r,column=6,value=f"{'+' if pnl>=0 else ''}{pnl}%").font=Font(name="Arial",bold=True,size=11,color=pnl_col)
        ws.cell(row=r,column=7,value=sh).font=Font(name="Arial",size=10,color=BLUE)
        ws.cell(row=r,column=8,value=best_pat).font=Font(name="Arial",size=10,color=GOLD)
        for col in [2,3,4,5,6,7,8]:
            ws.cell(row=r,column=col).alignment=Alignment(horizontal="center")
        ws.row_dimensions[r].height=22

    for i,w in enumerate([14,10,8,10,12,14,10,16],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.sheet_properties.tabColor="E8A838"

    # ── Per-coin sheets ──
    for res in results:
        sym  = res.get("symbol","UNKNOWN")
        s    = res.get("stats",{})
        trades=res.get("trades",[])
        wsn  = wb.create_sheet(sym[:28])

        wsn.merge_cells("A1:G1")
        wsn["A1"]=sym
        wsn["A1"].font=Font(name="Arial",bold=True,size=14,color=GOLD)
        wsn["A1"].fill=hf(DARK); wsn["A1"].alignment=Alignment(horizontal="center",vertical="center")
        wsn.row_dimensions[1].height=30

        if res.get("error"):
            wsn["A3"]=f"Error: {res['error']}"
            wsn["A3"].font=Font(name="Arial",size=10,color=RED)
            continue

        kpis=[("Win Rate",f"{s.get('win_rate',0)}%",GREEN if s.get('win_rate',0)>=60 else GOLD),
              ("Total P&L",f"{'+' if s.get('total_pnl_pct',0)>=0 else ''}{s.get('total_pnl_pct',0)}%",GREEN if s.get('total_pnl_pct',0)>=0 else RED),
              ("Sharpe",str(s.get('sharpe',0)),BLUE),
              ("Trades",str(s.get('total_trades',0)),WHITE),
              ("Wins",str(s.get('wins',0)),GREEN),
              ("Losses",str(s.get('losses',0)),RED)]
        for col,(label,val,color) in enumerate(kpis,1):
            wsn.cell(row=3,column=col,value=label).font=Font(name="Arial",size=8,color="FF6B7590")
            wsn.cell(row=3,column=col).fill=hf(GREY2)
            wsn.cell(row=3,column=col).alignment=Alignment(horizontal="center")
            wsn.cell(row=4,column=col,value=val).font=Font(name="Arial",bold=True,size=13,color=color)
            wsn.cell(row=4,column=col).fill=hf(GREY)
            wsn.cell(row=4,column=col).alignment=Alignment(horizontal="center")
        wsn.row_dimensions[3].height=18; wsn.row_dimensions[4].height=28; wsn.row_dimensions[5].height=10

        by_p=s.get("by_pattern",{})
        if by_p:
            wsn.cell(row=6,column=1,value="BY PATTERN").font=Font(name="Arial",bold=True,size=9,color="FF444D65")
            wsn.cell(row=6,column=1).fill=hf(DARK)
            hdr(wsn,7,["Pattern","Trades","Win Rate","P&L %"])
            for pi,(pat,pv) in enumerate(by_p.items(),1):
                pr=7+pi; pwr=pv.get("win_rate",0); ppnl=pv.get("pnl",0)
                pf=hf(DARK) if pi%2==0 else hf(GREY)
                for col in range(1,5):
                    wsn.cell(row=pr,column=col).fill=pf; wsn.cell(row=pr,column=col).border=bd()
                wsn.cell(row=pr,column=1,value=pat).font=Font(name="Arial",size=10,color=GOLD)
                wsn.cell(row=pr,column=2,value=pv.get("total",0)).font=Font(name="Arial",size=10,color=WHITE)
                wsn.cell(row=pr,column=2).alignment=Alignment(horizontal="center")
                wsn.cell(row=pr,column=3,value=f"{pwr}%").font=Font(name="Arial",bold=True,size=10,color=GREEN if pwr>=60 else GOLD if pwr>=45 else RED)
                wsn.cell(row=pr,column=3).alignment=Alignment(horizontal="center")
                wsn.cell(row=pr,column=4,value=f"{'+' if ppnl>=0 else ''}{ppnl}%").font=Font(name="Arial",size=10,color=GREEN if ppnl>=0 else RED)
                wsn.cell(row=pr,column=4).alignment=Alignment(horizontal="center")
                wsn.row_dimensions[pr].height=20
            trade_start = 7+len(by_p)+2
        else:
            trade_start = 8

        if trades:
            wsn.cell(row=trade_start,column=1,value="TRADE LOG").font=Font(name="Arial",bold=True,size=9,color="FF444D65")
            wsn.cell(row=trade_start,column=1).fill=hf(DARK)
            hdr(wsn,trade_start+1,["#","Pattern","Dir","Entry","Exit","Outcome","P&L %","RSI"])
            for ti,t in enumerate(trades,1):
                tr=trade_start+1+ti
                tf=hf(DARK) if ti%2==0 else hf(GREY)
                for col in range(1,9):
                    wsn.cell(row=tr,column=col).fill=tf; wsn.cell(row=tr,column=col).border=bd()
                wsn.cell(row=tr,column=1,value=ti).font=Font(name="Arial",size=9,color="FF6B7590")
                wsn.cell(row=tr,column=2,value=t["pattern"]).font=Font(name="Arial",size=9,color=GOLD)
                dir_sym="▲" if t["direction"]==1 else "▼"
                dir_col=GREEN if t["direction"]==1 else RED
                wsn.cell(row=tr,column=3,value=dir_sym).font=Font(name="Arial",bold=True,size=10,color=dir_col)
                wsn.cell(row=tr,column=3).alignment=Alignment(horizontal="center")
                wsn.cell(row=tr,column=4,value=t["entry"]).font=Font(name="Arial",size=9,color=WHITE)
                wsn.cell(row=tr,column=5,value=t["exit"]).font=Font(name="Arial",size=9,color=WHITE)
                out_col=GREEN if t["outcome"]=="TP" else RED
                wsn.cell(row=tr,column=6,value=t["outcome"]).font=Font(name="Arial",bold=True,size=9,color=out_col)
                wsn.cell(row=tr,column=6).alignment=Alignment(horizontal="center")
                pnl=t["pnl_pct"]
                wsn.cell(row=tr,column=7,value=f"{'+' if pnl>=0 else ''}{pnl}%").font=Font(name="Arial",bold=True,size=9,color=GREEN if pnl>=0 else RED)
                wsn.cell(row=tr,column=7).alignment=Alignment(horizontal="center")
                wsn.cell(row=tr,column=8,value=t["rsi"]).font=Font(name="Arial",size=9,color=WHITE)
                wsn.row_dimensions[tr].height=18

        for i,w in enumerate([8,12,6,12,12,10,10,8],1):
            wsn.column_dimensions[get_column_letter(i)].width=w
        wr=s.get("win_rate",0)
        wsn.sheet_properties.tabColor="2DD4A0" if wr>=60 else "E8A838" if wr>=45 else "E8504A"

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"HarmonicBot_{time.strftime('%Y%m%d_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/export_orders_excel", methods=["POST"])
def export_orders_excel():
    d       = request.json
    orders  = d.get("orders", [])
    history = d.get("history", [])
    wb = openpyxl.Workbook()

    GOLD="FFE8A838"; DARK="FF0D0F14"; GREEN="FF2DD4A0"
    RED="FFE8504A"; WHITE="FFD8DCE8"; GREY="FF12151C"; GREY2="FF171B25"

    def hf(c): return PatternFill("solid", fgColor=c)
    def bd():
        s=Side(style='thin',color="FF1E2433")
        return Border(left=s,right=s,top=s,bottom=s)
    def hdr(ws,row,cols):
        for col,h in enumerate(cols,1):
            c=ws.cell(row=row,column=col,value=h)
            c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
            c.fill=hf(GREY2); c.alignment=Alignment(horizontal="center"); c.border=bd()
        ws.row_dimensions[row].height=20

    # ── Sheet 1: Open Orders ──
    ws1=wb.active; ws1.title="Open Orders"
    ws1.merge_cells("A1:H1")
    ws1["A1"]=f"HarmonicBot — Open Orders  |  {time.strftime('%Y-%m-%d %H:%M')}"
    ws1["A1"].font=Font(name="Arial",bold=True,size=13,color=GOLD)
    ws1["A1"].fill=hf(DARK); ws1["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws1.row_dimensions[1].height=30

    hdr(ws1,2,["Symbol","Side","Type","Entry Price","Stop Loss","Take Profit","Qty","Status","Time"])
    for i,o in enumerate(orders,1):
        r=2+i
        f=hf(DARK) if i%2==0 else hf(GREY)
        side=o.get("side","")
        sc=GREEN if side=="BUY" else RED
        # find matching SL/TP from OCO pairs
        price=o.get("price","—")
        sl="—"; tp="—"
        for other in orders:
            if other.get("symbol")==o.get("symbol") and other.get("orderId")!=o.get("orderId"):
                ot=other.get("type","")
                if "STOP" in ot: sl=other.get("stopPrice") or other.get("price","—")
                if "LIMIT" in ot and "STOP" not in ot: tp=other.get("price","—")
        for col in range(1,10):
            ws1.cell(row=r,column=col).fill=f; ws1.cell(row=r,column=col).border=bd()
        ws1.cell(row=r,column=1,value=o.get("symbol","")).font=Font(name="Arial",bold=True,size=10,color=GOLD)
        ws1.cell(row=r,column=2,value=side).font=Font(name="Arial",bold=True,size=10,color=sc)
        ws1.cell(row=r,column=2).alignment=Alignment(horizontal="center")
        ws1.cell(row=r,column=3,value=o.get("type","")).font=Font(name="Arial",size=9,color=WHITE)
        ws1.cell(row=r,column=4,value=str(price)).font=Font(name="Arial",size=10,color=GOLD)
        ws1.cell(row=r,column=5,value=str(sl)).font=Font(name="Arial",size=10,color=RED)
        ws1.cell(row=r,column=6,value=str(tp)).font=Font(name="Arial",size=10,color=GREEN)
        ws1.cell(row=r,column=7,value=str(o.get("origQty",""))).font=Font(name="Arial",size=9,color=WHITE)
        ws1.cell(row=r,column=8,value=o.get("status","")).font=Font(name="Arial",size=9,color=WHITE)
        ws1.cell(row=r,column=9,value=str(o.get("time",""))).font=Font(name="Arial",size=9,color="FF6B7590")
        ws1.row_dimensions[r].height=20
    for i,w in enumerate([14,8,18,14,14,14,10,10,20],1):
        ws1.column_dimensions[get_column_letter(i)].width=w
    ws1.sheet_properties.tabColor="E8A838"

    # ── Sheet 2: Trade History ──
    ws2=wb.create_sheet("Trade History")
    ws2.merge_cells("A1:G1")
    ws2["A1"]=f"HarmonicBot — Trade History  |  {time.strftime('%Y-%m-%d %H:%M')}"
    ws2["A1"].font=Font(name="Arial",bold=True,size=13,color=GOLD)
    ws2["A1"].fill=hf(DARK); ws2["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[1].height=30
    hdr(ws2,2,["Symbol","Side","Buy Price","Sell Price","Qty","Total USDT","Time","P&L est."])
    buys={}
    rows=[]
    for t in sorted(history, key=lambda x:x.get("time",0)):
        sym=t.get("symbol",""); price=float(t.get("price",0)); qty=float(t.get("qty",0))
        is_buy=t.get("isBuyer",False)
        if is_buy: buys[sym]={"price":price,"qty":qty,"time":t.get("time",0)}
        rows.append({**t,"buy_price":buys.get(sym,{}).get("price") if not is_buy else None})
    for i,t in enumerate(sorted(rows,key=lambda x:x.get("time",0),reverse=True),1):
        r=2+i; f=hf(DARK) if i%2==0 else hf(GREY)
        is_buy=t.get("isBuyer",False)
        sc=GREEN if is_buy else RED
        price=float(t.get("price",0)); qty=float(t.get("qty",0))
        total=round(price*qty,2)
        buy_px=t.get("buy_price")
        pnl="—"
        if not is_buy and buy_px:
            pnl=round((price-buy_px)*qty,2)
        for col in range(1,9):
            ws2.cell(row=r,column=col).fill=f; ws2.cell(row=r,column=col).border=bd()
        ws2.cell(row=r,column=1,value=t.get("symbol","")).font=Font(name="Arial",bold=True,size=10,color=GOLD)
        ws2.cell(row=r,column=2,value="BUY" if is_buy else "SELL").font=Font(name="Arial",bold=True,size=10,color=sc)
        ws2.cell(row=r,column=2).alignment=Alignment(horizontal="center")
        ws2.cell(row=r,column=3,value=str(buy_px) if buy_px else ("BUY" if is_buy else "—")).font=Font(name="Arial",size=10,color=GREEN)
        ws2.cell(row=r,column=4,value=str(price) if not is_buy else "—").font=Font(name="Arial",size=10,color=RED)
        ws2.cell(row=r,column=5,value=qty).font=Font(name="Arial",size=9,color=WHITE)
        ws2.cell(row=r,column=6,value=total).font=Font(name="Arial",size=10,color=WHITE)
        ws2.cell(row=r,column=7,value=str(datetime.datetime.fromtimestamp(t.get("time",0)/1000).strftime('%Y-%m-%d %H:%M') if t.get("time") else "—")).font=Font(name="Arial",size=9,color="FF6B7590")
        pnl_col=GREEN if isinstance(pnl,float) and pnl>=0 else RED
        ws2.cell(row=r,column=8,value=f"{'+' if isinstance(pnl,float) and pnl>=0 else ''}{pnl} USDT" if pnl!="—" else "—").font=Font(name="Arial",bold=True,size=10,color=pnl_col)
        ws2.row_dimensions[r].height=20
    for i,w in enumerate([14,8,14,14,10,12,18,14],1):
        ws2.column_dimensions[get_column_letter(i)].width=w
    ws2.sheet_properties.tabColor="2DD4A0"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"Orders_{time.strftime('%Y%m%d_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/export_portfolio_excel", methods=["POST"])
def export_portfolio_excel():
    d       = request.json
    trades  = d.get("trades", [])
    capital = float(d.get("capital", 1000))
    wb      = openpyxl.Workbook()

    GOLD="FFE8A838"; DARK="FF0D0F14"; GREEN="FF2DD4A0"
    RED="FFE8504A"; BLUE="FF4A9DE8"; WHITE="FFD8DCE8"
    GREY="FF12151C"; GREY2="FF171B25"

    def hf(c): return PatternFill("solid", fgColor=c)
    def bd():
        s=Side(style='thin',color="FF1E2433")
        return Border(left=s,right=s,top=s,bottom=s)
    def hdr(ws,row,cols):
        for col,h in enumerate(cols,1):
            c=ws.cell(row=row,column=col,value=h)
            c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
            c.fill=hf(GREY2);c.alignment=Alignment(horizontal="center");c.border=bd()
        ws.row_dimensions[row].height=20

    # ── Sheet 1: Summary ──
    ws1=wb.active; ws1.title="Summary"
    ws1.merge_cells("A1:F1")
    ws1["A1"]=f"HarmonicBot Pro — Portfolio Report  |  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws1["A1"].font=Font(name="Arial",bold=True,size=14,color=GOLD)
    ws1["A1"].fill=hf(DARK);ws1["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws1.row_dimensions[1].height=32

    closed=[t for t in trades if t.get("status") in ["WIN","LOSS"]]
    wins=[t for t in closed if t.get("status")=="WIN"]
    losses=[t for t in closed if t.get("status")=="LOSS"]
    total_pnl=sum(float(t.get("pnl",0) or 0) for t in closed)
    current=capital+total_pnl
    wr=round(len(wins)/len(closed)*100,1) if closed else 0
    avg_win=round(sum(float(t.get("pnl",0) or 0) for t in wins)/len(wins),2) if wins else 0
    avg_loss=round(sum(float(t.get("pnl",0) or 0) for t in losses)/len(losses),2) if losses else 0
    best=max(closed,key=lambda t:float(t.get("pnl",0) or 0)) if closed else None
    worst=min(closed,key=lambda t:float(t.get("pnl",0) or 0)) if closed else None

    kpis=[
        ("Starting Capital", f"{capital:.2f} USDT", GOLD),
        ("Current Capital",  f"{current:.2f} USDT", GREEN if current>=capital else RED),
        ("Total P&L",        f"{'+' if total_pnl>=0 else ''}{total_pnl:.2f} USDT", GREEN if total_pnl>=0 else RED),
        ("Win Rate",         f"{wr}%", GREEN if wr>=55 else GOLD if wr>=45 else RED),
        ("Total Trades",     str(len(trades)), WHITE),
        ("Closed Trades",    f"{len(wins)}W / {len(losses)}L", WHITE),
        ("Avg Win",          f"+{avg_win} USDT", GREEN),
        ("Avg Loss",         f"{avg_loss} USDT", RED),
        ("Best Trade",       f"+{float(best.get('pnl',0)):.2f} USDT ({best.get('symbol','')} {best.get('pattern','')})" if best else "—", GREEN),
        ("Worst Trade",      f"{float(worst.get('pnl',0)):.2f} USDT ({worst.get('symbol','')} {worst.get('pattern','')})" if worst else "—", RED),
    ]

    hdr(ws1,3,["Metric","Value"])
    for i,(label,val,col) in enumerate(kpis,1):
        r=3+i; f=hf(DARK) if i%2==0 else hf(GREY)
        for c in range(1,3): ws1.cell(row=r,column=c).fill=f; ws1.cell(row=r,column=c).border=bd()
        ws1.cell(row=r,column=1,value=label).font=Font(name="Arial",size=10,color="FF6B7590")
        ws1.cell(row=r,column=2,value=val).font=Font(name="Arial",bold=True,size=10,color=col)
        ws1.row_dimensions[r].height=22

    # per pattern breakdown
    by_pat={}
    for t in closed:
        p=t.get("pattern","?")
        if p not in by_pat: by_pat[p]={"total":0,"wins":0,"pnl":0.0}
        by_pat[p]["total"]+=1
        if t.get("status")=="WIN": by_pat[p]["wins"]+=1
        by_pat[p]["pnl"]+=float(t.get("pnl",0) or 0)

    pr=3+len(kpis)+2
    ws1.merge_cells(f"A{pr}:F{pr}")
    ws1[f"A{pr}"]="BY PATTERN"
    ws1[f"A{pr}"].font=Font(name="Arial",bold=True,size=9,color="FF444D65")
    ws1[f"A{pr}"].fill=hf(DARK)
    hdr(ws1,pr+1,["Pattern","Trades","Wins","Win Rate","P&L"])
    for i,(pat,v) in enumerate(by_pat.items(),1):
        r=pr+1+i; f=hf(DARK) if i%2==0 else hf(GREY)
        pwr=round(v["wins"]/v["total"]*100,1) if v["total"] else 0
        for c in range(1,6): ws1.cell(row=r,column=c).fill=f; ws1.cell(row=r,column=c).border=bd()
        ws1.cell(row=r,column=1,value=pat).font=Font(name="Arial",bold=True,size=10,color=GOLD)
        ws1.cell(row=r,column=2,value=v["total"]).font=Font(name="Arial",size=10,color=WHITE)
        ws1.cell(row=r,column=3,value=v["wins"]).font=Font(name="Arial",size=10,color=GREEN)
        ws1.cell(row=r,column=4,value=f"{pwr}%").font=Font(name="Arial",bold=True,size=10,color=GREEN if pwr>=55 else GOLD if pwr>=45 else RED)
        ws1.cell(row=r,column=5,value=f"{'+' if v['pnl']>=0 else ''}{v['pnl']:.2f}").font=Font(name="Arial",bold=True,size=10,color=GREEN if v["pnl"]>=0 else RED)
        ws1.row_dimensions[r].height=20
    for i,w in enumerate([18,28,10,10,12,14],1):
        ws1.column_dimensions[get_column_letter(i)].width=w
    ws1.sheet_properties.tabColor="E8A838"

    # ── Sheet 2: Trade Journal ──
    ws2=wb.create_sheet("Trade Journal")
    ws2.merge_cells("A1:M1")
    ws2["A1"]="Trade Journal — Full Details"
    ws2["A1"].font=Font(name="Arial",bold=True,size=13,color=GOLD)
    ws2["A1"].fill=hf(DARK);ws2["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[1].height=28
    hdr(ws2,2,["#","Date","Symbol","Pattern","TF","Side","Entry","SL","TP","Size (USDT)","P&L (USDT)","P&L %","Status"])
    for i,t in enumerate(reversed(trades),1):
        r=2+i; f=hf(DARK) if i%2==0 else hf(GREY)
        side=t.get("side",""); pnl=t.get("pnl"); pnl_pct=t.get("pnl_pct")
        sc=GREEN if side=="BUY" else RED
        pc=GREEN if (pnl or 0)>=0 else RED
        status=t.get("status","OPEN")
        sc2=GREEN if status=="WIN" else RED if status=="LOSS" else GOLD
        for c in range(1,14): ws2.cell(row=r,column=c).fill=f; ws2.cell(row=r,column=c).border=bd()
        vals=[(str(len(trades)-i+1),WHITE),(t.get("date",""),  "FF6B7590"),
              (t.get("symbol",""),   GOLD),(t.get("pattern",""),WHITE),
              (t.get("tf",""),       "FF6B7590"),(side,sc),(t.get("entry",""),GOLD),
              (t.get("sl",""),       RED),(t.get("tp",""),GREEN),
              (t.get("size",""),     WHITE),
              (f"{'+' if (pnl or 0)>=0 else ''}{pnl:.2f}" if pnl is not None else "OPEN", pc),
              (f"{'+' if (pnl_pct or 0)>=0 else ''}{pnl_pct:.2f}%" if pnl_pct is not None else "—", pc),
              (status, sc2)]
        for ci,(v,col) in enumerate(vals,1):
            ws2.cell(row=r,column=ci,value=v).font=Font(name="Arial",size=9,color=col)
            ws2.cell(row=r,column=ci).alignment=Alignment(horizontal="center")
        ws2.row_dimensions[r].height=18
    for i,w in enumerate([5,18,10,10,6,6,12,12,12,10,12,10,8],1):
        ws2.column_dimensions[get_column_letter(i)].width=w
    ws2.sheet_properties.tabColor="2DD4A0"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf,as_attachment=True,
                     download_name=f"Portfolio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/export_intel_excel", methods=["POST"])
def export_intel_excel():
    d       = request.json
    signals = d.get("signals", [])
    wb      = openpyxl.Workbook()

    GOLD="FFE8A838"; DARK="FF0D0F14"; GREEN="FF2DD4A0"
    RED="FFE8504A"; BLUE="FF4A9DE8"; WHITE="FFD8DCE8"
    GREY="FF12151C"; GREY2="FF171B25"

    def hf(c): return PatternFill("solid", fgColor=c)
    def bd():
        s=Side(style='thin',color="FF1E2433")
        return Border(left=s,right=s,top=s,bottom=s)
    def hdr(ws,row,cols):
        for col,h in enumerate(cols,1):
            c=ws.cell(row=row,column=col,value=h)
            c.font=Font(name="Arial",bold=True,size=9,color=WHITE)
            c.fill=hf(GREY2);c.alignment=Alignment(horizontal="center");c.border=bd()
        ws.row_dimensions[row].height=20

    # ── Sheet 1: Summary ──
    ws1=wb.active; ws1.title="Summary"
    ws1.merge_cells("A1:F1")
    ws1["A1"]=f"Signal Intelligence Report  |  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws1["A1"].font=Font(name="Arial",bold=True,size=14,color=GOLD)
    ws1["A1"].fill=hf(DARK);ws1["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws1.row_dimensions[1].height=30

    bulls=[s for s in signals if s.get("direction")==1]
    bears=[s for s in signals if s.get("direction")==-1]
    checked=[s for s in signals if s.get("checked")]
    correct=[s for s in checked if s.get("correct")]
    bull_pct=round(len(bulls)/len(signals)*100,1) if signals else 0
    acc=round(len(correct)/len(checked)*100,1) if checked else None

    # coin counts
    coin_count={}
    for s in signals:
        coin_count[s.get("symbol","?")]= coin_count.get(s.get("symbol","?"),0)+1
    top_coin=max(coin_count,key=coin_count.get) if coin_count else "—"

    # pattern counts
    pat_count={}
    for s in signals:
        pat_count[s.get("pattern","?")]= pat_count.get(s.get("pattern","?"),0)+1
    top_pat=max(pat_count,key=pat_count.get) if pat_count else "—"

    kpis=[
        ("Total Signals",   str(len(signals)),          GOLD),
        ("Bull Signals",    f"{len(bulls)} ({bull_pct}%)", GREEN),
        ("Bear Signals",    f"{len(bears)} ({100-bull_pct}%)", RED),
        ("Market Bias",     "▲ BULL" if bull_pct>55 else "▼ BEAR" if bull_pct<45 else "⟺ NEUTRAL",
                            GREEN if bull_pct>55 else RED if bull_pct<45 else GOLD),
        ("Top Coin",        top_coin,                   GOLD),
        ("Top Pattern",     top_pat,                    GOLD),
        ("Signals Checked", str(len(checked)),          WHITE),
        ("Accuracy",        f"{acc}%" if acc else "—",  GREEN if (acc or 0)>=60 else GOLD if (acc or 0)>=45 else RED),
    ]

    hdr(ws1,3,["Metric","Value"])
    for i,(label,val,col) in enumerate(kpis,1):
        r=3+i; f=hf(DARK) if i%2==0 else hf(GREY)
        for c in range(1,3): ws1.cell(row=r,column=c).fill=f; ws1.cell(row=r,column=c).border=bd()
        ws1.cell(row=r,column=1,value=label).font=Font(name="Arial",size=10,color="FF6B7590")
        ws1.cell(row=r,column=2,value=val).font=Font(name="Arial",bold=True,size=11,color=col)
        ws1.row_dimensions[r].height=22

    # per-pattern accuracy
    pat_acc={}
    for s in checked:
        p=s.get("pattern","?")
        if p not in pat_acc: pat_acc[p]={"total":0,"correct":0}
        pat_acc[p]["total"]+=1
        if s.get("correct"): pat_acc[p]["correct"]+=1

    pr=3+len(kpis)+2
    ws1.merge_cells(f"A{pr}:F{pr}")
    ws1[f"A{pr}"]="ACCURACY BY PATTERN"
    ws1[f"A{pr}"].font=Font(name="Arial",bold=True,size=9,color="FF444D65")
    ws1[f"A{pr}"].fill=hf(DARK)
    hdr(ws1,pr+1,["Pattern","Checked","Correct","Accuracy"])
    for i,(p,v) in enumerate(sorted(pat_acc.items(),key=lambda x:-x[1]["total"]),1):
        r=pr+1+i; f=hf(DARK) if i%2==0 else hf(GREY)
        pct=round(v["correct"]/v["total"]*100,1) if v["total"] else 0
        for c in range(1,5): ws1.cell(row=r,column=c).fill=f; ws1.cell(row=r,column=c).border=bd()
        ws1.cell(row=r,column=1,value=p).font=Font(name="Arial",bold=True,size=10,color=GOLD)
        ws1.cell(row=r,column=2,value=v["total"]).font=Font(name="Arial",size=10,color=WHITE)
        ws1.cell(row=r,column=3,value=v["correct"]).font=Font(name="Arial",size=10,color=GREEN)
        ws1.cell(row=r,column=4,value=f"{pct}%").font=Font(name="Arial",bold=True,size=10,
            color=GREEN if pct>=60 else GOLD if pct>=45 else RED)
        ws1.row_dimensions[r].height=20

    # per-coin breakdown
    coin_data={}
    for s in signals:
        sym=s.get("symbol","?")
        if sym not in coin_data: coin_data[sym]={"total":0,"bull":0,"bear":0}
        coin_data[sym]["total"]+=1
        if s.get("direction")==1: coin_data[sym]["bull"]+=1
        else: coin_data[sym]["bear"]+=1

    cr=pr+1+len(pat_acc)+2
    ws1.merge_cells(f"A{cr}:F{cr}")
    ws1[f"A{cr}"]="SIGNALS BY COIN"
    ws1[f"A{cr}"].font=Font(name="Arial",bold=True,size=9,color="FF444D65")
    ws1[f"A{cr}"].fill=hf(DARK)
    hdr(ws1,cr+1,["Symbol","Total","Bull","Bear","Bias"])
    for i,(sym,v) in enumerate(sorted(coin_data.items(),key=lambda x:-x[1]["total"]),1):
        r=cr+1+i; f=hf(DARK) if i%2==0 else hf(GREY)
        bias="▲ BULL" if v["bull"]>v["bear"] else "▼ BEAR" if v["bear"]>v["bull"] else "⟺ NEUTRAL"
        bc=GREEN if v["bull"]>v["bear"] else RED if v["bear"]>v["bull"] else GOLD
        for c in range(1,6): ws1.cell(row=r,column=c).fill=f; ws1.cell(row=r,column=c).border=bd()
        ws1.cell(row=r,column=1,value=sym).font=Font(name="Arial",bold=True,size=10,color=GOLD)
        ws1.cell(row=r,column=2,value=v["total"]).font=Font(name="Arial",size=10,color=WHITE)
        ws1.cell(row=r,column=3,value=v["bull"]).font=Font(name="Arial",size=10,color=GREEN)
        ws1.cell(row=r,column=4,value=v["bear"]).font=Font(name="Arial",size=10,color=RED)
        ws1.cell(row=r,column=5,value=bias).font=Font(name="Arial",bold=True,size=10,color=bc)
        ws1.row_dimensions[r].height=20

    for i,w in enumerate([20,28,12,12,12,14],1):
        ws1.column_dimensions[get_column_letter(i)].width=w
    ws1.sheet_properties.tabColor="E8A838"

    # ── Sheet 2: Full Signal Log ──
    ws2=wb.create_sheet("Signal Log")
    ws2.merge_cells("A1:J1")
    ws2["A1"]="Full Signal Log"
    ws2["A1"].font=Font(name="Arial",bold=True,size=13,color=GOLD)
    ws2["A1"].fill=hf(DARK);ws2["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws2.row_dimensions[1].height=28
    hdr(ws2,2,["#","Time","Symbol","Pattern","TF","Direction","Entry","RSI","Checked","Correct"])
    for i,s in enumerate(reversed(signals),1):
        r=2+i; f=hf(DARK) if i%2==0 else hf(GREY)
        bull=s.get("direction")==1
        checked_s=s.get("checked",False)
        correct_s=s.get("correct")
        for c in range(1,11): ws2.cell(row=r,column=c).fill=f; ws2.cell(row=r,column=c).border=bd()
        vals=[
            (str(len(signals)-i+1), "FF6B7590"),
            (s.get("timeStr","—"),   "FF6B7590"),
            (s.get("symbol",""),     GOLD),
            (s.get("pattern",""),    WHITE),
            (s.get("tf",""),         BLUE),
            ("▲ BULL" if bull else "▼ BEAR", GREEN if bull else RED),
            (str(round(s.get("entry",0),4)), GOLD),
            (str(s.get("rsi","—")),  WHITE),
            ("✓" if checked_s else "—", GREEN if checked_s else "FF6B7590"),
            ("✓ YES" if correct_s else "✗ NO" if correct_s is False else "—",
             GREEN if correct_s else RED if correct_s is False else "FF6B7590"),
        ]
        for ci,(v,col) in enumerate(vals,1):
            ws2.cell(row=r,column=ci,value=v).font=Font(name="Arial",size=9,color=col)
            ws2.cell(row=r,column=ci).alignment=Alignment(horizontal="center")
        ws2.row_dimensions[r].height=18
    for i,w in enumerate([5,16,12,10,6,10,12,8,10,10],1):
        ws2.column_dimensions[get_column_letter(i)].width=w
    ws2.sheet_properties.tabColor="4A9DE8"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf,as_attachment=True,
                     download_name=f"Intelligence_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ─── DCA ───────────────────────────────────────────────────────────────────

@app.route("/api/dca/place", methods=["POST"])
def dca_place():
    """
    Place multiple DCA orders at intervals below/above entry.
    Each order is a LIMIT order at a calculated price.
    Returns all order IDs + average entry calculation.
    """
    d        = request.json
    testnet  = d.get("testnet", False)
    symbol   = d["symbol"]
    side     = d["side"]          # BUY or SELL
    entry    = float(d["entry"])  # reference price
    total_usdt = float(d["total_usdt"])  # total budget
    n_orders = int(d.get("n_orders", 3))   # number of DCA orders
    drop_pct = float(d.get("drop_pct", 1.0)) / 100  # % between each order
    sl       = float(d["sl"])
    tp       = float(d["tp"])

    if n_orders < 1 or n_orders > 5:
        return jsonify({"ok": False, "error": "n_orders must be 1-5"}), 400

    step, precision = get_lot_size(symbol, testnet)
    usdt_per_order  = total_usdt / n_orders
    orders_placed   = []
    prices          = []
    errors          = []

    for i in range(n_orders):
        # Bull: each order is drop_pct% lower than entry
        # Bear: each order is drop_pct% higher than entry
        if side == "BUY":
            price = round(entry * (1 - drop_pct * i), 8)
        else:
            price = round(entry * (1 + drop_pct * i), 8)

        qty = round_qty(usdt_per_order / price, step, precision)
        if qty <= 0:
            errors.append(f"Order {i+1}: qty too small")
            continue

        prices.append(price)
        try:
            result = binance_signed(d["api_key"], d["api_secret"], "POST", "/api/v3/order", {
                "symbol":      symbol,
                "side":        side,
                "type":        "LIMIT",
                "timeInForce": "GTC",
                "quantity":    str(qty),
                "price":       str(round(price, 8)),
            }, testnet=testnet)
            orders_placed.append({
                "order_num":  i + 1,
                "price":      price,
                "qty":        qty,
                "usdt":       round(usdt_per_order, 2),
                "orderId":    result.get("orderId"),
                "status":     result.get("status"),
            })
        except Exception as e:
            errors.append(f"Order {i+1} @ {price}: {str(e)}")

    # Calculate average entry from placed orders
    if orders_placed:
        total_qty  = sum(o["qty"] for o in orders_placed)
        avg_entry  = sum(o["price"] * o["qty"] for o in orders_placed) / total_qty
        # Recalculate SL/TP from average entry
        atr_val    = abs(entry - sl) / 1.5  # reverse-engineer ATR from original SL
        if side == "BUY":
            avg_sl = round(avg_entry - atr_val * 1.5, 8)
            avg_tp = round(avg_entry + atr_val * 3.0, 8)
        else:
            avg_sl = round(avg_entry + atr_val * 1.5, 8)
            avg_tp = round(avg_entry - atr_val * 3.0, 8)
    else:
        avg_entry = entry
        avg_sl    = sl
        avg_tp    = tp

    return jsonify({
        "ok":           len(orders_placed) > 0,
        "orders":       orders_placed,
        "avg_entry":    round(avg_entry, 8),
        "avg_sl":       avg_sl,
        "avg_tp":       avg_tp,
        "total_qty":    round(sum(o["qty"] for o in orders_placed), 6),
        "total_usdt":   round(sum(o["usdt"] for o in orders_placed), 2),
        "errors":       errors,
    })

# ─── TELEGRAM ──────────────────────────────────────────────────────────────

TG_BASE = "https://api.telegram.org/bot{token}/{method}"

def tg_send(token, chat_id, text, parse_mode="HTML"):
    """Send a Telegram message. Returns True if sent, False if failed."""
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": parse_mode},
            timeout=8
        )
        return r.ok, r.json()
    except Exception as e:
        return False, {"error": str(e)}

def tg_signal_msg(sig, sym, tf):
    """Build a beautiful signal message for Telegram."""
    bull      = sig.get("direction") == 1
    arrow     = "🟢" if bull else "🔴"
    direction = "BULLISH ▲" if bull else "BEARISH ▼"
    entry     = sig.get("entry", 0)
    dp        = 8 if entry < 0.001 else 6 if entry < 0.01 else 5 if entry < 1 else 3 if entry < 100 else 2
    fmt       = lambda v: f"{float(v):.{dp}f}" if v else "—"

    score     = sig.get('score', 0)
    score_lbl = '🔥 STRONG' if score>=80 else '✅ GOOD' if score>=60 else '⚠ MODERATE' if score>=40 else '❌ WEAK'
    lines = [
        f"{arrow} <b>{direction} SIGNAL</b> — HarmonicBot",
        f"",
        f"⭐ <b>Score:</b>      {score}/100 {score_lbl}",
        f"📊 <b>Pattern:</b>    {sig.get('pattern','—')}",
        f"💰 <b>Coin:</b>       {sym.replace('USDT', '/USDT')}",
        f"⏱ <b>Timeframe:</b>  {tf}",
        f"",
        f"📈 <b>Entry:</b>      {fmt(entry)}",
        f"🛑 <b>Stop Loss:</b>  {fmt(sig.get('hard_sl') or sig.get('sl_atr'))} (ATR)",
        f"🎯 <b>Take Profit:</b> {fmt(sig.get('hard_tp') or sig.get('tp_geo') or sig.get('tp_atr'))}",
        f"",
        f"📉 <b>RSI:</b>        {sig.get('rsi', '—')}",
        f"📊 <b>ATR:</b>        {fmt(sig.get('atr'))}",
        f"🔄 <b>Trend:</b>      {'▲ UP' if sig.get('trend_up') else '▼ DOWN'}",
        f"",
        f"⟳ <b>TSL Callback:</b>   {sig.get('tsl_callback','—')}%",
        f"⚡ <b>TSL Activation:</b> {fmt(sig.get('tsl_activation'))}",
        f"",
        f"⏰ {datetime.datetime.now().strftime('%H:%M:%S')}",
    ]
    return "\n".join(lines)

def tg_order_msg(sym, side, qty, entry, sl, tp, pattern, tf, pnl=None):
    """Build an order confirmation message."""
    arrow = "🟢" if side == "BUY" else "🔴"
    dp    = 8 if float(entry) < 0.001 else 6 if float(entry) < 0.01 else 3 if float(entry) < 100 else 2
    fmt   = lambda v: f"{float(v):.{dp}f}" if v else "—"
    lines = [
        f"{arrow} <b>ORDER PLACED</b> — HarmonicBot",
        f"",
        f"💱 <b>Symbol:</b>  {sym.replace('USDT','/USDT')}",
        f"📊 <b>Pattern:</b> {pattern} ({tf})",
        f"📋 <b>Side:</b>    {side}",
        f"📈 <b>Entry:</b>   {fmt(entry)}",
        f"📦 <b>Qty:</b>     {qty}",
        f"🛑 <b>SL:</b>      {fmt(sl)}",
        f"🎯 <b>TP:</b>      {fmt(tp)}",
    ]
    if pnl is not None:
        lines += [f"", f"💵 <b>P&L:</b> {'+' if pnl >= 0 else ''}{pnl:.2f} USDT"]
    lines.append(f"⏰ {datetime.datetime.now().strftime('%H:%M:%S')}")
    return "\n".join(lines)

def tg_tsl_msg(sym, status, entry, exit_price, pnl):
    """Build a TSL fired message."""
    is_win = status == "TP"
    lines = [
        f"{'✅' if is_win else '❌'} <b>{'TAKE PROFIT HIT' if is_win else 'STOP LOSS HIT'}</b>",
        f"",
        f"💰 <b>Symbol:</b>    {sym.replace('USDT','/USDT')}",
        f"📈 <b>Entry:</b>     {entry}",
        f"📉 <b>Exit:</b>      {exit_price}",
        f"💵 <b>P&L:</b>       {'+' if pnl >= 0 else ''}{pnl:.2f} USDT",
        f"🔔 <b>Reason:</b>    {'Take Profit ✅' if is_win else 'Stop Loss / TSL ❌'}",
        f"⏰ {datetime.datetime.now().strftime('%H:%M:%S')}",
    ]
    return "\n".join(lines)

def tg_confluence_msg(symbol, direction, timeframes):
    """Confluence alert — same coin on multiple timeframes."""
    bull = direction == 1
    lines = [
        f"⚡ <b>CONFLUENCE ALERT</b> — HarmonicBot",
        f"",
        f"💰 <b>Coin:</b>       {symbol.replace('USDT','/USDT')}",
        f"{'🟢' if bull else '🔴'} <b>Direction:</b>  {'BULLISH ▲' if bull else 'BEARISH ▼'}",
        f"📊 <b>Timeframes:</b> {' + '.join(timeframes)}",
        f"",
        f"🔥 <b>STRONG SIGNAL</b> — confirmed on {len(timeframes)} timeframes!",
        f"⏰ {datetime.datetime.now().strftime('%H:%M:%S')}",
    ]
    return "\n".join(lines)

@app.route("/api/telegram/test", methods=["POST"])
def telegram_test():
    d = request.json
    token   = d.get("token","").strip()
    chat_id = d.get("chat_id","").strip()
    if not token or not chat_id:
        return jsonify({"ok": False, "error": "Token and Chat ID required"}), 400
    msg = (
        "✅ <b>HarmonicBot Connected!</b>\n\n"
        "Your Telegram alerts are working.\n"
        "You will now receive:\n"
        "• 📊 New signal detections\n"
        "• ✅ Order confirmations\n"
        "• 🛑 TSL / Stop Loss hits\n"
        "• 🎯 Take Profit hits\n"
        "• ⚡ Confluence alerts\n\n"
        f"⏰ {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    ok, resp = tg_send(token, chat_id, msg)
    if ok:
        return jsonify({"ok": True, "message": "Test message sent!"})
    return jsonify({"ok": False, "error": resp.get("description","Failed")}), 400

@app.route("/api/telegram/send_signal", methods=["POST"])
def telegram_send_signal():
    d = request.json
    ok, resp = tg_send(d["token"], d["chat_id"],
                       tg_signal_msg(d["signal"], d["symbol"], d["tf"]))
    return jsonify({"ok": ok})

@app.route("/api/telegram/send_order", methods=["POST"])
def telegram_send_order():
    d = request.json
    ok, resp = tg_send(d["token"], d["chat_id"],
                       tg_order_msg(d["symbol"], d["side"], d["qty"],
                                    d["entry"], d["sl"], d["tp"],
                                    d.get("pattern","—"), d.get("tf","—")))
    return jsonify({"ok": ok})

@app.route("/api/telegram/send_tsl", methods=["POST"])
def telegram_send_tsl():
    d = request.json
    ok, resp = tg_send(d["token"], d["chat_id"],
                       tg_tsl_msg(d["symbol"], d["status"],
                                  d["entry"], d["exit_price"], float(d["pnl"])))
    return jsonify({"ok": ok})

@app.route("/api/telegram/send_confluence", methods=["POST"])
def telegram_send_confluence():
    d = request.json
    ok, resp = tg_send(d["token"], d["chat_id"],
                       tg_confluence_msg(d["symbol"], d["direction"], d["timeframes"]))
    return jsonify({"ok": ok})

def tg_session_report_msg(data):
    """Build a session summary report message for Telegram."""
    signals      = data.get("signals", [])
    trades       = data.get("trades", [])
    duration_min = data.get("duration_min", 0)
    balance      = data.get("balance", 0)
    start_bal    = data.get("start_balance", 0)

    # Signal stats
    total_sigs   = len(signals)
    bull_sigs    = len([s for s in signals if s.get("direction")==1])
    bear_sigs    = total_sigs - bull_sigs
    avg_score    = round(sum(s.get("score",0) for s in signals)/total_sigs, 1) if total_sigs else 0
    strong_sigs  = len([s for s in signals if s.get("score",0)>=80])

    # Pattern breakdown
    pat_count = {}
    for s in signals:
        p = s.get("pattern","?")
        pat_count[p] = pat_count.get(p,0)+1
    top_pat = max(pat_count, key=pat_count.get) if pat_count else "—"

    # Trade stats
    closed  = [t for t in trades if t.get("status") in ["WIN","LOSS"]]
    wins    = [t for t in closed if t.get("status")=="WIN"]
    losses  = [t for t in closed if t.get("status")=="LOSS"]
    open_tr = [t for t in trades if t.get("status")=="OPEN"]
    total_pnl = sum(float(t.get("pnl",0) or 0) for t in closed)
    win_rate  = round(len(wins)/len(closed)*100,1) if closed else 0

    # Duration
    h = duration_min // 60
    m = duration_min % 60
    duration_str = f"{h}h {m}m" if h else f"{m}m"

    # P&L vs balance
    pnl_str = f"{'+' if total_pnl>=0 else ''}{total_pnl:.2f} USDT"
    bal_str  = f"{balance:.2f} USDT" if balance else "—"

    # Emoji for overall performance
    perf_emoji = "🔥" if win_rate>=70 else "✅" if win_rate>=50 else "⚠" if closed else "📊"

    lines = [
        f"{perf_emoji} <b>SESSION REPORT</b> — HarmonicBot",
        f"⏱ Duration: {duration_str}",
        f"",
        f"📊 <b>SIGNALS</b>",
        f"   Total:      {total_sigs}  (🐂 {bull_sigs} bull / 🐻 {bear_sigs} bear)",
        f"   Avg Score:  {avg_score}/100",
        f"   🔥 Strong:  {strong_sigs} signals (score ≥80)",
        f"   Top Pattern: {top_pat} ({pat_count.get(top_pat,0)}×)",
        f"",
    ]

    if closed:
        lines += [
            f"💼 <b>TRADES</b>",
            f"   Closed:    {len(closed)}  (✅ {len(wins)}W / ❌ {len(losses)}L)",
            f"   Win Rate:  {win_rate}%",
            f"   P&L:       {pnl_str}",
        ]
        if open_tr:
            lines.append(f"   Open:      {len(open_tr)} position(s) still running")
        lines.append(f"")
    elif open_tr:
        lines += [f"💼 <b>TRADES</b>", f"   {len(open_tr)} open position(s) — no closed yet", f""]

    lines += [
        f"💰 <b>BALANCE</b>",
        f"   Current: {bal_str}",
        f"",
        f"⏰ {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
    ]
    return "\n".join(lines)

@app.route("/api/telegram/send_report", methods=["POST"])
def telegram_send_report():
    d = request.json
    ok, resp = tg_send(d["token"], d["chat_id"],
                       tg_session_report_msg(d.get("data",{})))
    return jsonify({"ok": ok, "error": resp.get("description","") if not ok else ""})

if __name__=="__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") != "production"
    print("="*52)
    print(f"  HarmonicBot Pro — Backend  port {port}")
    print(f"  Password protection: {'ON' if BOT_PASSWORD else 'OFF'}")
    print("  Press CTRL+C to stop")
    print("="*52)
    app.run(debug=debug, host="0.0.0.0", port=port)
