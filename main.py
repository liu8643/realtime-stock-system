# v5.2.0 Phase4 Takeover 機構級交易引擎接管版
# 已整合：
# 1. 波浪理論結構化欄位 wave_stage / wave_score / wave_risk_flag
# 2. 費波南西位置欄位 fibo_position / fibo_score / fibo_risk_flag
# 3. Decision Layer：final_decision / execution_ready / decision_reason
# 4. UI / CSV / PDF / TXT 同步顯示波費與最終決策
# 5. RR 閘門與禁追風險優先級
# 6. Phase4：進場區判斷 / 倉位管理 / 波浪RR風險資金配置 / 驗收規則

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from functools import lru_cache
import pandas as pd
import yfinance as yf
import requests
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import os
import csv
import logging

APP_TITLE = "GTC 股票專業版看盤分析系統"
APP_VERSION = "v5.2.0-Phase4-Takeover-FINAL"
DECISION_MODEL_VERSION = "EXEC-P4-TAKEOVER-20260428"
AUTO_REFRESH_MS = 30000
DEFAULT_ACCOUNT_CAPITAL = 1000000
DEFAULT_RISK_PCT = 1.0
MIN_BUY_RR = 1.5
MIN_BUY_ALLOCATION_SCORE = 70

LOG_FILE = "gtc_phase4_decision.log"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)


def setup_pdf_font():
    candidates = [
        r"C:\Windows\Fonts\msjh.ttc",
        r"C:\Windows\Fonts\msjh.ttf",
        r"C:\Windows\Fonts\mingliu.ttc",
        r"C:\Windows\Fonts\kaiu.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("CH_FONT", path))
                return "CH_FONT"
            except Exception:
                pass
    return "Helvetica"

def normalize_symbol(symbol: str) -> list[str]:
    s = symbol.strip().upper()
    if not s:
        return []
    if "." in s:
        return [s]
    if s.isdigit():
        if len(s) == 4:
            return [f"{s}.TWO", f"{s}.TW"]
        return [s]
    return [s]

@lru_cache(maxsize=1)
def get_tw_name_map():
    mapping = {}
    sources = [
        "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
        "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O",
    ]
    for url in sources:
        try:
            df = pd.read_json(url)
            code_col = None
            name_col = None
            for c in df.columns:
                c_str = str(c).strip()
                if code_col is None and ("代號" in c_str or "Code" in c_str):
                    code_col = c
                if name_col is None and ("簡稱" in c_str or "名稱" in c_str or "Name" in c_str):
                    name_col = c
            if code_col is None or name_col is None:
                continue
            for _, row in df.iterrows():
                code = str(row[code_col]).strip()
                name = str(row[name_col]).strip()
                if code.isdigit() and len(code) == 4 and name:
                    mapping[code] = name
        except Exception:
            continue
    return mapping

def get_stock_name(input_symbol: str, yf_symbol: str) -> str:
    if input_symbol.isdigit() and len(input_symbol) == 4:
        tw_map = get_tw_name_map()
        if input_symbol in tw_map:
            return tw_map[input_symbol]
    try:
        ticker = yf.Ticker(yf_symbol)
        info = ticker.info
        name = info.get("shortName") or info.get("longName")
        if name:
            return str(name)
    except Exception:
        pass
    return yf_symbol

def download_symbol_data(symbol: str, period: str = "12mo") -> tuple[str, pd.DataFrame]:
    candidates = normalize_symbol(symbol)
    last_error = None
    for yf_symbol in candidates:
        try:
            df = yf.download(
                yf_symbol,
                period=period,
                interval="1d",
                auto_adjust=False,
                progress=False,
                threads=False,
            )
            if df is None or df.empty:
                continue
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [c[0] for c in df.columns]
            needed = ["Open", "High", "Low", "Close", "Volume"]
            if not all(c in df.columns for c in needed):
                continue
            df = df.dropna(subset=["Close"]).copy()
            if df.empty:
                continue
            return yf_symbol, df
        except Exception as e:
            last_error = e
    if last_error:
        raise ValueError(f"查無資料：{symbol} / {last_error}")
    raise ValueError(f"查無資料：{symbol}")

def round_price(v: float) -> float:
    return round(float(v), 2)

def safe_float(v, default=None):
    try:
        if v in (None, "", "-", "--"):
            return default
        return float(v)
    except Exception:
        return default

def safe_int(v, default=None):
    try:
        if v in (None, "", "-", "--"):
            return default
        return int(float(v))
    except Exception:
        return default

def split_prices(text):
    if not text:
        return []
    vals = []
    for x in str(text).split("_"):
        v = safe_float(x)
        if v is not None and v > 0:
            vals.append(round_price(v))
    return vals

def split_ints(text):
    if not text:
        return []
    vals = []
    for x in str(text).split("_"):
        v = safe_int(x)
        if v is not None and v >= 0:
            vals.append(v)
    return vals

def get_orderbook_bias(bid_vols, ask_vols):
    buy_qty = sum(bid_vols[:5]) if bid_vols else 0
    sell_qty = sum(ask_vols[:5]) if ask_vols else 0
    if buy_qty == 0 and sell_qty == 0:
        return {"buy_qty": 0, "sell_qty": 0, "ratio": "-", "bias": "無有效五檔"}
    if sell_qty == 0:
        return {"buy_qty": buy_qty, "sell_qty": sell_qty, "ratio": "∞", "bias": "買盤明顯偏強"}
    ratio = buy_qty / sell_qty
    if ratio >= 1.5:
        bias = "買盤偏強"
    elif ratio <= 0.67:
        bias = "賣盤偏強"
    else:
        bias = "多空均衡"
    return {"buy_qty": buy_qty, "sell_qty": sell_qty, "ratio": f"{ratio:.2f}", "bias": bias}

def detect_market(input_symbol: str, yf_symbol: str) -> str:
    if yf_symbol.endswith(".TW"):
        return "台股上市"
    if yf_symbol.endswith(".TWO"):
        return "台股上櫃"
    if input_symbol.isalpha():
        return "美股/海外"
    return "其他"

def get_tw_realtime_quote(symbol: str, market: str) -> dict | None:
    if market not in ("台股上市", "台股上櫃"):
        return None
    ex_prefix = "tse" if market == "台股上市" else "otc"
    ex_ch = f"{ex_prefix}_{symbol}.tw"
    url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
    params = {"ex_ch": ex_ch, "json": "1", "delay": "0", "_": str(int(datetime.now().timestamp() * 1000))}
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://mis.twse.com.tw/stock/index.jsp"}
    try:
        r = requests.get(url, params=params, headers=headers, timeout=8)
        r.raise_for_status()
        data = r.json()
        msg_array = data.get("msgArray", [])
        if not msg_array:
            return None
        item = msg_array[0]
        last_trade = safe_float(item.get("z"))
        open_price = safe_float(item.get("o"))
        high_price = safe_float(item.get("h"))
        low_price = safe_float(item.get("l"))
        prev_close = safe_float(item.get("y"))
        ask_prices = split_prices(item.get("a"))
        bid_prices = split_prices(item.get("b"))
        ask_vols = split_ints(item.get("f"))
        bid_vols = split_ints(item.get("g"))
        indicative_price = None
        if bid_prices and ask_prices:
            indicative_price = round_price((bid_prices[0] + ask_prices[0]) / 2)
        elif bid_prices:
            indicative_price = bid_prices[0]
        elif ask_prices:
            indicative_price = ask_prices[0]
        if last_trade is not None:
            display_price = round_price(last_trade)
            display_note = "即時成交價"
        elif indicative_price is not None:
            display_price = round_price(indicative_price)
            display_note = "當下無成交，改用買一/賣一中間價"
        elif prev_close is not None:
            display_price = round_price(prev_close)
            display_note = "當下無成交且無五檔，暫以昨收顯示"
        else:
            return None
        ob = get_orderbook_bias(bid_vols, ask_vols)
        return {
            "close": display_price,
            "display_price": display_price,
            "display_note": display_note,
            "last_trade": round_price(last_trade) if last_trade is not None else None,
            "indicative_price": round_price(indicative_price) if indicative_price is not None else None,
            "prev_close": round_price(prev_close if prev_close is not None else display_price),
            "open": round_price(open_price if open_price is not None else display_price),
            "high": round_price(high_price if high_price is not None else display_price),
            "low": round_price(low_price if low_price is not None else display_price),
            "bid_prices": bid_prices,
            "ask_prices": ask_prices,
            "bid_vols": bid_vols,
            "ask_vols": ask_vols,
            "buy_qty": ob["buy_qty"],
            "sell_qty": ob["sell_qty"],
            "orderbook_ratio": ob["ratio"],
            "orderbook_bias": ob["bias"],
            "quote_time": item.get("t") or item.get("tt") or "",
            "source": "TWSE MIS 即時",
        }
    except Exception:
        return None

def get_us_yahoo_quote(yf_symbol: str, fallback_close: float, fallback_prev_close: float, fallback_open: float, fallback_high: float, fallback_low: float) -> dict:
    live_price = fallback_close
    prev_close = fallback_prev_close
    open_price = fallback_open
    high_price = fallback_high
    low_price = fallback_low
    try:
        ticker = yf.Ticker(yf_symbol)
        try:
            fi = ticker.fast_info
            if fi:
                lp = fi.get("lastPrice")
                pc = fi.get("previousClose")
                day_high = fi.get("dayHigh")
                day_low = fi.get("dayLow")
                day_open = fi.get("open")
                if lp is not None:
                    live_price = round(float(lp), 2)
                if pc is not None:
                    prev_close = round(float(pc), 2)
                if day_high is not None:
                    high_price = round(float(day_high), 2)
                if day_low is not None:
                    low_price = round(float(day_low), 2)
                if day_open is not None:
                    open_price = round(float(day_open), 2)
        except Exception:
            pass
        try:
            info = ticker.info
            rp = info.get("regularMarketPrice")
            pcp = info.get("regularMarketPreviousClose")
            day_high = info.get("regularMarketDayHigh")
            day_low = info.get("regularMarketDayLow")
            day_open = info.get("regularMarketOpen")
            if rp is not None:
                live_price = round(float(rp), 2)
            if pcp is not None:
                prev_close = round(float(pcp), 2)
            if day_high is not None:
                high_price = round(float(day_high), 2)
            if day_low is not None:
                low_price = round(float(day_low), 2)
            if day_open is not None:
                open_price = round(float(day_open), 2)
        except Exception:
            pass
    except Exception:
        pass
    return {"close": live_price, "prev_close": prev_close, "open": open_price, "high": high_price, "low": low_price, "source": "Yahoo Finance"}

def calc_indicators(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["MA5"] = df["Close"].rolling(5).mean()
    df["MA10"] = df["Close"].rolling(10).mean()
    df["MA20"] = df["Close"].rolling(20).mean()
    df["MA60"] = df["Close"].rolling(60).mean()
    delta = df["Close"].diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(14).mean()
    avg_loss = loss.rolling(14).mean()
    rs = avg_gain / avg_loss.replace(0, pd.NA)
    df["RSI"] = 100 - (100 / (1 + rs))
    df["RSI"] = df["RSI"].fillna(50)
    ema12 = df["Close"].ewm(span=12, adjust=False).mean()
    ema26 = df["Close"].ewm(span=26, adjust=False).mean()
    df["MACD"] = ema12 - ema26
    df["MACD_SIGNAL"] = df["MACD"].ewm(span=9, adjust=False).mean()
    low9 = df["Low"].rolling(9).min()
    high9 = df["High"].rolling(9).max()
    rsv = (df["Close"] - low9) / (high9 - low9) * 100
    df["K"] = rsv.ewm(com=2).mean()
    df["D"] = df["K"].ewm(com=2).mean()
    return df

def calc_professional_sr(df: pd.DataFrame) -> dict:
    recent20 = df.tail(20)
    recent40 = df.tail(40)
    close = float(df["Close"].iloc[-1])
    support_20 = float(recent20["Low"].min())
    resistance_20 = float(recent20["High"].max())
    swing_low = float(recent40["Low"].min())
    swing_high = float(recent40["High"].max())
    last_bar = df.iloc[-1]
    pivot = (float(last_bar["High"]) + float(last_bar["Low"]) + float(last_bar["Close"])) / 3
    r1 = pivot * 2 - float(last_bar["Low"])
    s1 = pivot * 2 - float(last_bar["High"])
    support_candidates = [support_20, swing_low, s1]
    resistance_candidates = [resistance_20, swing_high, r1]
    supports_below = [x for x in support_candidates if x <= close]
    main_support = max(supports_below) if supports_below else min(support_candidates)
    resistances_above = [x for x in resistance_candidates if x >= close]
    main_resistance = min(resistances_above) if resistances_above else max(resistance_candidates)
    return {
        "support": round_price(main_support),
        "resistance": round_price(main_resistance),
        "support20": round_price(support_20),
        "resistance20": round_price(resistance_20),
        "swing_low": round_price(swing_low),
        "swing_high": round_price(swing_high),
        "pivot": round_price(pivot),
        "s1": round_price(s1),
        "r1": round_price(r1),
    }

def build_trade_advice(close, ma20, ma60, score, rsi, support, resistance, change_pct, intraday_score=0, open_price=None, prev_close=None, trend_score=0, orderbook_bias="無"):
    if change_pct <= -9.0:
        return "觀望為主"
    if close < support:
        return "減碼/防守"
    if close > resistance and trend_score >= 82 and intraday_score >= 78 and orderbook_bias in ("買盤偏強", "買盤明顯偏強"):
        return "突破追價"
    if score >= 95 and trend_score >= 90 and intraday_score >= 85:
        return "拉回加碼"
    if score >= 82 and trend_score >= 80 and intraday_score >= 70 and change_pct > 0.5:
        return "低接布局"
    if score >= 45 and support <= close <= resistance:
        return "區間操作"
    if rsi > 70:
        return "減碼/防守"
    if close < ma20 and close < ma60 and change_pct < 0:
        return "減碼/防守"
    return "觀望為主"


def classify_trade_type(state_bucket: str, signal: str, advice: str) -> str:
    if signal == "整理偏多":
        return "整理偏多"
    if signal == "突破強勢" or "突破可追" in advice:
        return "突破追價"
    if state_bucket == "strong":
        return "拉回承接"
    if state_bucket == "bullish":
        return "偏多低接"
    if state_bucket == "range":
        return "區間低接"
    return "觀望"


def build_risk_note(close, support, resistance, rsi, score, change_pct=None,
                    wave_stage="-", fibo_position="-", fibo_risk_flag=False,
                    wave_risk_flag=False, rr_valid=True, price_valid=True):
    notes = []
    if not price_valid:
        notes.append("報價非即時成交或為回退資料，不可直接下單")
    if fibo_risk_flag:
        notes.append(f"費波位置為「{fibo_position}」，觸發禁追風險")
    if wave_risk_flag:
        notes.append(f"波浪定位為「{wave_stage}」，需防末升或修正風險")
    if not rr_valid:
        notes.append("RR未達有效門檻，買進條件不足")
    if change_pct is not None and change_pct <= -7:
        notes.append("當日跌幅偏大，短線波動風險升高")
    if change_pct is not None and change_pct <= -9:
        notes.append("接近或達跌停級別，避免把急跌誤判為強勢買點")
    if close <= support * 1.01:
        notes.append("接近支撐，觀察是否守穩")
    if close < support:
        notes.append("已跌破支撐，需提高風險控管")
    if close >= resistance * 0.99:
        notes.append("逼近壓力，留意獲利了結賣壓")
    if close > resistance:
        notes.append("已突破壓力，觀察是否假突破")
    if rsi >= 70:
        notes.append("RSI 偏高，短線過熱風險上升")
    if rsi <= 30:
        notes.append("RSI 偏低，可能進入超跌區")
    if score < 30:
        notes.append("綜合評分偏弱，不宜積極追價")
    if not notes:
        notes.append("目前技術面無明顯異常，但仍須控管部位")
    return "；".join(notes)

def build_ai_analysis(data: dict) -> str:
    close = data["close"]
    ma20 = data["ma20"]
    ma60 = data["ma60"]
    rsi = data["rsi"]
    score = data["score"]
    trend_score = data.get("trend_score", score)
    intraday_score = data.get("intraday_score", score)
    support = data["support"]
    resistance = data["resistance"]
    signal = data["signal"]
    advice = data["advice"]
    orderbook_bias = data.get("orderbook_bias", "無")
    orderbook_ratio = data.get("orderbook_ratio", "-")
    change_pct = data.get("change_pct", 0.0)
    if close >= ma20 and close >= ma60:
        trend_text = "目前股價位於20日線與60日線之上，中期趨勢偏強。"
        trend = "偏多"
    elif close >= ma20 and close < ma60:
        trend_text = "目前股價站上20日線，但仍在60日線下方，屬短強中性結構。"
        trend = "盤整偏多"
    elif close < ma20 and close >= ma60:
        trend_text = "目前股價跌破20日線但仍守住60日線，短線轉弱、中期待觀察。"
        trend = "盤整偏弱"
    else:
        trend_text = "目前股價位於20日線與60日線下方，技術面偏弱。"
        trend = "偏空"
    if close < support:
        pos_text = f"目前股價 {close} 已跌破支撐 {support}，位置偏弱。"
    elif close > resistance:
        pos_text = f"目前股價 {close} 已突破壓力 {resistance}，位置轉強。"
    else:
        pos_text = f"目前股價位於支撐 {support} 與壓力 {resistance} 之間，仍屬區間內。"
    if rsi >= 70:
        rsi_text = f"RSI為 {rsi}，已接近或進入過熱區，短線需留意震盪與拉回。"
    elif rsi <= 30:
        rsi_text = f"RSI為 {rsi}，已進入相對低檔區，若量價配合有機會出現反彈。"
    elif rsi >= 55:
        rsi_text = f"RSI為 {rsi}，動能偏強，但仍需觀察是否能持續放大。"
    elif rsi >= 40:
        rsi_text = f"RSI為 {rsi}，動能中性偏弱，屬整理觀察區。"
    else:
        rsi_text = f"RSI為 {rsi}，動能偏弱，短線仍需保守。"
    ob_text = f"五檔力道為「{orderbook_bias}」，委買/委賣比為 {orderbook_ratio}。"
    if score >= 80:
        score_text = "綜合評分屬高分區，結構偏強。"
    elif score >= 65:
        score_text = "綜合評分中上，偏多但仍需確認續航力。"
    elif score >= 45:
        score_text = "綜合評分中性，屬區間整理型。"
    else:
        score_text = "綜合評分偏弱，先以風險控制優先。"
    if change_pct <= -9:
        drop_text = f"當日跌幅 {change_pct:+.2f}% 已屬高風險急跌，不宜僅因均線與歷史分數誤判為強勢買點。"
    elif change_pct <= -5:
        drop_text = f"當日跌幅 {change_pct:+.2f}% 偏大，需提高風險意識。"
    elif change_pct >= 5:
        drop_text = f"當日漲幅 {change_pct:+.2f}% 偏強，需觀察是否放量續攻。"
    else:
        drop_text = f"當日漲跌幅 {change_pct:+.2f}% 屬正常波動區間。"
    final_text = f"AI綜合判斷：趨勢偏向「{trend}」，訊號為「{signal}」，建議採取「{advice}」策略。"
    return "\n".join([
        "【AI個股分析】",
        f"1. 趨勢判讀：{trend_text}",
        f"2. 位置判讀：{pos_text}",
        f"3. 動能狀態：{rsi_text}",
        f"4. 五檔力道：{ob_text}",
        f"5. 當日強弱：{drop_text}",
        f"6. 分數解讀：{score_text}（波段分={trend_score} / 盤中分={intraday_score} / 總分={score}）",
        f"7. AI結論：{final_text}",
    ])

def detect_local_pivots(series: pd.Series, left: int = 2, right: int = 2):
    pivots = []
    values = series.tolist()
    for i in range(left, len(values) - right):
        window = values[i - left:i + right + 1]
        center = values[i]
        if center == max(window):
            pivots.append((i, "H", float(center)))
        elif center == min(window):
            pivots.append((i, "L", float(center)))
    return pivots

def summarize_wave(df: pd.DataFrame, period: int, label: str) -> str:
    part = df.tail(period).copy()
    if len(part) < 15:
        return f"{label}：資料不足，暫無法判讀。"
    close_start = float(part["Close"].iloc[0])
    close_end = float(part["Close"].iloc[-1])
    highest = float(part["High"].max())
    lowest = float(part["Low"].min())
    amplitude_pct = ((highest - lowest) / lowest * 100) if lowest != 0 else 0
    ma20_last = float(part["Close"].rolling(20).mean().iloc[-1]) if len(part) >= 20 else close_end
    ma60_last = float(part["Close"].rolling(60).mean().iloc[-1]) if len(part) >= 60 else close_end
    pivots = detect_local_pivots(part["Close"], left=2, right=2)
    recent_pivots = pivots[-6:] if len(pivots) >= 6 else pivots
    if close_end > close_start and close_end >= ma20_last:
        if len(recent_pivots) >= 5:
            wave_hint = "較偏推動浪結構，可能處於第3浪或第5浪延伸區。"
        else:
            wave_hint = "偏多推升結構，可能處於推動浪初升段。"
    elif close_end < close_start and close_end < ma20_last:
        if len(recent_pivots) >= 4:
            wave_hint = "較偏修正浪結構，可能位於 A / C 浪下修階段。"
        else:
            wave_hint = "偏弱修正結構，較像回檔整理波。"
    else:
        wave_hint = "目前較像整理浪或轉折確認階段，尚未形成明確單邊波段。"
    if close_end >= ma20_last and close_end >= ma60_last:
        trend_hint = "均線結構偏多。"
    elif close_end >= ma20_last and close_end < ma60_last:
        trend_hint = "短線偏強，但中期壓力仍在。"
    elif close_end < ma20_last and close_end >= ma60_last:
        trend_hint = "短線轉弱，中期尚未完全破壞。"
    else:
        trend_hint = "短中期均線結構偏弱。"
    return f"{label}：區間波動約 {amplitude_pct:.2f}% ，{wave_hint}{trend_hint}"

def build_wave_analysis(df: pd.DataFrame) -> str:
    return "\n".join([
        "【波浪理論分析】",
        f"1. {summarize_wave(df, 20, '短期')}",
        f"2. {summarize_wave(df, 60, '中期')}",
        f"3. {summarize_wave(df, 120, '長期')}",
    ])

def calc_fibonacci_targets(df: pd.DataFrame) -> dict:
    lookback = df.tail(120).copy()
    if len(lookback) < 30:
        close_now = float(df["Close"].iloc[-1])
        return {
            "direction": "資料不足",
            "base_low": round_price(close_now),
            "base_high": round_price(close_now),
            "range": 0.0,
            "target_1_0": round_price(close_now),
            "target_1_382": round_price(close_now),
            "target_1_618": round_price(close_now),
            "next_target": round_price(close_now),
            "summary": "資料不足，暫無法估算費波南西目標位。",
        }
    close_now = float(lookback["Close"].iloc[-1])
    low_val = float(lookback["Low"].min())
    high_val = float(lookback["High"].max())
    price_range = high_val - low_val
    low_idx = lookback["Low"].idxmin()
    high_idx = lookback["High"].idxmax()
    upward = low_idx < high_idx
    if price_range <= 0:
        return {
            "direction": "整理",
            "base_low": round_price(low_val),
            "base_high": round_price(high_val),
            "range": round_price(price_range),
            "target_1_0": round_price(close_now),
            "target_1_382": round_price(close_now),
            "target_1_618": round_price(close_now),
            "next_target": round_price(close_now),
            "summary": "區間過小，暫不適合估算費波南西延伸目標。",
        }
    if upward:
        direction = "上升波"
        target_1_0 = high_val
        target_1_382 = low_val + price_range * 1.382
        target_1_618 = low_val + price_range * 1.618
        if close_now < target_1_0:
            next_target = target_1_0
        elif close_now < target_1_382:
            next_target = target_1_382
        else:
            next_target = target_1_618
        summary = f"目前較偏上升波段，近波段低點 {round_price(low_val)} 至高點 {round_price(high_val)}。若續強，下一觀察目標依序為 1.0={round_price(target_1_0)}、1.382={round_price(target_1_382)}、1.618={round_price(target_1_618)}。"
    else:
        direction = "下降波"
        target_1_0 = low_val
        target_1_382 = high_val - price_range * 1.382
        target_1_618 = high_val - price_range * 1.618
        if close_now > target_1_0:
            next_target = target_1_0
        elif close_now > target_1_382:
            next_target = target_1_382
        else:
            next_target = target_1_618
        summary = f"目前較偏下降修正波，近波段高點 {round_price(high_val)} 至低點 {round_price(low_val)}。若續弱，下一觀察目標依序為 1.0={round_price(target_1_0)}、1.382={round_price(target_1_382)}、1.618={round_price(target_1_618)}。"
    return {
        "direction": direction,
        "base_low": round_price(low_val),
        "base_high": round_price(high_val),
        "range": round_price(price_range),
        "target_1_0": round_price(target_1_0),
        "target_1_382": round_price(target_1_382),
        "target_1_618": round_price(target_1_618),
        "next_target": round_price(next_target),
        "summary": summary,
    }

def build_fibonacci_analysis(fibo: dict) -> str:
    return "\n".join([
        "【費波南西目標位】",
        f"1. 波段方向：{fibo['direction']}",
        f"2. 波段低點：{fibo['base_low']} / 波段高點：{fibo['base_high']}",
        f"3. 1.0 目標位：{fibo['target_1_0']}",
        f"4. 1.382 目標位：{fibo['target_1_382']}",
        f"5. 1.618 目標位：{fibo['target_1_618']}",
        f"6. 下一目標價：{fibo['next_target']}",
        f"7. 判讀：{fibo['summary']}",
    ])

def build_bull_bear_path(data: dict) -> str:
    support = data["support"]
    resistance = data["resistance"]
    next_target = data["fibo"]["next_target"]
    signal = data["signal"]
    advice = data["advice"]
    return "\n".join([
        "【多空路徑圖示】",
        "◎ 多方路徑：",
        f"→ 多方路徑①：守住支撐 {support}",
        f"→ 多方路徑②：重新挑戰壓力 {resistance}",
        f"→ 多方路徑③：若有效突破壓力，下一目標看 {next_target}",
        "",
        "◎ 空方路徑：",
        f"→ 空方路徑①：若跌破支撐 {support}",
        "→ 空方路徑②：短線結構轉弱，恐回測更低整理區",
        f"→ 空方路徑③：若反彈無法站回壓力 {resistance}，弱勢格局延續",
        "",
        f"【路徑結論】當前訊號為「{signal}」，操作建議為「{advice}」。",
    ])





def structured_wave_analysis(df: pd.DataFrame) -> dict:
    part = df.tail(120).copy()
    if len(part) < 30:
        return {
            "wave_stage": "資料不足",
            "wave_score": 0,
            "wave_reason": "日線資料少於30筆，暫不納入波浪判定。",
            "wave_risk_flag": False,
        }

    close_now = float(part["Close"].iloc[-1])
    close_start_20 = float(part["Close"].tail(20).iloc[0]) if len(part) >= 20 else float(part["Close"].iloc[0])
    ma20 = float(part["Close"].rolling(20).mean().iloc[-1]) if len(part) >= 20 else close_now
    ma60 = float(part["Close"].rolling(60).mean().iloc[-1]) if len(part) >= 60 else ma20
    rsi = float(df["RSI"].iloc[-1]) if "RSI" in df.columns and pd.notna(df["RSI"].iloc[-1]) else 50.0
    pivots = detect_local_pivots(part["Close"], left=2, right=2)
    recent_pivots = pivots[-6:]
    high_120 = float(part["High"].max())
    low_120 = float(part["Low"].min())
    range_pct = ((high_120 - low_120) / low_120 * 100) if low_120 else 0.0
    above_ma = close_now >= ma20 >= ma60
    below_ma20 = close_now < ma20
    momentum_up = close_now > close_start_20
    near_high = close_now >= high_120 * 0.94 if high_120 else False
    near_low = close_now <= low_120 * 1.08 if low_120 else False

    if above_ma and momentum_up and len(recent_pivots) >= 5 and 45 <= rsi <= 72:
        stage = "第3浪"
        score = 15
        reason = "站上MA20/MA60、20日動能向上且轉折點足夠，偏主升推動浪。"
        risk = False
    elif above_ma and momentum_up and (near_high or rsi > 72):
        stage = "第5浪"
        score = -8
        reason = "價格接近120日高點或RSI偏熱，偏末升延伸區。"
        risk = True
    elif below_ma20 and not momentum_up:
        stage = "A/C修正浪"
        score = -12
        reason = "跌破MA20且20日動能轉弱，偏修正浪。"
        risk = True
    elif near_low and rsi <= 40:
        stage = "第2浪/回測浪"
        score = 6
        reason = "接近波段低位且RSI偏低，偏回測觀察區。"
        risk = False
    elif close_now >= ma20:
        stage = "整理偏多"
        score = 5
        reason = "價格站上MA20但主升條件未完全成立，屬整理偏多。"
        risk = False
    else:
        stage = "整理/待確認"
        score = 0
        reason = f"波段振幅約{range_pct:.2f}%，尚未形成明確推動或修正結構。"
        risk = False

    return {
        "wave_stage": stage,
        "wave_score": int(score),
        "wave_reason": reason,
        "wave_risk_flag": bool(risk),
    }


def classify_fibo_position(close: float, fibo: dict) -> dict:
    direction = fibo.get("direction", "資料不足")
    t10 = safe_float(fibo.get("target_1_0"))
    t1382 = safe_float(fibo.get("target_1_382"))
    t1618 = safe_float(fibo.get("target_1_618"))

    if direction == "資料不足" or t10 is None or t1382 is None or t1618 is None:
        return {
            "fibo_position": "資料不足",
            "fibo_score": 0,
            "fibo_risk_flag": False,
            "fibo_reason": "費波南西資料不足，僅保留價格參考，不做交易升級。",
        }

    if direction == "下降波":
        if close <= t1382:
            return {
                "fibo_position": "下降延伸/破位",
                "fibo_score": -12,
                "fibo_risk_flag": True,
                "fibo_reason": "價格落在下降延伸區，優先風控。",
            }
        if close <= t10:
            return {
                "fibo_position": "跌破1.0",
                "fibo_score": -8,
                "fibo_risk_flag": True,
                "fibo_reason": "價格跌破下降波1.0目標，偏弱勢延續。",
            }
        return {
            "fibo_position": "下降波反彈區",
            "fibo_score": -2,
            "fibo_risk_flag": False,
            "fibo_reason": "下降波中反彈，需等待轉強確認。",
        }

    if close < t10 * 0.985:
        pos = "挑戰1.0前"
        score = 4
        risk = False
        reason = "尚未站上1.0目標，屬低接或突破前觀察區。"
    elif close < t10 * 1.015:
        pos = "站上/測試1.0"
        score = 10
        risk = False
        reason = "價格位於1.0目標附近，若量價配合可視為轉強確認。"
    elif close < t1382 * 0.985:
        pos = "1.0~1.382主升區"
        score = 12
        risk = False
        reason = "價格位於1.0與1.382之間，屬主升延伸有效區。"
    elif close < t1618 * 0.97:
        pos = "挑戰1.382"
        score = 8
        risk = False
        reason = "價格挑戰1.382延伸，仍可追蹤但需控管追價風險。"
    elif close <= t1618 * 1.02:
        pos = "接近1.618禁追區"
        score = -10
        risk = True
        reason = "價格接近1.618延伸目標，屬高檔禁追區。"
    else:
        pos = "突破1.618過熱區"
        score = -15
        risk = True
        reason = "價格已超過1.618延伸目標，追價風險過高。"

    return {
        "fibo_position": pos,
        "fibo_score": int(score),
        "fibo_risk_flag": bool(risk),
        "fibo_reason": reason,
    }


def build_wave_fibo_decision_note(result: dict) -> str:
    wave_stage = result.get("wave_stage", "-")
    fibo_position = result.get("fibo_position", "-")
    rr_valid = result.get("rr_valid", False)
    fibo_risk = result.get("fibo_risk_flag", False)
    wave_risk = result.get("wave_risk_flag", False)

    if fibo_risk or wave_risk:
        return f"禁追/風控：{wave_stage} + {fibo_position}，避免高檔追價。"
    if wave_stage == "第3浪" and fibo_position in ("站上/測試1.0", "1.0~1.382主升區", "挑戰1.382") and rr_valid:
        return "主升確認：第3浪 + 費波主升區 + RR有效，可拉回加碼。"
    if wave_stage in ("第2浪/回測浪", "整理偏多") and not fibo_risk:
        return f"低接觀察：{wave_stage} + {fibo_position}，等待支撐止穩。"
    if wave_stage == "A/C修正浪":
        return "修正風險：A/C修正浪，未重新站回關鍵均線前觀望。"
    return f"波費待確認：{wave_stage} + {fibo_position}，以原技術訊號為主。"



def classify_entry_zone(result: dict) -> dict:
    close = safe_float(result.get("close"), 0.0) or 0.0
    entry_low = safe_float(result.get("entry_low"), 0.0) or 0.0
    entry_high = safe_float(result.get("entry_high"), 0.0) or 0.0
    support = safe_float(result.get("support"), 0.0) or 0.0
    resistance = safe_float(result.get("resistance"), 0.0) or 0.0
    stop_loss = safe_float(result.get("stop_loss"), 0.0) or 0.0
    signal = result.get("signal", "")
    rr_valid = bool(result.get("rr_valid", False))
    fibo_risk = bool(result.get("fibo_risk_flag", False) or result.get("wave_risk_flag", False))
    state = result.get("state_bucket", "range")

    if close <= 0 or not result.get("trade_plan_valid", False):
        status = "NO_PLAN"
        ready = False
        reason = "交易計畫無效或價格不足，禁止下單。"
    elif fibo_risk:
        status = "NO_CHASE"
        ready = False
        reason = "波浪/費波觸發禁追，禁止追價。"
    elif (support > 0 and close < support) or (stop_loss > 0 and close < stop_loss):
        status = "BROKEN"
        ready = False
        reason = "跌破支撐或停損線，交易條件失效。"
    elif entry_low <= close <= entry_high:
        status = "IN_ZONE"
        ready = True
        reason = "目前價格位於建議進場區間內。"
    elif close < entry_low:
        status = "WAIT_PULLBACK"
        ready = False
        reason = "價格低於進場區，等待止穩或回到有效區間。"
    elif resistance > 0 and close > resistance and signal in ("主升突破", "突破強勢") and rr_valid and state == "strong":
        status = "BREAKOUT_CONFIRM"
        ready = True
        reason = "價格突破壓力且主升/突破訊號成立，允許小倉突破確認。"
    elif close > entry_high:
        status = "ABOVE_ENTRY"
        ready = False
        reason = "價格高於建議進場區，不追價，等待回測。"
    else:
        status = "WAIT_CONFIRM"
        ready = False
        reason = "尚未符合進場條件，等待確認。"

    if entry_low > 0 and close > 0:
        if close < entry_low:
            distance = (entry_low - close) / entry_low * 100
        elif close > entry_high and entry_high > 0:
            distance = (close - entry_high) / entry_high * 100
        else:
            distance = 0.0
    else:
        distance = 0.0

    chase_risk = bool(status in ("ABOVE_ENTRY", "NO_CHASE") or (resistance > 0 and close >= resistance * 0.99))
    if status in ("NO_CHASE", "BROKEN", "NO_PLAN"):
        order_type = "禁止"
    elif status == "IN_ZONE":
        order_type = "低接限價"
    elif status == "BREAKOUT_CONFIRM":
        order_type = "突破小倉"
    else:
        order_type = "等待"

    return {
        "entry_zone_status": status,
        "entry_zone_ready": bool(ready),
        "entry_zone_reason": reason,
        "distance_to_entry_pct": round(distance, 2),
        "chase_risk_flag": chase_risk,
        "order_type_hint": order_type,
    }


def calc_wave_rr_risk_allocation(result: dict) -> dict:
    rr = safe_float(result.get("rr"), 0.0) or 0.0
    wave_stage = result.get("wave_stage", "-")
    state = result.get("state_bucket", "range")
    entry_zone_status = result.get("entry_zone_status", "NO_PLAN")
    fibo_risk = bool(result.get("fibo_risk_flag", False) or result.get("wave_risk_flag", False))
    price_valid = bool(result.get("price_valid", False))
    signal = result.get("signal", "")

    if not price_valid:
        return {
            "allocation_score": 0,
            "allocation_grade": "BLOCK",
            "allocation_multiplier": 0.0,
            "phase4_block_reason": "報價非即時有效，資金配置歸零。",
        }
    if fibo_risk or wave_stage in ("第5浪", "A/C修正浪") or entry_zone_status in ("NO_CHASE", "BROKEN", "NO_PLAN"):
        return {
            "allocation_score": 0,
            "allocation_grade": "BLOCK",
            "allocation_multiplier": 0.0,
            "phase4_block_reason": f"波浪/費波或進場區觸發阻擋：{wave_stage}/{entry_zone_status}。",
        }
    if rr < 1.0:
        return {
            "allocation_score": 0,
            "allocation_grade": "BLOCK",
            "allocation_multiplier": 0.0,
            "phase4_block_reason": "RR小於1，資金配置歸零。",
        }

    score = 50
    if wave_stage == "第3浪":
        score += 25
    elif wave_stage in ("整理偏多", "第2浪/回測浪"):
        score += 10
    else:
        score += 0

    if rr >= 2.0:
        score += 20
    elif rr >= MIN_BUY_RR:
        score += 12
    elif rr >= 1.0:
        score += 3

    if entry_zone_status == "IN_ZONE":
        score += 15
    elif entry_zone_status == "BREAKOUT_CONFIRM":
        score += 5
    elif entry_zone_status == "ABOVE_ENTRY":
        score -= 20

    if state == "strong" or signal == "主升突破":
        score += 8
    elif state == "bullish":
        score += 3
    elif state in ("weak", "range"):
        score -= 8

    score = max(0, min(100, int(score)))
    if score >= 85:
        grade = "A"
        multiplier = 1.0
        block = ""
    elif score >= 70:
        grade = "B"
        multiplier = 0.65
        block = ""
    elif score >= 55:
        grade = "C"
        multiplier = 0.35
        block = "配置分未達主攻，僅允許小倉觀察。"
    else:
        grade = "D"
        multiplier = 0.0
        block = "配置分低於55，禁止下單。"

    return {
        "allocation_score": score,
        "allocation_grade": grade,
        "allocation_multiplier": round(multiplier, 2),
        "phase4_block_reason": block,
    }


def calc_position_sizing(result: dict, account_capital=DEFAULT_ACCOUNT_CAPITAL, risk_pct=DEFAULT_RISK_PCT) -> dict:
    close = safe_float(result.get("close"), 0.0) or 0.0
    entry_high = safe_float(result.get("entry_high"), 0.0) or 0.0
    stop_loss = safe_float(result.get("stop_loss"), 0.0) or 0.0
    rr = safe_float(result.get("rr"), 0.0) or 0.0
    wave_stage = result.get("wave_stage", "-")
    entry_zone_status = result.get("entry_zone_status", "NO_PLAN")
    allocation_multiplier = safe_float(result.get("allocation_multiplier"), 0.0) or 0.0
    allocation_grade = result.get("allocation_grade", "BLOCK")

    base_pct = 0.0
    if wave_stage == "第3浪" and rr >= 2.0 and entry_zone_status == "IN_ZONE":
        base_pct = 10.0
    elif wave_stage == "第3浪" and entry_zone_status == "BREAKOUT_CONFIRM":
        base_pct = 5.0
    elif wave_stage == "整理偏多" and rr >= 1.5 and entry_zone_status == "IN_ZONE":
        base_pct = 4.0
    elif rr >= 1.5 and entry_zone_status == "IN_ZONE":
        base_pct = 3.0

    if allocation_grade in ("BLOCK", "D") or allocation_multiplier <= 0:
        position_pct = 0.0
    else:
        position_pct = round(base_pct * allocation_multiplier, 2)

    # 限制最大單檔曝險與風險預算
    max_loss_pct = 0.0
    if entry_high > 0 and stop_loss > 0 and entry_high > stop_loss:
        max_loss_pct = round((entry_high - stop_loss) / entry_high * 100, 2)

    risk_budget_pct = round(min(float(risk_pct), position_pct * max_loss_pct / 100), 2) if position_pct > 0 else 0.0
    suggested_capital = account_capital * position_pct / 100
    suggested_shares = int(suggested_capital // close) if close > 0 and position_pct > 0 else 0
    if suggested_shares > 0:
        suggested_shares = (suggested_shares // 1000) * 1000 if close < 500 else (suggested_shares // 100) * 100

    return {
        "position_size_pct": position_pct,
        "risk_budget_pct": risk_budget_pct,
        "suggested_shares": suggested_shares,
        "max_loss_pct": max_loss_pct,
    }


def ensure_phase4_fields(result: dict) -> dict:
    """Phase4欄位完整性防呆：避免UI/CSV/PDF因缺欄造成KeyError或錯誤下單。"""
    defaults = {
        "entry_zone_status": "NO_PLAN",
        "entry_zone_ready": False,
        "entry_zone_reason": "Phase4欄位未完整產生，系統降級等待。",
        "distance_to_entry_pct": 0.0,
        "chase_risk_flag": False,
        "allocation_score": 0,
        "allocation_grade": "BLOCK",
        "allocation_multiplier": 0.0,
        "position_size_pct": 0.0,
        "risk_budget_pct": 0.0,
        "suggested_shares": 0,
        "max_loss_pct": 0.0,
        "phase4_block_reason": "Phase4欄位未完整產生，禁止下單。",
        "order_type_hint": "等待",
        "final_decision": "WAIT",
        "execution_ready": False,
        "decision_reason": "Phase4欄位未完整產生，系統降級等待。",
        "final_block_reason": "Phase4欄位未完整產生，系統降級等待。",
        "display_advice": result.get("advice", "觀望為主"),
        "display_trade_type": result.get("trade_type", "觀望"),
    }
    missing = []
    for key, value in defaults.items():
        if key not in result or result.get(key) is None:
            result[key] = value
            missing.append(key)
    if missing:
        logging.warning("PHASE4_MISSING_FIELDS symbol=%s missing=%s", result.get("input_symbol", "-"), ",".join(missing))
    return result


def sync_display_semantics(result: dict) -> dict:
    """用final_decision統一UI/CSV/PDF的人類語義，避免強勢追蹤但不可下單的誤判。"""
    decision = result.get("final_decision", "WAIT")
    entry_status = result.get("entry_zone_status", "-")
    block_reason = result.get("final_block_reason") or result.get("phase4_block_reason") or result.get("decision_reason", "")
    raw_advice = result.get("advice", "觀望為主")
    raw_trade_type = result.get("trade_type", "觀望")

    if decision == "BUY":
        result["display_advice"] = raw_advice
        result["display_trade_type"] = result.get("order_type_hint") or raw_trade_type
    elif decision == "AVOID":
        result["display_advice"] = "禁止進場/風控"
        result["display_trade_type"] = "禁止"
        result["final_block_reason"] = block_reason or "Phase4風控條件未通過，禁止下單。"
    else:
        if entry_status == "ABOVE_ENTRY":
            result["display_advice"] = "等待回測/不追價"
        elif entry_status == "WAIT_PULLBACK":
            result["display_advice"] = "等待止穩/僅觀察"
        elif result.get("allocation_score", 0) < MIN_BUY_ALLOCATION_SCORE:
            result["display_advice"] = "等待條件改善"
        else:
            result["display_advice"] = "等待確認/僅觀察"
        result["display_trade_type"] = "等待"
        result["final_block_reason"] = block_reason or result.get("decision_reason", "未符合可下單條件。")
    return result


def log_decision_trace(result: dict) -> None:
    """記錄每檔股票Phase4 Gate，方便EXE問題追蹤。"""
    try:
        logging.info(
            "DECISION symbol=%s price=%s source=%s entry=%s ready_entry=%s rr=%s rr_valid=%s alloc=%s grade=%s position=%s decision=%s execution_ready=%s order=%s reason=%s",
            result.get("input_symbol"), result.get("close"), result.get("source"),
            result.get("entry_zone_status"), result.get("entry_zone_ready"),
            result.get("rr"), result.get("rr_valid"), result.get("allocation_score"),
            result.get("allocation_grade"), result.get("position_size_pct"),
            result.get("final_decision"), result.get("execution_ready"),
            result.get("order_type_hint"), result.get("decision_reason")
        )
    except Exception as e:
        logging.warning("DECISION_LOG_FAILED symbol=%s error=%s", result.get("input_symbol", "-"), e)

def build_final_decision(result: dict) -> dict:
    price_valid = bool(result.get("price_valid", False))
    signal = result.get("signal", "")
    advice = result.get("advice", "")
    state = result.get("state_bucket", "range")
    rr = safe_float(result.get("rr"), None)
    rr_valid = bool(result.get("rr_valid", False))
    fibo_risk = bool(result.get("fibo_risk_flag", False))
    wave_risk = bool(result.get("wave_risk_flag", False))
    entry_zone_ready = bool(result.get("entry_zone_ready", False))
    entry_zone_status = result.get("entry_zone_status", "NO_PLAN")
    position_size_pct = safe_float(result.get("position_size_pct"), 0.0) or 0.0
    allocation_score = safe_float(result.get("allocation_score"), 0.0) or 0.0
    allocation_grade = result.get("allocation_grade", "BLOCK")
    phase4_block_reason = result.get("phase4_block_reason", "")
    entry_zone_reason = result.get("entry_zone_reason", "不符合進場區")
    order_type_hint = result.get("order_type_hint", "等待")
    support_broken = signal in ("急跌風險", "跌破支撐", "轉弱警戒")

    if not price_valid:
        decision = "WAIT"
        ready = False
        reason = "報價非即時成交或為回退資料，禁止下單。"
        order_type_hint = "等待"
    elif support_broken:
        decision = "AVOID"
        ready = False
        reason = f"命中風險訊號：{signal}。"
        order_type_hint = "禁止"
    elif fibo_risk or wave_risk:
        decision = "AVOID"
        ready = False
        reason = "波浪/費波觸發禁追或末升風險，不允許追價。"
        order_type_hint = "禁止"
    elif allocation_grade == "BLOCK":
        # Phase4硬Gate：資金等級BLOCK必須直接阻擋，避免被後續WAIT覆蓋成可觀察語義。
        decision = "AVOID"
        ready = False
        reason = phase4_block_reason or "Phase4資金等級為BLOCK，禁止下單。"
        order_type_hint = "禁止"
    elif entry_zone_status in ("ABOVE_ENTRY", "NO_CHASE", "BROKEN", "NO_PLAN"):
        decision = "AVOID" if entry_zone_status in ("NO_CHASE", "BROKEN") else "WAIT"
        ready = False
        reason = f"Phase4進場區阻擋：{entry_zone_reason}。"
        if decision == "AVOID":
            order_type_hint = "禁止"
        else:
            order_type_hint = "等待"
    elif not entry_zone_ready:
        decision = "WAIT"
        ready = False
        reason = f"尚未進入可執行進場區：{entry_zone_reason}。"
        order_type_hint = "等待"
    elif rr is None or rr < 1:
        decision = "WAIT"
        ready = False
        reason = "RR小於1或無法計算，禁止買進。"
        order_type_hint = "等待"
    elif rr < MIN_BUY_RR:
        decision = "WAIT"
        ready = False
        reason = f"RR介於1.0~{MIN_BUY_RR}，僅觀察或等待更佳風險報酬。"
        order_type_hint = "等待"
    elif position_size_pct <= 0:
        decision = "WAIT"
        ready = False
        reason = phase4_block_reason or "倉位計算為0，禁止下單。"
        order_type_hint = "等待"
    elif allocation_score < MIN_BUY_ALLOCATION_SCORE:
        decision = "WAIT"
        ready = False
        reason = phase4_block_reason or f"配置分低於{MIN_BUY_ALLOCATION_SCORE}，等待更佳進場條件。"
        order_type_hint = "等待"
    elif state == "strong" and rr_valid and entry_zone_ready and advice in ("突破可追", "拉回加碼"):
        decision = "BUY"
        ready = True
        reason = f"Phase4通過：進場區有效、RR有效、配置分={allocation_score}、建議倉位={position_size_pct}%。"
    elif state == "bullish" and rr_valid:
        decision = "WAIT"
        ready = False
        reason = "偏多但未達強勢買進，等待拉回或突破確認。"
        order_type_hint = "等待"
    else:
        decision = "WAIT"
        ready = False
        reason = "未符合可下單條件。"
        order_type_hint = "等待"

    # 後置Gate保護：未來維護時即使上方誤判BUY，也不能繞過Phase4硬條件。
    gate_ready = (
        decision == "BUY" and price_valid and entry_zone_ready and
        position_size_pct > 0 and allocation_score >= MIN_BUY_ALLOCATION_SCORE and
        rr_valid and rr is not None and rr >= MIN_BUY_RR and
        allocation_grade not in ("BLOCK", "D") and
        entry_zone_status in ("IN_ZONE", "BREAKOUT_CONFIRM") and
        not (fibo_risk or wave_risk or support_broken)
    )
    if decision == "BUY" and not gate_ready:
        decision = "WAIT"
        ready = False
        reason = "Phase4後置Gate未通過，降級等待。"
        order_type_hint = "等待"
    else:
        ready = bool(gate_ready)

    if decision != "BUY":
        position_size_pct = 0.0

    final_block_reason = "" if decision == "BUY" else (phase4_block_reason or reason)

    return {
        "final_decision": decision,
        "execution_ready": ready,
        "decision_reason": reason,
        "order_type_hint": order_type_hint,
        "position_size_pct": position_size_pct,
        "final_block_reason": final_block_reason,
    }

def get_light(signal, score, change_pct, intraday_score=None, fibo_risk_flag=False, wave_risk_flag=False, final_decision=None):
    intraday_score = intraday_score or 0
    if final_decision == "BUY":
        return "🔵"
    if signal == "急跌風險" or change_pct <= -9.0:
        return "🔴"
    if signal in ("跌破支撐", "轉弱警戒") or final_decision == "AVOID":
        return "🟠"
    if fibo_risk_flag or wave_risk_flag:
        return "🟠"
    if signal == "突破強勢":
        return "🔵"
    if signal in ("偏多觀察", "強勢追蹤", "主升突破"):
        return "🟢"
    if signal == "區間整理":
        return "🟡"
    if score >= 45 or intraday_score >= 45:
        return "🟡"
    return "🟠"


def evaluate_trade_state(close, prev_close, open_price, support, resistance, change_pct,
                         trend_score, intraday_score, score, orderbook_bias, ma20=0, ma60=0, rsi=50,
                         wave_stage="-", fibo_position="-", fibo_risk_flag=False,
                         wave_risk_flag=False, rr_valid=False, price_valid=True):
    near_resistance = close >= resistance * 0.988 if resistance else False
    at_breakout = close >= resistance * 0.998 if resistance else False
    above_open = close >= open_price
    above_prev = close >= prev_close
    bullish_orderbook = orderbook_bias in ("買盤偏強", "買盤明顯偏強")
    structure_bullish = (close >= ma20 and close >= ma60 and ma20 >= ma60) if ma20 and ma60 else False

    if not price_valid:
        return "資料待確認", "觀望為主", "weak", "D00", "非即時成交或日線回退，不允許下單"

    if change_pct <= -9.0 or intraday_score <= 15:
        return "急跌風險", "觀望為主", "weak", "R01", "當日急跌或盤中分過低，先處理風險"

    if close < support * 0.997 or (close < support and intraday_score < 42):
        return "跌破支撐", "減碼/防守", "weak", "R02", "跌破主支撐且盤中力道不足"

    if fibo_risk_flag or wave_risk_flag:
        return "末升/禁追風險", "不追高", "weak", "WF_RISK", "波浪或費波觸發末升/禁追條件"

    if (
        wave_stage == "第3浪" and
        fibo_position in ("站上/測試1.0", "1.0~1.382主升區", "挑戰1.382") and
        rr_valid and trend_score >= 80 and intraday_score >= 70 and score >= 78
    ):
        return "主升突破", "拉回加碼", "strong", "WF_BUY", "第3浪+費波主升區+RR有效"

    if score >= 95 and trend_score >= 90 and intraday_score >= 85 and rr_valid:
        if at_breakout and bullish_orderbook:
            return "突破強勢", "突破可追", "strong", "S01", "高分突破且五檔買盤偏強"
        return "強勢追蹤", "拉回加碼", "strong", "S02", "高分強勢但未完成有效突破"

    if (
        close > resistance and trend_score >= 82 and intraday_score >= 78 and score >= 86 and
        change_pct >= 1.8 and above_open and above_prev and bullish_orderbook and rr_valid
    ):
        return "突破強勢", "突破可追", "strong", "S03", "突破壓力、量價轉強且RR有效"

    if (
        trend_score >= 82 and intraday_score >= 70 and score >= 82 and
        change_pct >= 0.8 and above_open and above_prev and bullish_orderbook and rr_valid
    ):
        return "強勢追蹤", "拉回加碼", "strong", "S04", "波段與盤中分數同步偏強"

    if (
        trend_score >= 80 and intraday_score >= 70 and score >= 75 and structure_bullish
    ):
        return "整理偏多", "低接布局", "bullish", "B01", "均線結構偏多但未達可追條件"

    if (
        trend_score >= 82 and intraday_score >= 68 and score >= 78 and
        structure_bullish and change_pct >= 1.5 and
        orderbook_bias in ("買盤偏強", "買盤明顯偏強") and 35 <= rsi <= 68 and
        (not resistance or close <= resistance * 1.03)
    ):
        return "整理偏多", "低接布局", "bullish", "B02", "站穩中期均線且五檔偏多，但未完成有效突破"

    if (
        trend_score >= 72 and intraday_score >= 58 and score >= 70 and
        change_pct >= 0.3 and (above_open or structure_bullish)
    ):
        return "偏多觀察", "低接布局", "bullish", "B03", "偏多觀察，但仍需等待確認"

    if score >= 45 and support <= close <= resistance:
        return "區間整理", "區間操作", "range", "N01", "價格位於支撐與壓力間，屬區間操作"

    if score >= 30:
        return "轉弱警戒", "減碼/防守", "weak", "W01", "分數不足且結構轉弱"

    return "轉弱警戒", "減碼/防守", "weak", "W02", "綜合條件偏弱"


def is_main_trend_candidate(data: dict) -> bool:
    close = data.get("close", 0)
    open_price = data.get("open", 0)
    prev_close = data.get("prev_close", 0)
    resistance = data.get("resistance", 0)
    trend = data.get("trend_score", 0)
    intra = data.get("intraday_score", 0)
    score = data.get("score", 0)
    rsi = data.get("rsi", 0)
    ma20 = data.get("ma20", 0)
    ma60 = data.get("ma60", 0)
    signal = data.get("signal", "")
    orderbook = data.get("orderbook_bias", "無")
    change_pct = data.get("change_pct", 0)
    wave_stage = data.get("wave_stage", "-")
    fibo_position = data.get("fibo_position", "-")
    rr_valid = bool(data.get("rr_valid", False))
    fibo_risk = bool(data.get("fibo_risk_flag", False) or data.get("wave_risk_flag", False))

    bullish_orderbook = orderbook in ("買盤偏強", "買盤明顯偏強", "多空均衡")
    not_too_far_from_resistance = close <= resistance * 1.01 if resistance else True
    healthy_strength = signal in ("強勢追蹤", "突破強勢", "偏多觀察", "主升突破")
    wave_fibo_ok = (
        wave_stage == "第3浪" and
        fibo_position in ("站上/測試1.0", "1.0~1.382主升區", "挑戰1.382")
    )

    return (
        not fibo_risk and rr_valid and
        score >= 85 and trend >= 80 and intra >= 70 and
        45 <= rsi <= 72 and close > ma20 >= ma60 and
        close >= open_price and close >= prev_close and
        change_pct >= 0.5 and bullish_orderbook and healthy_strength and
        not_too_far_from_resistance and wave_fibo_ok
    )


def classify_leader_stage(data: dict) -> str:
    if bool(data.get("fibo_risk_flag", False) or data.get("wave_risk_flag", False)):
        return "末升/禁追"

    wave_stage = data.get("wave_stage", "-")
    fibo_position = data.get("fibo_position", "-")
    rr_valid = bool(data.get("rr_valid", False))

    if is_main_trend_candidate(data):
        return "是"

    close = data.get("close", 0)
    ma20 = data.get("ma20", 0)
    ma60 = data.get("ma60", 0)
    resistance = data.get("resistance", 0)
    trend = data.get("trend_score", 0)
    intra = data.get("intraday_score", 0)
    score = data.get("score", 0)
    rsi = data.get("rsi", 0)
    signal = data.get("signal", "")
    orderbook = data.get("orderbook_bias", "無")

    if (
        score >= 80 and trend >= 75 and intra >= 65 and
        close > ma20 >= ma60 and 42 <= rsi <= 72 and
        signal in ("強勢追蹤", "突破強勢", "偏多觀察", "整理偏多", "主升突破") and
        orderbook != "賣盤偏強" and close <= resistance * 1.01 and
        wave_stage in ("第3浪", "整理偏多", "第2浪/回測浪") and not rr_valid
    ):
        return "觀察"

    if wave_stage == "第3浪" and fibo_position in ("站上/測試1.0", "1.0~1.382主升區"):
        return "觀察"

    return "-"

def get_strategy_level(score: int) -> str:
    if score >= 85:
        return "A"
    if score >= 70:
        return "B"
    if score >= 55:
        return "C"
    return "D"

def get_strategy_level_score(level: str) -> int:
    mapping = {"A": 4, "B": 3, "C": 2, "D": 1}
    return mapping.get(str(level).strip().upper(), 0)


def normalize_rr_display(rr):
    return "-" if rr is None else rr


def get_display_target(target, signal: str, state_bucket: str):
    if signal in ("轉弱警戒", "急跌風險", "跌破支撐", "區間整理") or state_bucket in ("weak", "range"):
        return "-"
    return target



def calc_trade_plan(data: dict) -> dict:
    support = float(data.get("support", 0) or 0)
    resistance = float(data.get("resistance", 0) or 0)
    fibo_target = float(data.get("fibo", {}).get("next_target", resistance) or resistance)
    state = data.get("state_bucket", "range")
    fibo_risk = bool(data.get("fibo_risk_flag", False) or data.get("wave_risk_flag", False))

    if state == "strong" and not fibo_risk:
        entry_low = support * 1.002
        entry_high = min(support * 1.012, resistance * 0.995) if resistance > 0 else support * 1.012
        stop = support * 0.982
    elif state == "bullish" and not fibo_risk:
        entry_low = support * 1.000
        entry_high = min(support * 1.010, resistance * 0.992) if resistance > 0 else support * 1.010
        stop = support * 0.978
    elif state == "range" and not fibo_risk:
        entry_low = support * 0.998
        entry_high = min(support * 1.006, resistance * 0.988) if resistance > 0 else support * 1.006
        stop = support * 0.972
    else:
        entry_low = 0.0
        entry_high = 0.0
        stop = support * 0.968 if support else 0.0

    target = max(resistance, fibo_target) if state in ("strong", "bullish") else resistance
    target_source = "Fibo/壓力擇高" if state in ("strong", "bullish") else "壓力"

    entry_mid = (entry_low + entry_high) / 2 if entry_low > 0 and entry_high > 0 else 0.0
    entry_band_width_pct = ((entry_high - entry_low) / entry_low * 100) if entry_low > 0 and entry_high >= entry_low else 0.0
    stop_distance_pct = ((entry_high - stop) / entry_high * 100) if entry_high > 0 and stop > 0 and entry_high > stop else 0.0
    reward_to_target_pct = ((target - entry_high) / entry_high * 100) if entry_high > 0 and target > entry_high else 0.0
    trade_plan_valid = bool(entry_low > 0 and entry_high > 0 and stop > 0 and target > entry_high and entry_high > stop)

    risk = entry_high - stop
    reward = target - entry_high
    rr = round(reward / risk, 2) if trade_plan_valid and risk > 0 and reward > 0 else None
    if rr is None:
        rr_valid = False
        rr_level = "無效"
    elif rr >= MIN_BUY_RR:
        rr_valid = True
        rr_level = "有效"
    elif rr >= 1.0:
        rr_valid = False
        rr_level = "觀察"
    else:
        rr_valid = False
        rr_level = "不足"

    return {
        "entry_low": round_price(entry_low) if entry_low else 0.0,
        "entry_high": round_price(entry_high) if entry_high else 0.0,
        "entry_mid": round_price(entry_mid) if entry_mid else 0.0,
        "entry_band_width_pct": round(entry_band_width_pct, 2),
        "stop_loss": round_price(stop) if stop else 0.0,
        "stop_distance_pct": round(stop_distance_pct, 2),
        "target_price": round_price(target) if target else 0.0,
        "reward_to_target_pct": round(reward_to_target_pct, 2),
        "trade_plan_valid": trade_plan_valid,
        "rr": rr,
        "rr_valid": rr_valid,
        "rr_level": rr_level,
        "trade_target": round_price(target) if target else 0.0,
        "fibo_target": round_price(fibo_target) if fibo_target else 0.0,
        "resistance_target": round_price(resistance) if resistance else 0.0,
        "target_source": target_source,
    }

def build_trade_scripts(data: dict) -> dict:
    support = data["support"]
    resistance = data["resistance"]
    next_target = data["fibo"]["next_target"]
    bucket = data.get("state_bucket", "range")

    if bucket == "strong":
        return {
            "script_a": f"劇本A（強勢突破）: 若站穩 {resistance} 之上且量能續強，可順勢追蹤，下一目標看 {next_target}",
            "script_b": f"劇本B（拉回承接）: 若回測 {support} 附近不破，可分批承接；失守則降級為偏多/整理",
            "script_c": f"劇本C（壓力震盪）: 若接近 {resistance} 但量能不足，先等縮量整理後再攻，不宜盲目追高",
        }
    if bucket == "bullish":
        return {
            "script_a": f"劇本A（偏多延續）: 守住 {support} 可維持偏多觀察，等待再次挑戰 {resistance}",
            "script_b": f"劇本B（回測確認）: 若回測 {support} 但止穩，可偏向低接；跌破則先退場觀望",
            "script_c": f"劇本C（轉強升級）: 若有效突破 {resistance} 並量價配合，可由偏多觀察升級為強勢追蹤",
        }
    if bucket == "weak":
        return {
            "script_a": f"劇本A（弱勢反彈）: 若反彈至 {resistance} 下方仍無法突破，先視為弱勢反彈，不宜追價",
            "script_b": f"劇本B（跌破續弱）: 若失守 {support}，優先控管部位，避免逆勢攤平",
            "script_c": f"劇本C（止穩觀察）: 只有重新站回 {support} 並伴隨量價轉強，才考慮恢復偏多",
        }
    return {
        "script_a": f"劇本A（區間低接）: 靠近 {support} 可觀察承接力道，未見止穩前不急著進場",
        "script_b": f"劇本B（跌破下緣）: 若跌破 {support}，區間整理失效，先轉為保守觀察",
        "script_c": f"劇本C（突破上緣）: 若有效突破 {resistance} 並量能配合，可由整理升級為偏多追蹤",
    }


def calc_intraday_score(close, prev_close, open_price, high_price, low_price, support, resistance, orderbook_bias, change_pct):
    score = 50
    comments = []

    if change_pct >= 3:
        score += 20; comments.append("當日漲幅偏強")
    elif change_pct >= 1:
        score += 10; comments.append("當日漲幅為正")
    elif change_pct <= -9:
        score -= 35; comments.append("急跌風險")
    elif change_pct <= -5:
        score -= 20; comments.append("當日跌幅偏大")
    elif change_pct < 0:
        score -= 8; comments.append("當日走弱")

    if close >= open_price:
        score += 8; comments.append("站上開盤")
    else:
        score -= 8; comments.append("跌破開盤")

    if close >= prev_close:
        score += 8; comments.append("站上昨收")
    else:
        score -= 8; comments.append("跌破昨收")

    day_range = max(high_price - low_price, 0.01)
    pos = (close - low_price) / day_range
    if pos >= 0.8:
        score += 12; comments.append("接近日高")
    elif pos <= 0.2:
        score -= 12; comments.append("接近日低")

    if close > resistance:
        score += 18; comments.append("突破壓力")
    elif close >= resistance * 0.995:
        score += 6; comments.append("逼近壓力")
    elif close < support:
        score -= 18; comments.append("跌破支撐")

    if change_pct >= 1.5 and close >= open_price and close >= prev_close:
        score += 10; comments.append("盤中續強")

    if orderbook_bias == "買盤明顯偏強":
        score += 12; comments.append("五檔買盤明顯偏強")
    elif orderbook_bias == "買盤偏強":
        score += 7; comments.append("五檔買盤偏強")
    elif orderbook_bias == "賣盤偏強":
        score -= 8; comments.append("五檔賣盤偏強")

    score = max(0, min(100, int(score)))
    return score, "；".join(comments)


def analyze_symbol(symbol: str) -> dict:
    yf_symbol, df = download_symbol_data(symbol)
    market = detect_market(symbol, yf_symbol)
    stock_name = get_stock_name(symbol, yf_symbol)
    df = calc_indicators(df)
    last = df.iloc[-1]

    fallback_close = round_price(last["Close"])
    fallback_prev_close = round_price(df.iloc[-2]["Close"]) if len(df) >= 2 else fallback_close
    fallback_open = round_price(last["Open"])
    fallback_high = round_price(last["High"])
    fallback_low = round_price(last["Low"])

    if market in ("台股上市", "台股上櫃"):
        rt = get_tw_realtime_quote(symbol, market)
        if rt is None:
            rt = {
                "close": fallback_close, "display_price": fallback_close, "display_note": "日線回退",
                "last_trade": None, "indicative_price": None, "prev_close": fallback_prev_close,
                "open": fallback_open, "high": fallback_high, "low": fallback_low,
                "bid_prices": [], "ask_prices": [], "bid_vols": [], "ask_vols": [],
                "buy_qty": 0, "sell_qty": 0, "orderbook_ratio": "-", "orderbook_bias": "無有效五檔",
                "quote_time": "", "source": "日線回退",
            }
    else:
        rt = get_us_yahoo_quote(
            yf_symbol=yf_symbol,
            fallback_close=fallback_close,
            fallback_prev_close=fallback_prev_close,
            fallback_open=fallback_open,
            fallback_high=fallback_high,
            fallback_low=fallback_low,
        )
        rt["display_price"] = rt["close"]
        rt["display_note"] = "即時/近即時成交價"
        rt["last_trade"] = rt["close"]
        rt["indicative_price"] = rt["close"]
        rt["bid_prices"] = []
        rt["ask_prices"] = []
        rt["bid_vols"] = []
        rt["ask_vols"] = []
        rt["buy_qty"] = 0
        rt["sell_qty"] = 0
        rt["orderbook_ratio"] = "-"
        rt["orderbook_bias"] = "不適用"
        rt["quote_time"] = ""

    close = rt["close"]
    prev_close = rt["prev_close"]
    open_price = rt["open"]
    high_price = rt["high"]
    low_price = rt["low"]

    change = round_price(close - prev_close)
    change_pct = round((change / prev_close) * 100, 2) if prev_close != 0 else 0.0

    ma5 = round_price(last["MA5"]) if pd.notna(last["MA5"]) else close
    ma10 = round_price(last["MA10"]) if pd.notna(last["MA10"]) else close
    ma20 = round_price(last["MA20"]) if pd.notna(last["MA20"]) else close
    ma60 = round_price(last["MA60"]) if pd.notna(last["MA60"]) else close
    rsi = round(float(last["RSI"]), 2) if pd.notna(last["RSI"]) else 50.0

    sr = calc_professional_sr(df)
    support = sr["support"]
    resistance = sr["resistance"]

    trend_score = 50
    comments = []

    if close >= ma5:
        trend_score += 4; comments.append("站上5日線")
    else:
        trend_score -= 4; comments.append("跌破5日線")
    if close >= ma10:
        trend_score += 6; comments.append("站上10日線")
    else:
        trend_score -= 5; comments.append("跌破10日線")
    if close >= ma20:
        trend_score += 10; comments.append("站上20日線")
    else:
        trend_score -= 10; comments.append("跌破20日線")
    if close >= ma60:
        trend_score += 15; comments.append("站上60日線")
    else:
        trend_score -= 12; comments.append("跌破60日線")
    if float(last["MACD"]) >= float(last["MACD_SIGNAL"]):
        trend_score += 8; comments.append("MACD偏多")
    else:
        trend_score -= 6; comments.append("MACD偏弱")
    if pd.notna(last["K"]) and pd.notna(last["D"]):
        if float(last["K"]) >= float(last["D"]):
            trend_score += 6; comments.append("KD偏多")
        else:
            trend_score -= 4; comments.append("KD偏空")
    if rsi < 30:
        trend_score += 8; comments.append("RSI超跌")
    elif rsi > 70:
        trend_score -= 8; comments.append("RSI過熱")
    if len(df) >= 20:
        vol5 = df["Volume"].tail(5).mean()
        vol20 = df["Volume"].tail(20).mean()
        if pd.notna(vol5) and pd.notna(vol20) and vol5 > vol20:
            trend_score += 4; comments.append("量能放大")

    trend_score = max(0, min(100, int(trend_score)))
    intraday_score, intraday_comment = calc_intraday_score(
        close, prev_close, open_price, high_price, low_price, support, resistance,
        rt.get("orderbook_bias", "無"), change_pct
    )
    score = max(0, min(100, int(round(trend_score * 0.6 + intraday_score * 0.4))))

    extra_comment = (
        f"{'；'.join(comments)}"
        f"；盤中={intraday_comment}"
        f"；20日支撐={sr['support20']}"
        f"；20日壓力={sr['resistance20']}"
        f"；波段低點={sr['swing_low']}"
        f"；波段高點={sr['swing_high']}"
        f"；Pivot={sr['pivot']}"
        f"；來源={rt['source']}"
    )

    fibo = calc_fibonacci_targets(df)
    wave = structured_wave_analysis(df)
    fibo_pos = classify_fibo_position(close, fibo)
    price_valid = bool(rt.get("last_trade") is not None and rt.get("source") != "日線回退")

    signal, advice, state_bucket, rule_id, signal_reason = evaluate_trade_state(
        close, prev_close, open_price, support, resistance, change_pct,
        trend_score, intraday_score, score, rt.get("orderbook_bias", "無"),
        ma20=ma20, ma60=ma60, rsi=rsi,
        wave_stage=wave["wave_stage"],
        fibo_position=fibo_pos["fibo_position"],
        fibo_risk_flag=fibo_pos["fibo_risk_flag"],
        wave_risk_flag=wave["wave_risk_flag"],
        rr_valid=True,
        price_valid=price_valid
    )

    result = {
        "input_symbol": symbol, "name": stock_name, "yf_symbol": yf_symbol, "market": market,
        "close": close, "display_price": rt.get("display_price", close), "display_note": rt.get("display_note", ""),
        "last_trade": rt.get("last_trade"), "indicative_price": rt.get("indicative_price"),
        "prev_close": prev_close, "open": open_price, "high": high_price, "low": low_price,
        "change": change, "change_pct": change_pct, "signal": signal, "advice": advice, "score": score,
        "trend_score": trend_score, "intraday_score": intraday_score,
        "support": support, "resistance": resistance, "rsi": rsi, "ma5": ma5, "ma10": ma10,
        "ma20": ma20, "ma60": ma60, "comment": extra_comment,
        "source": rt["source"], "fibo": fibo, "bid_prices": rt.get("bid_prices", []),
        "ask_prices": rt.get("ask_prices", []), "bid_vols": rt.get("bid_vols", []),
        "ask_vols": rt.get("ask_vols", []), "buy_qty": rt.get("buy_qty", 0),
        "sell_qty": rt.get("sell_qty", 0), "orderbook_ratio": rt.get("orderbook_ratio", "-"),
        "orderbook_bias": rt.get("orderbook_bias", "無"), "quote_time": rt.get("quote_time", ""),
        "state_bucket": state_bucket,
        "strategy_level": get_strategy_level(score),
        "strategy_level_score": get_strategy_level_score(get_strategy_level(score)),
        "target_price": fibo.get("next_target", resistance),
        "price_valid": price_valid,
        "execution_price_valid": price_valid,
        "rule_id": rule_id,
        "signal_reason": signal_reason,
        "wave_stage": wave["wave_stage"],
        "wave_score": wave["wave_score"],
        "wave_reason": wave["wave_reason"],
        "wave_risk_flag": wave["wave_risk_flag"],
        "fibo_position": fibo_pos["fibo_position"],
        "fibo_score": fibo_pos["fibo_score"],
        "fibo_risk_flag": fibo_pos["fibo_risk_flag"],
        "fibo_reason": fibo_pos["fibo_reason"],
        "trend_score_detail": "；".join(comments),
        "intraday_score_detail": intraday_comment,
        "wave_score_detail": wave["wave_reason"],
        "fibo_score_detail": fibo_pos["fibo_reason"],
        "decision_model_version": DECISION_MODEL_VERSION,
    }
    result.update(calc_trade_plan(result))

    signal, advice, state_bucket, rule_id, signal_reason = evaluate_trade_state(
        close, prev_close, open_price, support, resistance, change_pct,
        trend_score, intraday_score, score, rt.get("orderbook_bias", "無"),
        ma20=ma20, ma60=ma60, rsi=rsi,
        wave_stage=result["wave_stage"],
        fibo_position=result["fibo_position"],
        fibo_risk_flag=result["fibo_risk_flag"],
        wave_risk_flag=result["wave_risk_flag"],
        rr_valid=result["rr_valid"],
        price_valid=price_valid
    )
    result.update({
        "signal": signal,
        "advice": advice,
        "state_bucket": state_bucket,
        "rule_id": rule_id,
        "signal_reason": signal_reason,
    })
    result.update(calc_trade_plan(result))
    result.update(classify_entry_zone(result))
    result.update(calc_wave_rr_risk_allocation(result))
    result.update(calc_position_sizing(result, account_capital=DEFAULT_ACCOUNT_CAPITAL, risk_pct=DEFAULT_RISK_PCT))
    result["wave_fibo_signal"] = build_wave_fibo_decision_note(result)
    result.update(build_final_decision(result))
    result = ensure_phase4_fields(result)
    result = sync_display_semantics(result)
    log_decision_trace(result)
    result["risk_note"] = build_risk_note(
        close, support, resistance, rsi, score, change_pct,
        wave_stage=result["wave_stage"],
        fibo_position=result["fibo_position"],
        fibo_risk_flag=result["fibo_risk_flag"],
        wave_risk_flag=result["wave_risk_flag"],
        rr_valid=result["rr_valid"],
        price_valid=price_valid
    )
    result["trade_type"] = classify_trade_type(state_bucket, signal, advice)
    result = sync_display_semantics(result)
    result["leader_candidate"] = classify_leader_stage(result)
    result["leader_stage"] = result["leader_candidate"]
    risk_penalty = 18 if (result.get("fibo_risk_flag") or result.get("wave_risk_flag")) else 0
    result["rank_score"] = (
        result["score"] * 0.34 +
        result["trend_score"] * 0.24 +
        result["intraday_score"] * 0.16 +
        result.get("allocation_score", 0) * 0.18 +
        result.get("wave_score", 0) * 0.7 +
        result.get("fibo_score", 0) * 0.7 +
        result["change_pct"] * 1.2 +
        (20 if result.get("execution_ready") else 0) +
        (12 if result.get("entry_zone_ready") else 0) +
        (10 if result.get("position_size_pct", 0) > 0 else 0) +
        (15 if result["leader_candidate"] == "是" else 0) +
        (6 if result["leader_candidate"] == "觀察" else 0) -
        risk_penalty -
        (25 if result.get("allocation_grade") == "BLOCK" else 0) -
        (12 if result.get("final_decision") == "WAIT" else 0) -
        (28 if result.get("final_decision") == "AVOID" else 0)
    )
    result["light"] = get_light(
        result["signal"], result["score"], result["change_pct"],
        intraday_score=result["intraday_score"],
        fibo_risk_flag=result["fibo_risk_flag"],
        wave_risk_flag=result["wave_risk_flag"],
        final_decision=result.get("final_decision")
    )
    result["display_target_price"] = get_display_target(result.get("target_price"), result["signal"], result["state_bucket"])
    result["display_rr"] = normalize_rr_display(result.get("rr"))
    result["summary_block"] = "\n".join([
        "【速讀摘要】",
        f"現價 / 漲跌幅 / 報價：{result['display_price']} / {result['change_pct']:+.2f}% / {result['display_note']}",
        f"總分 / 波段 / 盤中：{result['score']} / {result['trend_score']} / {result['intraday_score']}",
        f"波浪 / 費波 / 禁追：{result['wave_stage']} / {result['fibo_position']} / {'是' if (result['fibo_risk_flag'] or result['wave_risk_flag']) else '否'}",
        f"波費判定：{result['wave_fibo_signal']}",
        f"支撐 / 壓力 / 五檔：{result['support']} / {result['resistance']} / {result['orderbook_bias']}",
        f"燈號 / 訊號 / 建議 / 主升狀態：{result['light']} / {result['signal']} / {result.get('display_advice', result['advice'])} / {result['leader_candidate']}",
        f"交易類型 / 等級：{result.get('display_trade_type', result['trade_type'])} / {result['strategy_level']}",
        f"目標價 / RR / RR等級：{result['display_target_price']} / {result['display_rr']} / {result['rr_level']}",
        f"Phase4：進場狀態={result.get('entry_zone_status','-')} / 倉位={result.get('position_size_pct',0)}% / 資金等級={result.get('allocation_grade','-')} / 配置分={result.get('allocation_score','-')}",
        f"最終決策 / 可下單 / 下單類型：{result['final_decision']} / {result['execution_ready']} / {result.get('order_type_hint','-')} / {result['decision_reason']}",
        f"策略定位：狀態={result['state_bucket']} / 量價比={result['orderbook_ratio']} / RSI={result['rsi']}",
    ])
    result["ai_analysis"] = build_ai_analysis(result)
    result["wave_analysis"] = build_wave_analysis(df)
    result["fibo_analysis"] = build_fibonacci_analysis(fibo)
    result["path_analysis"] = build_bull_bear_path(result)
    result.update(build_trade_scripts(result))
    return result






def get_market_index_quote(symbol: str) -> dict:
    """使用 yfinance 抓取大盤指數；若失敗則回傳 None。"""
    try:
        ticker = yf.Ticker(symbol)
        hist = ticker.history(period="5d", interval="1d", auto_adjust=False)
        if hist is None or hist.empty:
            return None
        last = hist.iloc[-1]
        if len(hist) >= 2:
            prev_close = float(hist.iloc[-2]["Close"])
        else:
            prev_close = float(last["Close"])
        close = float(last["Close"])
        change = close - prev_close
        pct = (change / prev_close * 100) if prev_close else 0.0
        return {"close": round(close, 2), "change": round(change, 2), "pct": round(pct, 2), "source": "Yahoo Finance"}
    except Exception:
        return None

def infer_volume_status(results: list[dict]) -> str:
    if not results:
        return "未知"
    trend_up = sum(1 for r in results if r.get("trend_score", 0) >= 75)
    weak = sum(1 for r in results if r.get("trend_score", 0) < 40)
    if trend_up >= max(2, len(results) * 0.35):
        return "放量"
    if weak >= max(3, len(results) * 0.45):
        return "量縮"
    return "正常"

def _count_change_sign(v) -> int:
    if v in (None, "", "--", "---"):
        return 0
    s = str(v).strip().replace(",", "")
    if any(x in s for x in ["跌", "▼", "-"]):
        try:
            return -1 if float(s.replace("跌", "").replace("▼", "")) != 0 else 0
        except Exception:
            return -1
    if any(x in s for x in ["漲", "+", "▲"]):
        try:
            return 1 if float(s.replace("漲", "").replace("+", "").replace("▲", "")) != 0 else 0
        except Exception:
            return 1
    try:
        f = float(s)
        return 1 if f > 0 else (-1 if f < 0 else 0)
    except Exception:
        return 0

def fetch_twse_breadth() -> tuple[int, int, str]:
    urls = [
        "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
        "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=json",
    ]
    for url in urls:
        try:
            records = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"}).json()
            if isinstance(records, dict):
                records = records.get("data") or records.get("records") or []
            if not isinstance(records, list) or not records:
                continue
            up = down = 0
            for rec in records:
                if not isinstance(rec, dict):
                    continue
                code = str(rec.get("Code") or rec.get("證券代號") or rec.get("股票代號") or "").strip()
                if not (code.isdigit() and len(code) == 4):
                    continue
                sign_val = None
                for key in ("Change", "漲跌價差", "漲跌(+/-)", "漲跌"):
                    if key in rec:
                        sign_val = rec.get(key)
                        break
                sign = _count_change_sign(sign_val)
                if sign > 0:
                    up += 1
                elif sign < 0:
                    down += 1
            if up + down > 0:
                return up, down, "TWSE 官方"
        except Exception:
            continue
    return 0, 0, ""

def fetch_tpex_breadth() -> tuple[int, int, str]:
    candidate_urls = [
        "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes",
        "https://www.tpex.org.tw/openapi/v1/tpex_daily_market_value",
    ]
    for url in candidate_urls:
        try:
            records = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"}).json()
            if not isinstance(records, list) or not records:
                continue
            up = down = 0
            for rec in records:
                if not isinstance(rec, dict):
                    continue
                code = str(rec.get("SecuritiesCompanyCode") or rec.get("股票代號") or rec.get("證券代號") or rec.get("Code") or "").strip()
                if code and not (code.isdigit() and len(code) == 4):
                    continue
                sign_val = None
                for key in ("漲跌", "Change", "漲跌價差", "UpDown", "漲跌(+/-)"):
                    if key in rec:
                        sign_val = rec.get(key)
                        break
                sign = _count_change_sign(sign_val)
                if sign > 0:
                    up += 1
                elif sign < 0:
                    down += 1
            if up + down > 0:
                return up, down, "TPEX 官方"
        except Exception:
            continue
    return 0, 0, ""

def get_tsmc_market_quote() -> dict:
    rt = get_tw_realtime_quote("2330", "台股上市")
    if rt:
        change = round_price(rt["close"] - rt["prev_close"])
        pct = round((change / rt["prev_close"] * 100), 2) if rt["prev_close"] else 0.0
        return {"close": rt["close"], "change": change, "pct": pct, "source": rt.get("source", "TWSE MIS")}
    yq = get_us_yahoo_quote("2330.TW", 0.0, 0.0, 0.0, 0.0, 0.0)
    change = round_price(yq["close"] - yq["prev_close"])
    pct = round((change / yq["prev_close"] * 100), 2) if yq["prev_close"] else 0.0
    return {"close": yq["close"], "change": change, "pct": pct, "source": yq.get("source", "Yahoo Finance")}

def get_market_data(results: list[dict]) -> dict:
    twse = get_market_index_quote("^TWII") or {"close": 0.0, "change": 0.0, "pct": 0.0, "source": "Yahoo Finance"}
    tsmc = get_tsmc_market_quote()

    listed_up, listed_down, src1 = fetch_twse_breadth()
    otc_up, otc_down, src2 = fetch_tpex_breadth()
    up = listed_up + otc_up
    down = listed_down + otc_down
    breadth_source = " / ".join([s for s in (src1, src2) if s]).strip()

    if up + down == 0:
        up = sum(1 for r in results if r.get("change", 0) > 0)
        down = sum(1 for r in results if r.get("change", 0) < 0)
        breadth_source = "觀察池代理"

    return {
        "twse": twse,
        "tsmc": tsmc,
        "up": up,
        "down": down,
        "volume_status": infer_volume_status(results),
        "breadth_source": breadth_source,
        "source_note": f"加權={twse.get('source','Yahoo')} / 台積電={tsmc.get('source','TWSE MIS')} / 家數={breadth_source}",
    }

def get_market_mode(market: dict) -> str:
    twse_pct = market.get("twse", {}).get("pct", 0.0)
    tsmc_pct = market.get("tsmc", {}).get("pct", 0.0)
    up = market.get("up", 0)
    down = market.get("down", 0)
    if twse_pct >= 0.6 and tsmc_pct >= 0.8 and up > down:
        return "偏多震盪"
    if twse_pct <= -0.6 and tsmc_pct <= -0.5 and down > up:
        return "偏弱震盪"
    if twse_pct >= 0 and tsmc_pct >= 0 and up >= down * 0.9:
        return "震盪偏多"
    if twse_pct < 0 and tsmc_pct < 0 and down > up:
        return "震盪偏弱"
    return "區間震盪"

def get_today_strategy(market: dict, mode: str) -> str:
    twse_pct = market.get("twse", {}).get("pct", 0.0)
    tsmc_pct = market.get("tsmc", {}).get("pct", 0.0)
    breadth_balance = market.get("up", 0) - market.get("down", 0)
    if mode == "偏多震盪":
        if tsmc_pct >= 1.0:
            return "大盤與台積電同步偏強，只做主升與整理偏多，避免追高末升段"
        return "指數偏強但台積電未全面發動，以拉回承接為主，不追爆量長紅"
    if mode == "震盪偏多":
        return "大盤偏多但結構未全面擴散，以整理偏多與低接型主升股為主"
    if mode == "偏弱震盪":
        return "大盤與台積電偏弱，優先防守，不抄底弱勢股，只看支撐是否止穩"
    if mode == "震盪偏弱":
        return "盤面偏弱且家數落後，降低持股水位，反彈先看壓力不追價"
    if twse_pct > 0 or breadth_balance > 0:
        return "市場無明確主流但略有撐盤，只做型態完整個股"
    return "市場無明確優勢，觀望為主，等待大盤與台積電同步轉強"

def build_market_overview(results: list[dict]) -> str:
    if not results:
        return "加權：- ｜ 台積電：- ｜ 上漲/下跌：-/- ｜ 量能：未知\n市場模式：尚無資料 ｜ 今日策略：尚無資料"
    market = get_market_data(results)
    mode = get_market_mode(market)
    strategy = get_today_strategy(market, mode)

    twse = market["twse"]
    tsmc = market["tsmc"]
    twse_arrow = "▲" if twse["change"] >= 0 else "▼"
    tsmc_arrow = "▲" if tsmc["change"] >= 0 else "▼"
    line1 = (
        f"加權：{twse['close']} {twse_arrow}{abs(twse['change'])} ({twse['pct']:+.2f}%) ｜ "
        f"台積電：{tsmc['close']} {tsmc_arrow}{abs(tsmc['change'])} ({tsmc['pct']:+.2f}%) ｜ "
        f"上漲/下跌：{market['up']}/{market['down']} ｜ 量能：{market['volume_status']}"
    )
    line2 = f"市場模式：{mode} ｜ 今日策略：{strategy}"
    return line1 + "\n" + line2



class GTCProApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_TITLE} {APP_VERSION}")
        self.root.geometry("1920x1000")
        self.root.minsize(1500, 820)
        self.results = []
        self.current_sort_column = None
        self.sort_reverse = True
        self.auto_refresh_enabled = False
        self.next_refresh_sec = AUTO_REFRESH_MS // 1000
        self.last_update_time = None
        self._timer_job_id = None
        self.show_advanced_columns = False
        self.market_overview_var = tk.StringVar(value="加權：- ｜ 台積電：- ｜ 上漲/下跌：-/- ｜ 量能：未知\n市場模式：尚無資料 ｜ 今日策略：尚無資料")
        self.data_source_var = tk.StringVar(value="資料來源：尚無資料")
        self._build_ui()
        self.set_status(f"系統已就緒。當前版本：{APP_VERSION}")


    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        input_frame = ttk.Frame(top)
        input_frame.pack(side="left", fill="x", expand=True)

        ttk.Label(input_frame, text="股票代號（逗號分隔）").pack(side="left", padx=(0, 8))
        self.symbol_entry = ttk.Entry(input_frame, width=80)
        self.symbol_entry.pack(side="left", padx=(0, 8), fill="x", expand=True)
        self.symbol_entry.insert(0, "2330,2382,3231,2308,3017,4979,AAPL,NVDA,MSFT")

        action_frame = ttk.Frame(top)
        action_frame.pack(side="left", padx=(8, 0))

        ttk.Button(action_frame, text="執行分析", command=self.run_analysis).pack(side="left", padx=(0, 6))
        ttk.Button(action_frame, text="啟用自動刷新", command=self.enable_auto_refresh).pack(side="left", padx=(0, 6))
        ttk.Button(action_frame, text="停止自動刷新", command=self.disable_auto_refresh).pack(side="left", padx=(0, 6))
        ttk.Button(action_frame, text="切換進階欄位", command=self.toggle_advanced_columns).pack(side="left", padx=(0, 6))
        ttk.Button(action_frame, text="清空", command=self.clear_results).pack(side="left", padx=(0, 6))

        right_frame = ttk.Frame(top)
        right_frame.pack(side="right")

        self.download_btn = tk.Menubutton(right_frame, text="下載報告 ▼", relief="raised")
        self.download_menu = tk.Menu(self.download_btn, tearoff=0)
        self.download_btn.config(menu=self.download_menu)
        self.download_menu.add_command(label="PDF：總表摘要", command=self.export_pdf_summary)
        self.download_menu.add_command(label="PDF：目前選取個股", command=self.export_pdf_selected)
        self.download_menu.add_command(label="PDF：全部完整報告", command=self.export_pdf_full)
        self.download_menu.add_separator()
        self.download_menu.add_command(label="TXT：全部完整報告", command=self.export_txt_full)
        self.download_menu.add_command(label="CSV：主表資料", command=self.export_csv_table)
        self.download_btn.pack(side="right", padx=(8, 0))

        market_bar = ttk.Frame(self.root, padding=(10, 0, 10, 6))
        market_bar.pack(fill="x")
        ttk.Label(market_bar, textvariable=self.market_overview_var, justify="left").pack(anchor="w")

        center = ttk.Panedwindow(self.root, orient=tk.VERTICAL)
        center.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        top_frame = ttk.Frame(center)
        bottom_frame = ttk.Frame(center)
        center.add(top_frame, weight=3)
        center.add(bottom_frame, weight=2)
        self._build_table_area(top_frame)
        self._build_detail_area(bottom_frame)
    def _build_table_area(self, parent):
        columns = (
            "排名", "燈號", "市場", "代號", "名稱", "顯示價", "漲跌", "漲跌幅%",
            "訊號", "建議", "分數", "等級", "目標價", "RR", "主升候選", "波浪定位", "費波位置", "禁追提示", "波費判定", "波段分", "盤中分", "支撐", "壓力", "RSI",
            "五檔力道", "交易類型", "進場狀態", "進場可執行", "建議倉位%", "資金等級", "配置分", "阻擋原因", "下單類型", "最終決策", "可下單", "決策原因", "報價說明"
        )
        self.all_columns = columns
        self.core_columns = ("排名", "燈號", "市場", "代號", "名稱", "顯示價", "漲跌幅%", "訊號", "建議", "分數", "等級", "目標價", "RR", "主升候選", "波浪定位", "費波位置", "禁追提示", "波費判定", "進場狀態", "建議倉位%", "資金等級", "配置分", "最終決策", "可下單")
        self.advanced_columns = ("排名", "燈號", "市場", "代號", "名稱", "顯示價", "漲跌", "漲跌幅%", "訊號", "建議", "分數", "等級", "目標價", "RR", "主升候選", "波浪定位", "費波位置", "禁追提示", "波費判定", "波段分", "盤中分", "支撐", "壓力", "RSI", "五檔力道", "交易類型", "進場狀態", "進場可執行", "建議倉位%", "資金等級", "配置分", "阻擋原因", "下單類型", "最終決策", "可下單", "決策原因", "報價說明")
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", height=16)
        self.tree.configure(displaycolumns=self.core_columns)
        widths = {
            "排名": 55, "燈號": 55, "市場": 90, "代號": 80, "名稱": 190, "顯示價": 90,
            "漲跌": 90, "漲跌幅%": 95, "訊號": 100, "建議": 130, "分數": 65, "等級": 60, "目標價": 90, "RR": 70, "主升候選": 90,
            "波浪定位": 100, "費波位置": 130, "禁追提示": 90, "波費判定": 220,
            "波段分": 70, "盤中分": 70, "支撐": 90, "壓力": 90, "RSI": 70,
            "五檔力道": 110, "交易類型": 100, "進場狀態": 120, "進場可執行": 100, "建議倉位%": 100, "資金等級": 90, "配置分": 80, "阻擋原因": 260, "下單類型": 100, "最終決策": 90, "可下單": 80, "決策原因": 280, "報價說明": 180
        }
        for c in columns:
            self.tree.heading(c, text=c, command=lambda col=c: self.sort_by_column(col))
            self.tree.column(c, width=widths[c], anchor="center")
        self.tree.tag_configure("up", foreground="red", background="#ffecec")
        self.tree.tag_configure("down", foreground="green", background="#ecffec")
        self.tree.tag_configure("flat", foreground="black", background="white")
        self.tree.tag_configure("strong", background="#fff2b3")
        self.tree.tag_configure("watch", background="#eef5ff")
        self.tree.tag_configure("danger", background="#ffd9d9")
        self.tree.tag_configure("level_a", background="#fff2b3")
        self.tree.tag_configure("level_b", background="#eef5ff")
        self.tree.tag_configure("level_c", background="#f3f3f3")
        self.tree.tag_configure("level_d", background="#fce8e8")
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)
        yscroll = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(parent, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)

    def toggle_advanced_columns(self):
        self.show_advanced_columns = not self.show_advanced_columns
        if self.show_advanced_columns:
            self.tree.configure(displaycolumns=self.advanced_columns)
            self.set_status(f"已顯示進階欄位。版本：{APP_VERSION}")
        else:
            self.tree.configure(displaycolumns=self.core_columns)
            self.set_status(f"已切回核心欄位。版本：{APP_VERSION}")

    def _build_detail_area(self, parent):
        left = ttk.LabelFrame(parent, text="個股明細分析", padding=10)
        right = ttk.LabelFrame(parent, text="操作建議 / 風險提醒", padding=10)
        left.pack(side="left", fill="both", expand=True, padx=(0, 5))
        right.pack(side="left", fill="both", expand=True, padx=(5, 0))
        left_text_frame = ttk.Frame(left)
        left_text_frame.pack(fill="both", expand=True)
        self.detail_text = tk.Text(left_text_frame, height=14, wrap="none", font=("Microsoft JhengHei", 10))
        left_y_scroll = ttk.Scrollbar(left_text_frame, orient="vertical", command=self.detail_text.yview)
        left_x_scroll = ttk.Scrollbar(left_text_frame, orient="horizontal", command=self.detail_text.xview)
        self.detail_text.configure(yscrollcommand=left_y_scroll.set, xscrollcommand=left_x_scroll.set)
        self.detail_text.grid(row=0, column=0, sticky="nsew")
        left_y_scroll.grid(row=0, column=1, sticky="ns")
        left_x_scroll.grid(row=1, column=0, sticky="ew")
        left_text_frame.rowconfigure(0, weight=1)
        left_text_frame.columnconfigure(0, weight=1)
        right_text_frame = ttk.Frame(right)
        right_text_frame.pack(fill="both", expand=True)
        self.advice_text = tk.Text(right_text_frame, height=14, wrap="none", font=("Microsoft JhengHei", 10))
        right_y_scroll = ttk.Scrollbar(right_text_frame, orient="vertical", command=self.advice_text.yview)
        right_x_scroll = ttk.Scrollbar(right_text_frame, orient="horizontal", command=self.advice_text.xview)
        self.advice_text.configure(yscrollcommand=right_y_scroll.set, xscrollcommand=right_x_scroll.set)
        self.advice_text.grid(row=0, column=0, sticky="nsew")
        right_y_scroll.grid(row=0, column=1, sticky="ns")
        right_x_scroll.grid(row=1, column=0, sticky="ew")
        right_text_frame.rowconfigure(0, weight=1)
        right_text_frame.columnconfigure(0, weight=1)
        bottom = ttk.LabelFrame(self.root, text="系統訊息", padding=10)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        self.status_var = tk.StringVar(value="")
        bottom_row = ttk.Frame(bottom)
        bottom_row.pack(fill="x")
        ttk.Label(bottom_row, textvariable=self.status_var).pack(side="left", anchor="w")
        ttk.Label(bottom_row, textvariable=self.data_source_var, foreground="gray").pack(side="right", anchor="e")

    def set_status(self, text: str):
        now = datetime.now().strftime("%H:%M:%S")
        self.status_var.set(f"[{now}] {text}")

    def get_light(self, signal, score, change_pct, intraday_score=None):
        return get_light(signal, score, change_pct, intraday_score)

    def update_status_with_timer(self):
        if self.last_update_time:
            last = self.last_update_time.strftime("%H:%M:%S")
        else:
            last = "-"
        mode = "自動刷新開啟" if self.auto_refresh_enabled else "自動刷新關閉"
        self.status_var.set(
            f"最後更新：{last} ｜ 下次刷新：{self.next_refresh_sec} 秒 ｜ {mode} ｜ "
            f"追蹤檔數：{len(self.results)} ｜ 版本：{APP_VERSION}"
        )

    def update_data_source_bar(self):
        try:
            if self.results:
                market = get_market_data(self.results)
                now_text = datetime.now().strftime("%H:%M:%S")
                self.data_source_var.set(
                    f"資料來源：{market['source_note']} ｜ 更新時間：{now_text}"
                )
            else:
                self.data_source_var.set("資料來源：尚無資料")
        except Exception as e:
            self.data_source_var.set(f"資料來源：更新失敗 ({e})")

    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = []
        self.data_source_var.set("資料來源：尚無資料")
        self.detail_text.delete("1.0", tk.END)
        self.advice_text.delete("1.0", tk.END)
        self.set_status(f"已清空結果。版本：{APP_VERSION}")

    def parse_symbols(self):
        raw = self.symbol_entry.get().strip()
        if not raw:
            return []
        parts = [x.strip() for x in raw.replace("，", ",").split(",")]
        return [p for p in parts if p]

    def render_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, r in enumerate(self.results, start=1):
            tags = []
            if r["change"] > 0:
                tags.append("up")
            elif r["change"] < 0:
                tags.append("down")
            else:
                tags.append("flat")
            if r["signal"] == "急跌風險":
                tags.append("danger")
            elif r["score"] >= 80:
                tags.append("strong")
            elif r["score"] >= 65:
                tags.append("watch")

            level_tag_map = {
                "A": "level_a",
                "B": "level_b",
                "C": "level_c",
                "D": "level_d",
            }
            if r.get("strategy_level") in level_tag_map:
                tags.append(level_tag_map[r["strategy_level"]])

            light = r.get("light") or self.get_light(r["signal"], r["score"], r["change_pct"], r["intraday_score"])
            self.tree.insert(
                "", "end",
                values=(
                    idx, light, r["market"], r["input_symbol"], r["name"], r["display_price"],
                    f"{r['change']:+.2f}", f"{r['change_pct']:+.2f}%", r["signal"], r.get("display_advice", r["advice"]),
                    r["score"], r["strategy_level"], r["display_target_price"], r["display_rr"], r["leader_candidate"], r.get("wave_stage", "-"), r.get("fibo_position", "-"),
                    "是" if (r.get("fibo_risk_flag") or r.get("wave_risk_flag")) else "否", r.get("wave_fibo_signal", "-"),
                    r["trend_score"], r["intraday_score"], r["support"], r["resistance"],
                    r["rsi"], r["orderbook_bias"], r.get("display_trade_type", r["trade_type"]), r.get("entry_zone_status", "-"), "是" if r.get("entry_zone_ready") else "否", r.get("position_size_pct", 0), r.get("allocation_grade", "-"), r.get("allocation_score", 0), r.get("final_block_reason", r.get("phase4_block_reason", "-")), r.get("order_type_hint", "-"), r.get("final_decision", "-"), "是" if r.get("execution_ready") else "否", r.get("decision_reason", "-"), r["display_note"]
                ),
                tags=tuple(tags)
            )

    def run_analysis(self):
        symbols = self.parse_symbols()
        if not symbols:
            messagebox.showwarning("提醒", "請輸入至少一個股票代號。")
            return
        self.clear_results()
        self.set_status(f"開始抓取即時股票資料... 版本：{APP_VERSION}")
        self.root.update_idletasks()
        ok_results, errors = [], []
        for sym in symbols:
            try:
                result = analyze_symbol(sym)
                ok_results.append(result)
                self.set_status(f"完成：{sym} / 版本：{APP_VERSION}")
                self.root.update_idletasks()
            except Exception as e:
                errors.append(f"{sym}: {e}")
        self.results = sorted(ok_results, key=lambda x: x["rank_score"], reverse=True)
        self.render_results()
        self.market_overview_var.set(build_market_overview(self.results))
        self.update_data_source_bar()
        self.last_update_time = datetime.now()
        self.next_refresh_sec = AUTO_REFRESH_MS // 1000
        self.update_status_with_timer()
        if self.results:
            first_id = self.tree.get_children()[0]
            self.tree.selection_set(first_id)
            self.tree.focus(first_id)
            self.on_row_select()
        if errors:
            self.set_status(f"完成 {len(self.results)} 檔，失敗 {len(errors)} 檔。版本：{APP_VERSION}")
            messagebox.showwarning("部分股票失敗", "\n".join(errors[:10]))
        else:
            self.set_status(f"分析完成，共 {len(self.results)} 檔。版本：{APP_VERSION}")

    def enable_auto_refresh(self):
        self.auto_refresh_enabled = True
        self.next_refresh_sec = AUTO_REFRESH_MS // 1000
        self.update_status_with_timer()
        if self._timer_job_id is not None:
            try:
                self.root.after_cancel(self._timer_job_id)
            except Exception:
                pass
        self._timer_job_id = self.root.after(1000, self.auto_refresh_job)

    def disable_auto_refresh(self):
        self.auto_refresh_enabled = False
        if self._timer_job_id is not None:
            try:
                self.root.after_cancel(self._timer_job_id)
            except Exception:
                pass
            self._timer_job_id = None
        self.update_status_with_timer()

    def auto_refresh_job(self):
        if not self.auto_refresh_enabled:
            self._timer_job_id = None
            return
        self.next_refresh_sec -= 1
        if self.next_refresh_sec <= 0:
            symbols = self.parse_symbols()
            if symbols:
                try:
                    self.run_analysis()
                except Exception:
                    pass
            self.next_refresh_sec = AUTO_REFRESH_MS // 1000
        self.update_status_with_timer()
        self._timer_job_id = self.root.after(1000, self.auto_refresh_job)


    def _get_result_by_symbol(self, symbol: str):
        return next((r for r in self.results if r["input_symbol"] == symbol), None)

    def _build_detail_lines(self, target: dict):
        return [
            f"【{target['input_symbol']} {target['name']}】個股明細分析",
            f"市場：{target['market']}",
            f"資料來源：{target['source']}",
            f"報價時間：{target['quote_time']}",
            f"顯示價：{target['display_price']}",
            f"報價說明：{target['display_note']}",
            f"即時成交價：{target['last_trade'] if target['last_trade'] is not None else '-'}",
            f"參考價/中間價：{target['indicative_price'] if target['indicative_price'] is not None else '-'}",
            f"昨收：{target['prev_close']}",
            f"開盤：{target['open']}",
            f"最高：{target['high']}",
            f"最低：{target['low']}",
            f"漲跌：{target['change']:+.2f}",
            f"漲跌幅：{target['change_pct']:+.2f}%",
            "",
            target["summary_block"],
            f"交易計畫：進場={target.get('entry_low',0)}~{target.get('entry_high',0)} / 停損={target.get('stop_loss',0)} / 目標={target.get('display_target_price','-')} / RR={target.get('display_rr','-')}",
            "",
            "【五檔資訊】",
            f"買盤總量：{target['buy_qty']}",
            f"賣盤總量：{target['sell_qty']}",
            f"委買/委賣比：{target['orderbook_ratio']}",
            f"五檔力道：{target['orderbook_bias']}",
            f"買一：{target['bid_prices'][0] if target['bid_prices'] else '-'} / 量：{target['bid_vols'][0] if target['bid_vols'] else '-'}",
            f"賣一：{target['ask_prices'][0] if target['ask_prices'] else '-'} / 量：{target['ask_vols'][0] if target['ask_vols'] else '-'}",
            "",
            "【均線結構】",
            f"MA5：{target['ma5']}",
            f"MA10：{target['ma10']}",
            f"MA20：{target['ma20']}",
            f"MA60：{target['ma60']}",
            "",
            "【技術指標】",
            f"RSI：{target['rsi']}",
            f"綜合訊號：{target['signal']}",
            f"綜合分數：{target['score']}"
            ,f"策略等級：{target['strategy_level']}",
            f"波段分數：{target['trend_score']}",
            f"盤中分數：{target['intraday_score']}",
            f"Rule ID：{target.get('rule_id','-')}",
            f"訊號原因：{target.get('signal_reason','-')}",
            "",
            "【波浪費波決策判定】",
            f"波浪定位：{target.get('wave_stage','-')}",
            f"波浪原因：{target.get('wave_reason','-')}",
            f"費波位置：{target.get('fibo_position','-')}",
            f"費波原因：{target.get('fibo_reason','-')}",
            f"禁追提示：{'是' if (target.get('fibo_risk_flag') or target.get('wave_risk_flag')) else '否'}",
            f"RR有效：{target.get('rr_valid','-')} / RR等級：{target.get('rr_level','-')}",
            f"波費判定：{target.get('wave_fibo_signal','-')}",
            f"最終決策：{target.get('final_decision','-')} / 可下單：{target.get('execution_ready','-')}",
            f"決策原因：{target.get('decision_reason','-')}",
            "",
            "【Phase4交易引擎】",
            f"進場狀態：{target.get('entry_zone_status','-')} / 可執行：{'是' if target.get('entry_zone_ready') else '否'} / 距離進場%：{target.get('distance_to_entry_pct','-')}",
            f"追價風險：{'是' if target.get('chase_risk_flag') else '否'} / 下單類型：{target.get('order_type_hint','-')}",
            f"配置分：{target.get('allocation_score','-')} / 資金等級：{target.get('allocation_grade','-')} / 配置倍率：{target.get('allocation_multiplier','-')}",
            f"建議倉位%：{target.get('position_size_pct','-')} / 建議股數：{target.get('suggested_shares','-')} / 風險預算%：{target.get('risk_budget_pct','-')} / 最大損失%：{target.get('max_loss_pct','-')}",
            f"Phase4阻擋原因：{target.get('final_block_reason') or target.get('phase4_block_reason','') or '-'}",
            "",
            "【支撐壓力】",
            f"主支撐：{target['support']}",
            f"主壓力：{target['resistance']}",
            "",
            "【技術說明】",
            target["comment"],
            "",
            target["ai_analysis"],
            "",
            target["wave_analysis"],
            "",
            target["fibo_analysis"],
            "",
            target["path_analysis"],
        ]

    def _build_advice_lines(self, target: dict):
        rr_text = f"1:{target['rr']:.2f}" if target.get('rr') is not None else "-"
        entry_text = (
            f"{target['entry_low']} ~ {target['entry_high']}"
            if target.get('entry_high', 0) > 0 else "弱勢不建議主動進場"
        )
        return [
            f"【{target['input_symbol']} {target['name']}】交易決策報告",
            "【交易結論】",
            f"建議：{target.get('display_advice', target['advice'])}",
            f"訊號：{target['signal']} / 狀態：{target['state_bucket']} / 主升：{target['leader_candidate']}",
            "",
            "【交易計畫】",
            f"建議進場：{entry_text}",
            f"停損點：{target.get('stop_loss', 0)}",
            f"策略等級：{target['strategy_level']}",
            f"第一目標：{target.get('display_target_price', '-') }",
            f"風險報酬比：{rr_text}",
            f"RR等級：{target.get('rr_level','-')} / RR有效：{target.get('rr_valid','-')}",
            "",
            "【波浪費波決策】",
            f"波浪定位：{target.get('wave_stage','-')} / 費波位置：{target.get('fibo_position','-')}",
            f"禁追提示：{'是' if (target.get('fibo_risk_flag') or target.get('wave_risk_flag')) else '否'}",
            f"波費判定：{target.get('wave_fibo_signal','-')}",
            f"最終決策：{target.get('final_decision','-')} / 可下單：{target.get('execution_ready','-')}",
            f"決策原因：{target.get('decision_reason','-')}",
            "",
            "【Phase4交易引擎】",
            f"進場狀態：{target.get('entry_zone_status','-')} / 可執行：{'是' if target.get('entry_zone_ready') else '否'} / 進場距離%：{target.get('distance_to_entry_pct','-')}",
            f"建議倉位：{target.get('position_size_pct','-')}% / 建議股數：{target.get('suggested_shares','-')} / 風險預算：{target.get('risk_budget_pct','-')}%",
            f"資金等級：{target.get('allocation_grade','-')} / 配置分：{target.get('allocation_score','-')} / 下單類型：{target.get('order_type_hint','-')}",
            f"阻擋原因：{target.get('final_block_reason') or target.get('phase4_block_reason','') or '-'}",
            "",
            "【風險提醒】",
            target["risk_note"],
            "",
            "【關鍵價位】",
            f"主支撐：{target['support']}",
            f"主壓力：{target['resistance']}",
            f"下一目標價：{target['fibo']['next_target']}",
            f"1.0：{target['fibo']['target_1_0']}",
            f"1.382：{target['fibo']['target_1_382']}",
            f"1.618：{target['fibo']['target_1_618']}",
            "",
            "【多空路徑重點】",
            f"多方關鍵：守 {target['support']}、破 {target['resistance']}、看 {target['fibo']['next_target']}",
            f"空方關鍵：失守 {target['support']} 後，短線結構轉弱",
            f"狀態分類：{target['state_bucket']}",
            "",
            "【交易劇本】",
            target["script_a"],
            target["script_b"],
            target["script_c"],
            "",
            "【操作觀察重點】",
            f"1. 報價模式：{target['display_note']}",
            f"2. 支撐區：{target['support']} 附近是否守穩",
            f"3. 壓力區：{target['resistance']} 附近是否放量突破",
            f"4. RSI：{target['rsi']} 是否進一步轉強/轉弱",
            f"5. 五檔力道：{target['orderbook_bias']} / 比值={target['orderbook_ratio']}",
            f"6. 波段分 / 盤中分 / 總分：{target['trend_score']} / {target['intraday_score']} / {target['score']}",
            f"7. 主升狀態：{target['leader_candidate']} / 狀態：{target['state_bucket']}",
            f"8. 均線結構：MA20={target['ma20']} / MA60={target['ma60']}",
            f"9. 波浪定位：{target.get('wave_stage','-')} / 費波位置：{target.get('fibo_position','-')}",
            f"10. Phase4：進場={target.get('entry_zone_status','-')} / 倉位={target.get('position_size_pct','-')}% / 資金等級={target.get('allocation_grade','-')}",
            f"11. 最終決策：{target.get('final_decision','-')} / Rule={target.get('rule_id','-')}",
        ]

    def _render_selected_result(self, target: dict):
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, "\n".join(self._build_detail_lines(target)))
        self.advice_text.delete("1.0", tk.END)
        self.advice_text.insert(tk.END, "\n".join(self._build_advice_lines(target)))

    def on_row_select(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        values = item["values"]
        if not values:
            return
        symbol = str(values[3])
        target = self._get_result_by_symbol(symbol)
        if not target:
            return
        self._render_selected_result(target)

    def _draw_wrapped_lines(self, canvas_obj, font_name, lines, x, y, max_width, line_height=12):
        canvas_obj.setFont(font_name, 9)
        for raw_line in lines:
            text = "" if raw_line is None else str(raw_line)
            if text == "":
                y -= line_height
                if y < 42:
                    canvas_obj.showPage()
                    canvas_obj.setFont(font_name, 9)
                    y = 560
                continue

            current = ""
            for ch in text:
                candidate = current + ch
                if canvas_obj.stringWidth(candidate, font_name, 9) <= max_width:
                    current = candidate
                else:
                    canvas_obj.drawString(x, y, current)
                    y -= line_height
                    if y < 42:
                        canvas_obj.showPage()
                        canvas_obj.setFont(font_name, 9)
                        y = 560
                    current = ch
            if current:
                canvas_obj.drawString(x, y, current)
                y -= line_height
                if y < 42:
                    canvas_obj.showPage()
                    canvas_obj.setFont(font_name, 9)
                    y = 560
        return y

    def _export_pdf_header(self, canvas_obj, font_name, title):
        width, height = landscape(A4)
        canvas_obj.setFont(font_name, 16)
        canvas_obj.drawString(24, height - 28, f"{APP_TITLE} {APP_VERSION}")
        canvas_obj.setFont(font_name, 10)
        canvas_obj.drawString(24, height - 46, title)
        canvas_obj.setFont(font_name, 9)
        canvas_obj.drawString(24, height - 62, f"報告時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        return width, height
    def sort_by_column(self, col_name):
        if not self.results:
            return
        key_map = {
            "排名": "rank_score",
            "燈號": "light",
            "市場": "market",
            "代號": "input_symbol",
            "名稱": "name",
            "顯示價": "display_price",
            "漲跌": "change",
            "漲跌幅%": "change_pct",
            "訊號": "signal",
            "建議": "display_advice",
            "分數": "score",
            "等級": "strategy_level_score",
            "目標價": "target_price",
            "RR": "rr",
            "主升候選": "leader_candidate",
            "波浪定位": "wave_stage",
            "費波位置": "fibo_position",
            "禁追提示": "fibo_risk_flag",
            "波費判定": "wave_fibo_signal",
            "波段分": "trend_score",
            "盤中分": "intraday_score",
            "支撐": "support",
            "壓力": "resistance",
            "RSI": "rsi",
            "五檔力道": "orderbook_bias",
            "交易類型": "display_trade_type",
            "進場狀態": "entry_zone_status",
            "進場可執行": "entry_zone_ready",
            "建議倉位%": "position_size_pct",
            "資金等級": "allocation_grade",
            "配置分": "allocation_score",
            "阻擋原因": "final_block_reason",
            "下單類型": "order_type_hint",
            "最終決策": "final_decision",
            "可下單": "execution_ready",
            "決策原因": "decision_reason",
            "報價說明": "display_note",
        }
        real_key = key_map.get(col_name)
        if real_key is None:
            self.results = sorted(self.results, key=lambda x: x["rank_score"], reverse=True)
            self.render_results()
            return
        if self.current_sort_column == col_name:
            self.sort_reverse = not self.sort_reverse
        else:
            self.current_sort_column = col_name
            self.sort_reverse = True
        def sort_value(x):
            v = x.get(real_key)
            if real_key == "strategy_level_score":
                return int(v or 0)
            if real_key in ("target_price", "rr", "display_target_price", "display_rr"):
                if v in (None, "-"):
                    return float("-inf") if self.sort_reverse else float("inf")
                return float(v)
            return v
        self.results = sorted(self.results, key=sort_value, reverse=self.sort_reverse)
        self.render_results()


    def export_pdf_summary(self):
        if not self.results:
            messagebox.showwarning("提醒", "目前沒有分析結果可匯出。")
            return
        file_path = filedialog.asksaveasfilename(title="PDF：總表摘要", defaultextension=".pdf", filetypes=[("PDF file", "*.pdf"), ("All files", "*.*")])
        if not file_path:
            return
        font_name = setup_pdf_font()
        c = canvas.Canvas(file_path, pagesize=landscape(A4))
        width, height = self._export_pdf_header(c, font_name, "總表摘要")
        headers = ["排名", "燈", "代號", "名稱", "現價", "漲跌%", "訊號", "建議", "分數", "RR", "波浪", "費波", "進場", "倉位%", "資金", "決策", "可下單"]
        x_positions = [18, 45, 72, 115, 220, 268, 318, 375, 432, 468, 510, 558, 615, 672, 720, 765, 810]
        y = height - 82
        c.setFont(font_name, 8)
        for h, x in zip(headers, x_positions):
            c.drawString(x, y, h)
        y -= 14
        c.line(18, y + 8, width - 18, y + 8)

        for idx, r in enumerate(self.results, start=1):
            if y < 42:
                c.showPage()
                width, height = self._export_pdf_header(c, font_name, "總表摘要（續）")
                y = height - 50
                c.setFont(font_name, 8)
            row = [
                str(idx), r["light"], r["input_symbol"], r["name"][:8], str(r["display_price"]),
                f"{r['change_pct']:+.2f}%", r["signal"][:6], r.get("display_advice", r["advice"])[:7], str(r["score"]),
                str(r["display_rr"]), str(r.get("wave_stage", "-"))[:5], str(r.get("fibo_position", "-"))[:7],
                str(r.get("entry_zone_status", "-"))[:8], str(r.get("position_size_pct", 0)), str(r.get("allocation_grade", "-")),
                str(r.get("final_decision", "-")), str(r.get("execution_ready", "-"))
            ]
            for text, x in zip(row, x_positions):
                c.drawString(x, y, str(text))
            y -= 14
        c.save()
        self.set_status(f"已匯出 PDF 總表：{file_path} / 版本：{APP_VERSION}")
        messagebox.showinfo("完成", "PDF：總表摘要匯出成功。")

    def export_pdf_selected(self):
        if not self.results:
            messagebox.showwarning("提醒", "目前沒有分析結果可匯出。")
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提醒", "請先在表格中選取一檔個股。")
            return
        item = self.tree.item(selected[0])
        symbol = str(item["values"][3])
        target = self._get_result_by_symbol(symbol)
        if not target:
            messagebox.showwarning("提醒", "找不到目前選取個股資料。")
            return
        file_path = filedialog.asksaveasfilename(title="PDF：目前選取個股", defaultextension=".pdf", filetypes=[("PDF file", "*.pdf"), ("All files", "*.*")])
        if not file_path:
            return
        font_name = setup_pdf_font()
        c = canvas.Canvas(file_path, pagesize=landscape(A4))
        width, height = self._export_pdf_header(c, font_name, f"個股完整報告：{target['input_symbol']} {target['name']}")
        y = height - 82
        self._draw_wrapped_lines(c, font_name, self._build_detail_lines(target), 24, y, 520)
        c.showPage()
        width, height = self._export_pdf_header(c, font_name, f"操作建議與劇本：{target['input_symbol']} {target['name']}")
        y = height - 82
        self._draw_wrapped_lines(c, font_name, self._build_advice_lines(target), 24, y, 760)
        c.save()
        self.set_status(f"已匯出 PDF 個股：{file_path} / 版本：{APP_VERSION}")
        messagebox.showinfo("完成", "PDF：目前選取個股匯出成功。")

    def export_pdf_full(self):
        if not self.results:
            messagebox.showwarning("提醒", "目前沒有分析結果可匯出。")
            return
        file_path = filedialog.asksaveasfilename(title="PDF：全部完整報告", defaultextension=".pdf", filetypes=[("PDF file", "*.pdf"), ("All files", "*.*")])
        if not file_path:
            return
        font_name = setup_pdf_font()
        c = canvas.Canvas(file_path, pagesize=landscape(A4))
        width, height = self._export_pdf_header(c, font_name, f"盤勢總覽：{self.market_overview_var.get()}")
        y = height - 82
        overview_lines = [
            "【總表摘要】",
            f"追蹤檔數：{len(self.results)}",
            self.market_overview_var.get(),
            ""
        ]
        for idx, r in enumerate(self.results, start=1):
            overview_lines.append(
                f"{idx}. {r['input_symbol']} {r['name']} / 現價={r['display_price']} / 漲跌幅={r['change_pct']:+.2f}% / 訊號={r['signal']} / 建議={r.get('display_advice', r['advice'])} / 分數={r['score']} / 等級={r['strategy_level']} / 目標價={r['display_target_price']} / RR={r['display_rr']} / 主升={r['leader_candidate']} / 波浪={r.get('wave_stage','-')} / 費波={r.get('fibo_position','-')} / 決策={r.get('final_decision','-')} / 可下單={r.get('execution_ready','-')} / 阻擋={r.get('final_block_reason','-')}"
            )
        self._draw_wrapped_lines(c, font_name, overview_lines, 24, y, 760)

        for r in self.results:
            c.showPage()
            width, height = self._export_pdf_header(c, font_name, f"個股完整報告：{r['input_symbol']} {r['name']}")
            y = height - 82
            self._draw_wrapped_lines(c, font_name, self._build_detail_lines(r), 24, y, 520)
            c.showPage()
            width, height = self._export_pdf_header(c, font_name, f"操作建議與劇本：{r['input_symbol']} {r['name']}")
            y = height - 82
            self._draw_wrapped_lines(c, font_name, self._build_advice_lines(r), 24, y, 760)
        c.save()
        self.set_status(f"已匯出 PDF 完整報告：{file_path} / 版本：{APP_VERSION}")
        messagebox.showinfo("完成", "PDF：全部完整報告匯出成功。")

    def export_txt_full(self):
        if not self.results:
            messagebox.showwarning("提醒", "目前沒有分析結果可匯出。")
            return
        file_path = filedialog.asksaveasfilename(title="TXT：全部完整報告", defaultextension=".txt", filetypes=[("Text file", "*.txt"), ("All files", "*.*")])
        if not file_path:
            return
        lines = [
            f"{APP_TITLE} {APP_VERSION}",
            f"報告時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            self.market_overview_var.get(),
            "=" * 140
        ]
        for idx, r in enumerate(self.results, start=1):
            lines.append(f"[{idx}] {r['input_symbol']} {r['name']}")
            lines.extend(self._build_detail_lines(r))
            lines.append("")
            lines.extend(self._build_advice_lines(r))
            lines.append("-" * 140)
        with open(file_path, "w", encoding="utf-8-sig") as f:
            f.write("\n".join(lines))
        self.set_status(f"已匯出 TXT 完整報告：{file_path} / 版本：{APP_VERSION}")
        messagebox.showinfo("完成", "TXT：全部完整報告匯出成功。")

    def export_csv_table(self):
        if not self.results:
            messagebox.showwarning("提醒", "目前沒有分析結果可匯出。")
            return
        file_path = filedialog.asksaveasfilename(title="CSV：主表資料", defaultextension=".csv", filetypes=[("CSV file", "*.csv"), ("All files", "*.*")])
        if not file_path:
            return
        fieldnames = [
            "排名", "燈號", "市場", "代號", "名稱", "顯示價", "漲跌", "漲跌幅%", "訊號", "建議",
            "分數", "等級", "目標價", "RR", "主升候選", "波浪定位", "費波位置", "禁追提示", "波費判定",
            "進場狀態", "進場可執行", "進場距離%", "追價風險", "建議倉位%", "建議股數", "風險預算%", "最大損失%", "資金等級", "配置分", "阻擋原因", "下單類型", "最終決策", "可下單", "決策原因", "Rule ID", "訊號原因", "RR等級", "波段分", "盤中分", "支撐", "壓力", "RSI", "五檔力道", "交易類型", "報價說明"
        ]
        with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(fieldnames)
            for idx, r in enumerate(self.results, start=1):
                writer.writerow([
                    idx, r["light"], r["market"], r["input_symbol"], r["name"], r["display_price"], f"{r['change']:+.2f}",
                    f"{r['change_pct']:+.2f}%", r["signal"], r.get("display_advice", r["advice"]), r["score"], r["strategy_level"], r["display_target_price"], r["display_rr"], r["leader_candidate"],
                    r.get("wave_stage", "-"), r.get("fibo_position", "-"), "是" if (r.get("fibo_risk_flag") or r.get("wave_risk_flag")) else "否", r.get("wave_fibo_signal", "-"),
                    r.get("entry_zone_status", "-"), r.get("entry_zone_ready", "-"), r.get("distance_to_entry_pct", "-"), r.get("chase_risk_flag", "-"), r.get("position_size_pct", 0), r.get("suggested_shares", 0), r.get("risk_budget_pct", 0), r.get("max_loss_pct", 0), r.get("allocation_grade", "-"), r.get("allocation_score", 0), r.get("final_block_reason", r.get("phase4_block_reason", "-")), r.get("order_type_hint", "-"), r.get("final_decision", "-"), r.get("execution_ready", "-"), r.get("decision_reason", "-"), r.get("rule_id", "-"), r.get("signal_reason", "-"), r.get("rr_level", "-"),
                    r["trend_score"], r["intraday_score"], r["support"], r["resistance"], r["rsi"],
                    r["orderbook_bias"], r.get("display_trade_type", r["trade_type"]), r["display_note"]
                ])
        self.set_status(f"已匯出 CSV 主表：{file_path} / 版本：{APP_VERSION}")
        messagebox.showinfo("完成", "CSV：主表資料匯出成功。")

    # Backward-compatible wrappers
    def export_txt(self):
        self.export_txt_full()

    def export_pdf(self):
        self.export_pdf_summary()

def validate_phase4_decision_rules():
    """Phase4 A01-A12 防回歸驗收。"""
    cases = [
        {
            "id": "A01",
            "name": "BUY 必須進場區有效",
            "data": {"price_valid": True, "signal": "主升突破", "advice": "拉回加碼", "state_bucket": "strong", "rr": 2.1, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "IN_ZONE", "position_size_pct": 8, "allocation_score": 85, "allocation_grade": "A", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "BUY",
        },
        {
            "id": "A02",
            "name": "ABOVE_ENTRY不可下單",
            "data": {"price_valid": True, "signal": "強勢追蹤", "advice": "拉回加碼", "state_bucket": "strong", "rr": 2.2, "rr_valid": True, "entry_zone_ready": False, "entry_zone_status": "ABOVE_ENTRY", "position_size_pct": 0, "allocation_score": 82, "allocation_grade": "B", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "WAIT",
        },
        {
            "id": "A03",
            "name": "BUY必須倉位大於0",
            "data": {"price_valid": True, "signal": "主升突破", "advice": "拉回加碼", "state_bucket": "strong", "rr": 2.3, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "IN_ZONE", "position_size_pct": 0, "allocation_score": 90, "allocation_grade": "A", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "WAIT",
        },
        {
            "id": "A04",
            "name": "禁追不可BUY",
            "data": {"price_valid": True, "signal": "末升/禁追風險", "advice": "不追高", "state_bucket": "weak", "rr": 3.0, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "NO_CHASE", "position_size_pct": 8, "allocation_score": 90, "allocation_grade": "A", "fibo_risk_flag": True, "wave_risk_flag": False},
            "expected": "AVOID",
        },
        {
            "id": "A05",
            "name": "RR不足不可下單",
            "data": {"price_valid": True, "signal": "強勢追蹤", "advice": "拉回加碼", "state_bucket": "strong", "rr": 1.2, "rr_valid": False, "entry_zone_ready": True, "entry_zone_status": "IN_ZONE", "position_size_pct": 3, "allocation_score": 80, "allocation_grade": "B", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "WAIT",
        },
        {
            "id": "A06",
            "name": "非即時報價不可下單",
            "data": {"price_valid": False, "signal": "主升突破", "advice": "拉回加碼", "state_bucket": "strong", "rr": 3.0, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "IN_ZONE", "position_size_pct": 8, "allocation_score": 90, "allocation_grade": "A", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "WAIT",
        },
        {
            "id": "A07",
            "name": "Allocation低於70不可BUY",
            "data": {"price_valid": True, "signal": "主升突破", "advice": "拉回加碼", "state_bucket": "strong", "rr": 2.0, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "IN_ZONE", "position_size_pct": 4, "allocation_score": 60, "allocation_grade": "C", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "WAIT",
        },
        {
            "id": "A08",
            "name": "BLOCK必須AVOID",
            "data": {"price_valid": True, "signal": "強勢追蹤", "advice": "拉回加碼", "state_bucket": "strong", "rr": 2.5, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "IN_ZONE", "position_size_pct": 8, "allocation_score": 0, "allocation_grade": "BLOCK", "phase4_block_reason": "測試BLOCK", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "AVOID",
        },
        {
            "id": "A09",
            "name": "BROKEN必須AVOID",
            "data": {"price_valid": True, "signal": "跌破支撐", "advice": "減碼/防守", "state_bucket": "weak", "rr": 2.0, "rr_valid": True, "entry_zone_ready": False, "entry_zone_status": "BROKEN", "position_size_pct": 0, "allocation_score": 0, "allocation_grade": "BLOCK", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "AVOID",
        },
        {
            "id": "A10",
            "name": "BREAKOUT_CONFIRM小倉可BUY",
            "data": {"price_valid": True, "signal": "突破強勢", "advice": "突破可追", "state_bucket": "strong", "rr": 2.0, "rr_valid": True, "entry_zone_ready": True, "entry_zone_status": "BREAKOUT_CONFIRM", "position_size_pct": 3, "allocation_score": 75, "allocation_grade": "B", "fibo_risk_flag": False, "wave_risk_flag": False},
            "expected": "BUY",
        },
    ]
    for case in cases:
        decision = build_final_decision(case["data"])
        assert decision["final_decision"] == case["expected"], f"Phase4驗收失敗：{case['id']} {case['name']}，得到 {decision['final_decision']}"
        if decision["final_decision"] == "BUY":
            assert decision["execution_ready"] is True, "BUY 必須 execution_ready=True"
            assert case["data"].get("entry_zone_ready") is True, "BUY 必須 entry_zone_ready=True"
            assert case["data"].get("position_size_pct", 0) > 0, "BUY 必須 position_size_pct>0"
            assert case["data"].get("allocation_score", 0) >= MIN_BUY_ALLOCATION_SCORE, "BUY 必須 allocation_score達標"
        else:
            assert decision["execution_ready"] is False, "非BUY不可 execution_ready=True"
            assert decision.get("position_size_pct", 0) == 0, "非BUY倉位必須歸零"
    return True

def main():
    validate_phase4_decision_rules()
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    GTCProApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
