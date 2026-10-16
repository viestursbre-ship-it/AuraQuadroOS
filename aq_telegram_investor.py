# -*- coding: utf-8 -*-
import os
import json
import telebot
from telebot import types
from datetime import datetime, timedelta
import yfinance as yf
import google.generativeai as genai
import requests

# --- IELĀDĒJAM CONFIG UN ATSLĒGAS ---
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
PORTFOLIO_PATH = os.path.join(os.path.dirname(__file__), "portfolio.json")

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

config = load_config()

# Telegram token
tg_section = config.get("telegram", {})
if isinstance(tg_section, dict):
    TELEGRAM_TOKEN = tg_section.get("bot_token") or tg_section.get("token")
else:
    TELEGRAM_TOKEN = None

if not TELEGRAM_TOKEN:
    TELEGRAM_TOKEN = (
        config.get("telegram_token") 
        or config.get("bot_token") 
        or config.get("telegram_bot_token")
        or os.environ.get("TELEGRAM_BOT_TOKEN")
    )

if not TELEGRAM_TOKEN:
    raise ValueError("⚠️ Telegram tokens netika atrasts config.json failā!")

# Gemini atslēga
GEMINI_KEY = (
    config.get("gemini_api_key") 
    or config.get("aq_ai_api_key") 
    or os.environ.get("GEMINI_API_KEY") 
    or os.environ.get("AQ_AI_API_KEY", "")
)

# Finnhub atslēga
FINNHUB_KEY = config.get("finnhub_api_key") or os.environ.get("FINNHUB_API_KEY", "")

bot = telebot.TeleBot(TELEGRAM_TOKEN)

# --- PALĪGFUNKCIJAS ---
def get_portfolio_tickers():
    if os.path.exists(PORTFOLIO_PATH):
        try:
            with open(PORTFOLIO_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
                return data.get("tickers", ["MSFT", "ARM", "VOO"])
        except Exception:
            pass
    return ["MSFT", "ARM", "VOO"]

def build_portfolio_summary():
    tickers = get_portfolio_tickers()
    text = "💼 *AuraQuadro Investor Studio — Portfelis*\n"
    text += "──────────────────────────\n"
    
    for t in tickers:
        try:
            tk = yf.Ticker(t)
            price = tk.fast_info.get("lastPrice") or tk.info.get("currentPrice") or 0.0
            currency = tk.fast_info.get("currency") or "USD"
            text += f"• *{t}*: `{price:.2f} {currency}`\n"
        except Exception:
            text += f"• *{t}*: _Nav datu_\n"
            
    text += "──────────────────────────\n"
    text += "Statuss: 🟢 Dati sinhronizēti ar tirgu"
    return text

def build_dividend_summary():
    tickers = get_portfolio_tickers()
    text = "💰 *Dividenžu un Ienesīguma Pārskats*\n"
    text += "──────────────────────────\n"
    
    for t in tickers:
        try:
            tk = yf.Ticker(t)
            info = tk.info
            
            # Sakārtojam ienesīgumu (precīza normalizācija)
            yield_rate = info.get("dividendYield")
            if yield_rate is not None and yield_rate > 0:
                if yield_rate > 20.0:
                    val = yield_rate / 100.0
                elif yield_rate < 0.20:
                    val = yield_rate * 100.0
                else:
                    val = yield_rate
                yield_str = f"{val:.2f}%"
            else:
                yield_str = "Nav/0.0%"
            
            # Ex-dividend datums
            ex_div = info.get("exDividendDate")
            if ex_div and isinstance(ex_div, (int, float)):
                try:
                    ex_date_str = datetime.fromtimestamp(ex_div).strftime("%Y-%m-%d")
                except Exception:
                    ex_date_str = str(ex_div)[:10]
            elif isinstance(ex_div, str):
                ex_date_str = ex_div[:10]
            else:
                ex_date_str = "Nav fiksēts"
                
            text += f"• *{t}*: Yield: `{yield_str}` | Ex-Date: `{ex_date_str}`\n"
        except Exception:
            text += f"• *{t}*: _Dati nav pieejami_\n"
            
    text += "──────────────────────────"
    return text

def get_coingecko_global():
    """Iegūst globālo kripto likviditātes pulsu bez API atslēgas."""
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get("https://api.coingecko.com/api/v3/global", headers=headers, timeout=5)
        if r.status_code == 200:
            d = r.json().get("data", {})
            total_mcap = d.get("total_market_cap", {}).get("usd", 0) / 1e12
            change_24h = d.get("market_cap_change_percentage_24h_usd", 0.0)
            btc_dom = d.get("market_cap_percentage", {}).get("btc", 0.0)
            sign = "+" if change_24h >= 0 else ""
            return (
                f"• *Kopējā likviditāte:* `${total_mcap:.2f}T` ({sign}{change_24h:.2f}% 24h)\n"
                f"• *BTC Dominance:* `{btc_dom:.1f}%`"
            )
    except Exception:
        pass
    return "• *Globālā likviditāte:* _Dati nav pieejami_"

def get_market_sentiment():
    """Iegūst tirgus Baiļu un Alkatības rādītāju + Makro pulsu + Kripto likviditāti."""
    text = "🌡️ *TIRGUS KLIMATS & GLOBĀLĀ LIKVIDITĀTE*\n──────────────────────────\n"
    
    # 1. Fear & Greed
    try:
        url = "https://api.alternative.me/fng/?limit=1"
        res = requests.get(url, timeout=5).json()
        item = res["data"][0]
        val = item["value"]
        status = item["value_classification"]
        text += f"• *Fear & Greed:* `{val}/100` ({status})\n"
    except Exception:
        text += "• *Fear & Greed:* _Nav pieejams_\n"

    # 2. Makro pulss
    text += "\n🌐 *Valsts Kase & Drošība:*\n"
    macro_map = {"^TNX": ("ASV 10Y Yield", "%"), "GC=F": ("Zelts", "USD")}
    for sym, (name, unit) in macro_map.items():
        try:
            tk = yf.Ticker(sym)
            p = tk.fast_info.get("lastPrice") or tk.info.get("regularMarketPrice") or 0.0
            if unit == "%":
                text += f"• *{name}*: `{p:.2f}%`\n"
            else:
                text += f"• *{name}*: `{p:,.2f} {unit}`\n"
        except Exception:
            text += f"• *{name}*: _Nav datu_\n"

    # 3. CoinGecko globālais likviditātes bloks
    text += "\n🪙 *Kripto & Naudas Plūsma (CoinGecko):*\n"
    text += get_coingecko_global() + "\n"
    text += "──────────────────────────"
    return text

def get_finnhub_intel():
    """Iegūst insaideru noskaņojumu un svaigas ziņas no Finnhub."""
    if not FINNHUB_KEY:
        return "⚠️ Finnhub atslēga nav pievienota config.json failā (`finnhub_api_key`)!"
        
    tickers = [t for t in get_portfolio_tickers() if "." not in t][:4] # Paņemam pirmos ASV tikerus
    text = "🕵️ *Finnhub Tirgus Izlūkdati (Insaideri & Ziņas)*\n"
    text += "──────────────────────────\n"
    
    # 1. Insaideru noskaņojums (MSPR)
    text += "📊 *Insaideru noskaņojums (Pēdējais mēnesis):*\n"
    for t in tickers:
        try:
            url = f"https://finnhub.io/api/v1/stock/insider-sentiment?symbol={t}&token={FINNHUB_KEY}"
            res = requests.get(url, timeout=5).json()
            data = res.get("data", [])
            if data:
                latest = data[-1]
                mspr = latest.get("mspr", 0)
                change = latest.get("change", 0)
                signal = "🟢 Pērk" if mspr > 0 else "🔴 Pārdod" if mspr < 0 else "⚪ Neitrāls"
                text += f"• *{t}*: {signal} (MSPR: `{mspr:.2f}`, Net: `{change}`)\n"
            else:
                text += f"• *{t}*: _Nav nesenu darījumu_\n"
        except Exception:
            text += f"• *{t}*: _Kļūda datos_\n"
            
    text += "──────────────────────────\n"
    
    # 2. Viena svarīga jaunākā ziņa pirmajam tikerim
    if tickers:
        lead_ticker = tickers[0]
        try:
            today_str = datetime.now().strftime("%Y-%m-%d")
            week_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
            news_url = f"https://finnhub.io/api/v1/company-news?symbol={lead_ticker}&from={week_ago}&to={today_str}&token={FINNHUB_KEY}"
            n_res = requests.get(news_url, timeout=5).json()
            if n_res and len(n_res) > 0:
                first_news = n_res[0]
                headline = first_news.get("headline", "Bez virsraksta")
                news_src = first_news.get("source", "Avots")
                text += f"📰 *Top ziņa ({lead_ticker}):*\n_{headline}_ ({news_src})\n"
        except Exception:
            pass
            
    text += "──────────────────────────"
    return text

def run_quark_audit():
    if not GEMINI_KEY:
        return "⚠️ Nav atrasta Gemini API atslēga config.json failā!"
        
    tickers = get_portfolio_tickers()
    market_snapshot = []
    for t in tickers:
        try:
            tk = yf.Ticker(t)
            price = tk.fast_info.get("lastPrice") or 0.0
            market_snapshot.append(f"{t}: {price:.2f}")
        except Exception:
            market_snapshot.append(f"{t}: N/A")
            
    prompt = f"""Tu esi Komandants Kvarks — ass, precīzs un stratēģisks AQ-OS finanšu AI asistents.
Izvērtē šos portfeļa aktīvus:
{', '.join(market_snapshot)}

Dod īsu militāri precīzu auditu:
1. Galvenais risks vai stiprā puse.
2. Īss taktisks ieteikums pozīciju līdzsvarošanai vai brīvajai naudai.
Maksimums 4-6 teikumi, skaidrā latviešu valodā ar enerģisku toni."""

    try:
        genai.configure(api_key=GEMINI_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")
        res = model.generate_content(prompt)
        return f"🐾 *KOMANDANTA KVARKA AUDITS:*\n──────────────────────────\n{res.text}"
    except Exception as e:
        return f"⚠️ Kvarka audita kļūda: {e}"

# --- TIKERU SAGLABĀŠANA UN PĀRVALDĪBA ---
def save_portfolio_tickers(tickers):
    """Saglabā aktīvo tikeru sarakstu portfolio.json failā."""
    data = {"tickers": tickers}
    try:
        with open(PORTFOLIO_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        return True
    except Exception as e:
        print(f"Kļūda saglabājot portfeli: {e}")
        return False

# --- IZVĒLNE AR POGĀM ---
def main_menu_keyboard():
    markup = types.InlineKeyboardMarkup(row_width=2)
    b1 = types.InlineKeyboardButton("💼 Portfelis", callback_data="btn_portfolio")
    b2 = types.InlineKeyboardButton("💰 Dividendes", callback_data="btn_dividends")
    b3 = types.InlineKeyboardButton("🌡️ Tirgus Klimats", callback_data="btn_sentiment")
    b4 = types.InlineKeyboardButton("🕵️ Insaideri & Ziņas", callback_data="btn_intel")
    b5 = types.InlineKeyboardButton("🐾 Kvarka Audits", callback_data="btn_audit")
    b6 = types.InlineKeyboardButton("⚙️ Pārvaldīt Tikerus", callback_data="btn_manage")
    b7 = types.InlineKeyboardButton("🔄 Atjaunot", callback_data="btn_refresh")
    markup.add(b1, b2)
    markup.add(b3, b4)
    markup.add(b5)
    markup.add(b6, b7)
    return markup

# --- TELEGRAM HANDLERI ---
@bot.message_handler(commands=['start', 'menu'])
def send_welcome(message):
    welcome_text = (
        "⚡ *AuraQuadroOS Investor Studio (Mobilā Līnija)* ⚡\n\n"
        "Sveiks, Komandant! Sistēma ir gatava operatīvai uzraudzībai tieši no viedtālruņa.\n"
        "Izvēlies darbību:"
    )
    bot.send_message(message.chat.id, welcome_text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.callback_query_handler(func=lambda call: True)
def handle_callbacks(call):
    chat_id = call.message.chat.id
    
    if call.data == "btn_portfolio":
        bot.answer_callback_query(call.id, "Ielādē portfeli...")
        text = build_portfolio_summary()
        bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

    elif call.data == "btn_dividends":
        bot.answer_callback_query(call.id, "Aprēķina dividendes...")
        text = build_dividend_summary()
        bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

    elif call.data == "btn_sentiment":
        bot.answer_callback_query(call.id, "Nolasa tirgus noskaņojumu...")
        text = get_market_sentiment()
        bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

    elif call.data == "btn_intel":
        bot.answer_callback_query(call.id, "Ielādē Finnhub datus...")
        text = get_finnhub_intel()
        bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

    elif call.data == "btn_audit":
        bot.answer_callback_query(call.id, "Kvarks analizē situāciju...")
        bot.send_message(chat_id, "🐾 *Kvarks pieslēdzas analīzei...*", parse_mode="Markdown")
        text = run_quark_audit()
        bot.send_message(chat_id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

    elif call.data == "btn_refresh":
        bot.answer_callback_query(call.id, "Atjaunots!")
        bot.send_message(chat_id, "🔄 Izvēlne atjaunota:", reply_markup=main_menu_keyboard())

    elif call.data == "btn_manage":
        bot.answer_callback_query(call.id)
        tickers = get_portfolio_tickers()
        help_text = (
            "⚙️ *Portfeļa Tikeru Pārvaldība*\n"
            "──────────────────────────\n"
            f"Pašreizējie aktīvi:\n`{', '.join(tickers)}`\n\n"
            "Lai mainītu saturu, vienkārši ieraksti čatā:\n"
            "• `+ NVDA` — pievienot tikeru\n"
            "• `- USO` — izņemt tikeru\n"
            "• `set AAPL, NVDA, VOO` — pārrakstīt visu sarakstu"
        )
        bot.send_message(chat_id, help_text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.callback_query_handler(func=lambda call: call.data == "btn_manage")
def handle_manage_click(call):
    tickers = get_portfolio_tickers()
    help_text = (
        "⚙️ *Portfeļa Tikeru Pārvaldība*\n"
        "──────────────────────────\n"
        f"Aktīvie tikeri:\n`{', '.join(tickers)}`\n\n"
        "Lai mainītu sastāvu, ieraksti čatā:\n"
        "• `+ NVDA` — pievienot jaunu uzņēmumu\n"
        "• `- DAL` — izņemt uzņēmumu\n"
        "• `set AAPL, NVDA, MSFT` — iestatīt visu sarakstu no jauna"
    )
    bot.answer_callback_query(call.id)
    bot.send_message(call.message.chat.id, help_text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.message_handler(func=lambda msg: not msg.text.startswith('/'))
def handle_ticker_text_commands(message):
    txt = message.text.strip()
    tickers = get_portfolio_tickers()
    
    if txt.startswith('+'):
        new_ticker = txt[1:].strip().upper()
        if new_ticker and new_ticker not in tickers:
            tickers.append(new_ticker)
            save_portfolio_tickers(tickers)
            bot.reply_to(message, f"✅ Tikers *{new_ticker}* pievienots!\nAktīvi: `{', '.join(tickers)}`", parse_mode="Markdown")
        else:
            bot.reply_to(message, f"ℹ️ Tikers *{new_ticker}* jau ir portfelī vai nav derīgs.", parse_mode="Markdown")
            
    elif txt.startswith('-'):
        rem_ticker = txt[1:].strip().upper()
        if rem_ticker in tickers:
            tickers.remove(rem_ticker)
            save_portfolio_tickers(tickers)
            bot.reply_to(message, f"🗑️ Tikers *{rem_ticker}* izņemts!\nAktīvi: `{', '.join(tickers)}`", parse_mode="Markdown")
        else:
            bot.reply_to(message, f"⚠️ Tikers *{rem_ticker}* netika atrasts.", parse_mode="Markdown")
            
    elif txt.lower().startswith('set '):
        raw_list = txt[4:].replace(';', ',').split(',')
        cleaned = [t.strip().upper() for t in raw_list if t.strip()]
        if cleaned:
            save_portfolio_tickers(cleaned)
            bot.reply_to(message, f"🎯 Jaunais portfelis iestatīts:\n`{', '.join(cleaned)}`", parse_mode="Markdown")

if __name__ == "__main__":
    print("🚀 AQ Telegram Investor Bot veiksmīgi palaists!")
    print("Atver Telegram savā telefonā un nosūti savam botam /start")
    bot.infinity_polling()
