# -*- coding: utf-8 -*-
import os
import json
import telebot
from telebot import types
from datetime import datetime, timedelta
import yfinance as yf
import google.generativeai as genai
import requests
from pathlib import Path

# Import for Word document generation
from docx import Document
from docx.shared import Inches, Pt 
from docx.enum.text import WD_ALIGN_PARAGRAPH 

# Import for PDF text extraction
import PyPDF2

# --- IELĀDĒJAM CONFIG UN ATSLĒGAS ---
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
PORTFOLIO_PATH = os.path.join(os.path.dirname(__file__), "portfolio.json")
INBOX_DIR = Path(os.path.dirname(__file__)) / "storage" / "inbox" # Definējam iesūtnes mapi
OUTBOX_DIR = Path(os.path.dirname(__file__)) / "storage" / "outbox" # Definējam izejošo failu mapi

# Pārliecināmies, ka INBOX_DIR un OUTBOX_DIR eksistē
INBOX_DIR.mkdir(parents=True, exist_ok=True)
OUTBOX_DIR.mkdir(parents=True, exist_ok=True)

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

def _extract_text_from_docx(docx_path: Path) -> str:
    """Ekstrahē tekstu no .docx faila."""
    document = Document(docx_path)
    full_text = []
    for para in document.paragraphs:
        full_text.append(para.text)
    return "\n".join(full_text)

def _extract_text_from_pdf(pdf_path: Path) -> str:
    """Ekstrahē tekstu no .pdf faila."""
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text += page.extract_text() or ""
    except Exception as e:
        return f"Kļūda, ekstrahējot tekstu no PDF: {e}"
    return text

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
        model = genai.GenerativeModel("gemini-2.5-flash") # Šis modelis jau ir korekts
        res = model.generate_content(prompt)
        return f"🐾 *KOMANDANTA KVARKA AUDITS:*\n──────────────────────────\n{res.text}"
    except Exception as e:
        return f"⚠️ Kvarka audita kļūda: {e}"

def generate_word_report(original_file_name: str, report_data: dict) -> Path:
    """
    Generates a Word (.docx) report from the analysis data.
    Saves it to the OUTBOX_DIR and returns the path.
    """
    document = Document()
    
    # Title
    document.add_heading(f"AQ-OS Juridiskais Audits: {original_file_name}", level=1)
    
    # Risk Level
    document.add_heading("Riska Līmenis:", level=2)
    document.add_paragraph(report_data.get('risk_level', 'Nav definēts'))
    
    # Contract Subject and Parties
    document.add_heading("Līguma Priekšmets un Puses:", level=2)
    document.add_paragraph(report_data.get('contract_subject_parties', 'Nav definēts'))
    
    # Main Risks and Pitfalls
    document.add_heading("Galvenie Riski un Zemūdens Akmeņi:", level=2)
    if report_data.get('main_risks_pitfalls'):
        for risk in report_data['main_risks_pitfalls']:
            document.add_paragraph(risk, style='List Bullet')
    else:
        document.add_paragraph("Nav būtisku risku fiksēts.")
        
    # Date Warnings
    document.add_heading("Datumu Brīdinājumi:", level=2)
    if report_data.get('date_warnings'):
        for warning in report_data['date_warnings']:
            document.add_paragraph(warning, style='List Bullet')
    else:
        document.add_paragraph("Nav datumu brīdinājumu fiksēts.")
        
    # Summary
    document.add_heading("Kopsavilkums:", level=2)
    document.add_paragraph(report_data.get('summary', 'Nav kopsavilkuma pieejams.'))
    
    # Footer
    document.add_paragraph("\n──────────────────────────")
    document.add_paragraph(f"Auditu veikusi AuraQuadroOS Juridiskā Analīzes Sistēma, vadībā Komandants Kvarks. Izveidots: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Save the document
    # Get the base name without extension for a clean report filename
    base_name = Path(original_file_name).stem
    report_filename = f"DIGEST_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{base_name}.docx"
    output_path = OUTBOX_DIR / report_filename
    document.save(output_path)
    return output_path

def doc_digest_analysis(file_path: Path):
    """
    Analizē dokumentu, izmantojot Gemini AI un ģenerē Word failu ar rezultātiem.
    Ekstrahē tekstu no PDF/DOCX faila.
    Atgriež analīzes rezultātus kā dict un ģenerētā Word faila ceļu.
    """
    docx_file_path = None # Initialize as None
    document_content = "" # Initialize document content

    if not GEMINI_KEY:
        error_results = {
            "risk_level": "⚠️ Kritiskas Kļūda",
            "contract_subject_parties": "Nav definēts bez AI.",
            "main_risks_pitfalls": ["Gemini API atslēga nav konfigurēta."],
            "date_warnings": [],
            "summary": "Nevar veikt dokumenta analīzi bez API atslēgas. Lūdzu, konfigurējiet to."
        }
        try:
            docx_file_path = generate_word_report(file_path.name, error_results)
        except Exception as e:
            print(f"Kļūda ģenerējot Word atskaiti pie API kļūdas: {e}")
        return error_results, docx_file_path
    
    try:
        genai.configure(api_key=GEMINI_KEY)
        
        # Mēģinām izmantot gemini-2.5-flash, ar atkritienu uz gemini-2.0-flash
        try:
            model = genai.GenerativeModel("gemini-2.5-flash")
        except Exception:
            print("WARNING: gemini-2.5-flash nav pieejams, izmantoju gemini-2.0-flash.")
            model = genai.GenerativeModel("gemini-2.0-flash") # Atkritiens uz stabilu modeli
        
        # --- Real document text extraction ---
        try:
            if file_path.suffix == '.docx':
                document_content = _extract_text_from_docx(file_path)
            elif file_path.suffix == '.pdf':
                document_content = _extract_text_from_pdf(file_path)
            else:
                document_content = f"Neatbalstīts faila formāts ({file_path.suffix}). Nevar ekstrahēt tekstu."
                print(f"ERROR: {document_content}")
        except Exception as e:
            document_content = f"Kļūda, ekstrahējot tekstu no faila '{file_path.name}': {e}. Lūdzu, pārbaudiet failu."
            print(f"ERROR: {document_content}")

        if not document_content.strip():
            document_content = f"Dokuments '{file_path.name}' ir tukšs vai no tā nevarēja ekstrahēt tekstu. AI analizēs šo situāciju."
        
        prompt = f"""Tu esi "AQ-OS Juridiskais Analītiķis", Komandanta Kvarka labā roka dokumentu auditā.
Analizē šo dokumentu ({file_path.name}) saturu un struktūru, ņemot vērā tālāk sniegto tekstu.
Ja teksts ir īss vai tukšs, informē par to kopsavilkumā.

Dokumenta saturs:
---
{document_content}
---

Nodrošini strukturētu, militāri precīzu un kodolīgu atskaiti latviešu valodā. Izcel galvenos punktus un riskus.
Formāts, kas JĀIEVĒRO Dzelžaini:
Riska Līmenis: [Zems/Vidējs/Augsts/Kritisks/Nav definēts]
Līguma Priekšmets un Puses: [Īss apraksts, piem., "Investīciju līgums starp X un Y". Ja nav skaidri definēts, norādi "Nav skaidri definēts."]
Galvenie Riski un Zemūdens Akmeņi:
- [Specifisks risks, piem., "Automātiska pagarināšana"]
- [Specifisks risks, piem., "Slēptās izmaksas"]
- [Specifisks risks, piem., "Sankciju klauzulas"]
- [Ja nav būtisku risku, norādi tikai "Nav būtisku risku fiksēts."]
Datumu Brīdinājumi:
- [Datums un problēma, piem., "Termiņš pagājis: 2023-01-01"]
- [Datums un problēma, piem., "Neloģisks termiņš: sākums pēc beigām"]
- [Ja nav datumu brīdinājumu, norādi tikai "Nav datumu brīdinājumu fiksēts."]
Kopsavilkums: [Kodolīgs kopsavilkums cilvēcīgā valodā, ne vairāk kā 3-4 teikumi. Ja dokuments ir tukšs vai neizdevās ekstrahēt tekstu, norādi to.]
"""
        
        response = model.generate_content(prompt)
        response_text = response.text.strip()
        
        parsed_response = {
            "risk_level": "Neskaidrs",
            "contract_subject_parties": "Nav definēts.",
            "main_risks_pitfalls": [],
            "date_warnings": [],
            "summary": "Neizdevās analizēt Gemini atbildi."
        }
        
        lines = response_text.split('\n')
        current_section = None
        
        for line in lines:
            line = line.strip()
            if not line:
                continue

            if line.startswith("Riska Līmenis:"):
                parsed_response["risk_level"] = line.replace("Riska Līmenis:", "").strip()
                current_section = None
            elif line.startswith("Līguma Priekšmets un Puses:"):
                parsed_response["contract_subject_parties"] = line.replace("Līguma Priekšmets un Puses:", "").strip()
                current_section = None
            elif line.startswith("Galvenie Riski un Zemūdens Akmeņi:"):
                current_section = "main_risks_pitfalls"
            elif line.startswith("Datumu Brīdinājumi:"):
                current_section = "date_warnings"
            elif line.startswith("Kopsavilkums:"):
                parsed_response["summary"] = line.replace("Kopsavilkums:", "").strip()
                current_section = None 
            elif line.startswith("- ") and current_section in ["main_risks_pitfalls", "date_warnings"]:
                item = line[2:].strip()
                # Filter out generic 'no risks' messages from actual lists
                if current_section == "main_risks_pitfalls":
                    if item.lower() != "nav būtisku risku fiksēts.": 
                        parsed_response[current_section].append(item)
                elif current_section == "date_warnings":
                    if item.lower() != "nav datumu brīdinājumu fiksēts.":
                        parsed_response[current_section].append(item)
            elif parsed_response["summary"] == "Neizdevās analizēt Gemini atbildi." and current_section is None and not line.startswith('-'):
                parsed_response["summary"] = line # Fallback if summary is not parsed initially

        # Refined summary fallback if still default
        if parsed_response["summary"] == "Neizdevās analizēt Gemini atbildi.":
            if "Kļūda, ekstrahējot tekstu" in document_content:
                parsed_response["summary"] = "Kļūda, ekstrahējot tekstu no dokumenta. Analīze var būt nepilnīga."
            elif document_content.startswith("Dokuments") and ("tukšs" in document_content or "nevarēja ekstrahēt tekstu" in document_content):
                parsed_response["summary"] = "Dokuments ir tukšs vai no tā nevarēja ekstrahēt tekstu. AI atbildēja uz šo situāciju."
            else:
                parsed_response["summary"] = "Neizdevās iegūt kopsavilkumu no AI atbildes."

        # Ensure contract_subject_parties is not empty if summary suggests there was content.
        if parsed_response["contract_subject_parties"] == "Nav definēts." and parsed_response["summary"] != "Dokuments ir tukšs vai no tā nevarēja ekstrahēt tekstu. AI atbildēja uz šo situāciju.":
             parsed_response["contract_subject_parties"] = "Nav skaidri definēts."
        
        # Pamata pārbaudes, ja riska līmenis nav skaidri definēts
        if parsed_response["risk_level"] == "Neskaidrs":
            if parsed_response["main_risks_pitfalls"] or parsed_response["date_warnings"]:
                parsed_response["risk_level"] = "Vidējs" 
            else:
                parsed_response["risk_level"] = "Zems"

        # Generate the Word document
        try:
            docx_file_path = generate_word_report(file_path.name, parsed_response)
        except Exception as e:
            print(f"Kļūda ģenerējot Word atskaiti: {e}")
        
        return parsed_response, docx_file_path
    
    except Exception as e:
        error_summary_msg = f"Neizdevās veikt dokumenta analīzi. Sazinieties ar Komandantu. Detalizēta kļūda: {e}"
        if "Kļūda, ekstrahējot tekstu" in document_content:
            error_summary_msg = f"Kļūda, ekstrahējot tekstu no dokumenta, vai citas analīzes kļūdas dēļ: {e}. Sazinieties ar Komandantu."
        
        error_results = {
            "risk_level": "🔴 Kritiskas Kļūda",
            "contract_subject_parties": "Nav definēts kļūdas dēļ.",
            "main_risks_pitfalls": [f"Gemini analīzes kļūda: {e}"],
            "date_warnings": [],
            "summary": error_summary_msg
        }
        # Attempt to generate a Word file even with error results
        try:
            docx_file_path = generate_word_report(file_path.name, error_results)
        except Exception as e_docx:
            print(f"Kļūda ģenerējot Word atskaiti pie analīzes kļūdas: {e_docx}")
        return error_results, docx_file_path

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
    b6 = types.InlineKeyboardButton("📑 Līgumu Audits", callback_data="btn_contract_audit") # JAUNA POGA
    b7 = types.InlineKeyboardButton("⚙️ Pārvaldīt Tikerus", callback_data="btn_manage")
    b8 = types.InlineKeyboardButton("🔄 Atjaunot", callback_data="btn_refresh")
    markup.add(b1, b2)
    markup.add(b3, b4)
    markup.add(b5, b6) # Pievienojam jauno pogu esošajā rindā
    markup.add(b7, b8)
    return markup

# --- TELEGRAM HANDLERI ---

# Jauna komanda /boss un /kvarks
@bot.message_handler(commands=['boss', 'kvarks'])
def send_boss_status(message):
    boss_message = (
        "🐾 *KOMANDANTS KVARKS & GALVENAIS ANALĪTIĶIS LĀCIS ziņo:*\n"
        "──────────────────────────\n"
        "Liepājas garnizons ir pilnībā nocietināts un operatīvā gatavībā! "
        "Visas konservu fondu noliktavas tiek pastāvīgi uzraudzītas, "
        "nodrošinot kritisko resursu drošību. Portfelis ir stabilā stāvoklī "
        "un tiek rūpīgi pārvaldīts. "
        "Miers un drošība ir prioritāte. Turpinām darbu!"
    )
    bot.send_message(message.chat.id, boss_message, parse_mode="Markdown", reply_markup=main_menu_keyboard())

# Komanda /help
@bot.message_handler(commands=['help'])
def send_help_message(message):
    help_text = (
        "⚡ *AuraQuadroOS Komandu Saraksts* ⚡\n"
        "──────────────────────────\n"
        "*/start* vai */menu* — Atver galveno izvēlni.\n"
        "*/help* — Parāda šo komandu sarakstu.\n"
        "*/boss* vai */kvarks* — Saņem ziņojumu no Komandanta Kvarka un Galvenā analītiķa Lāča.\n"
        "*/portfolio* — Parāda portfeļa kopsavilkumu.\n"
        "*/dividends* — Parāda dividenžu pārskatu.\n"
        "*/sentiment* — Parāda tirgus noskaņojumu un globālo likviditāti.\n"
        "*/intel* — Parāda Finnhub insaideru datus un jaunākās ziņas.\n"
        "*/audit* — Iesniedz pieprasījumu Komandanta Kvarka auditam (ar AI).\n"
        "*/manage* — Atver tikeru pārvaldības saskarni.\n"
        "*/refresh* — Atjauno galveno izvēlni.\n\n"
        "Pārvaldības komandas čatā:\n"
        "`+ TIKER` — Pievienot tikeru (piem., `+ NVDA`)\n"
        "`- TIKER` — Izņemt tikeru (piem., `- DAL`)\n"
        "`set TIKER1, TIKER2` — Iestatīt jaunu portfeļa sarakstu (piem., `set AAPL, MSFT, GOOGL`)\n"
        "Sūtiet `.pdf` vai `.docx` dokumentu, lai veiktu juridisko auditu."
    )
    bot.send_message(message.chat.id, help_text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

# Tiešās komandas pogu funkcionalitātei
@bot.message_handler(commands=['portfolio'])
def command_portfolio(message):
    text = build_portfolio_summary()
    bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.message_handler(commands=['dividends'])
def command_dividends(message):
    text = build_dividend_summary()
    bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.message_handler(commands=['sentiment'])
def command_sentiment(message):
    text = get_market_sentiment()
    bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.message_handler(commands=['intel'])
def command_intel(message):
    text = get_finnhub_intel()
    bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.message_handler(commands=['audit'])
def command_audit(message):
    bot.send_message(message.chat.id, "🐾 *Kvarks pieslēdzas analīzei...*", parse_mode="Markdown")
    text = run_quark_audit()
    bot.send_message(message.chat.id, text, parse_mode="Markdown", reply_markup=main_menu_keyboard())

@bot.message_handler(commands=['manage'])
def command_manage(message):
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
    bot.send_message(message.chat.id, help_text, parse_mode="Markdown", reply_markup=main_menu_keyboard())
    
@bot.message_handler(commands=['refresh'])
def command_refresh(message):
    bot.send_message(message.chat.id, "🔄 Izvēlne atjaunota:", reply_markup=main_menu_keyboard())


@bot.message_handler(commands=['start', 'menu'])
def send_welcome(message):
    welcome_text = (
        "⚡ *AuraQuadroOS Investor Studio (Mobilā Līnija)* ⚡\n\n"
        "Sveiks, Komandant! Sistēma ir gatava operatīvai uzraudzībai tieši no viedtālruņa.\n"
        "Izvēlies darbību vai ieraksti */help*, lai redzētu komandu sarakstu:"
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
    
    # JAUNS HANDLERIS LĪGUMU AUDITA POGAI
    elif call.data == "btn_contract_audit":
        bot.answer_callback_query(call.id, "Komandants Kvarks gaida dokumentu skenēšanai...")
        instruction_message = (
            "📎 Lai veiktu līguma auditu, piespiediet saspraudes ikonu apakšā un nosūtiet "
            "`.pdf` vai `.docx` dokumentu. Komandants Kvarks to nekavējoties izskatīs!"
        )
        bot.send_message(chat_id, instruction_message, parse_mode="Markdown", reply_markup=main_menu_keyboard())

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

# --- JAUNS DOKUMENTU HANDLERIS ---
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    chat_id = message.chat.id
    
    # Initialize docx_file_path to None, so we know if it was successfully generated
    docx_file_path = None

    try:
        if not message.document:
            bot.send_message(chat_id, "⚠️ Kļūda: Fails netika atrasts ziņojumā.")
            return

        file_name = message.document.file_name
        file_extension = Path(file_name).suffix.lower() # Izmantojam Path, lai tīri iegūtu paplašinājumu

        # Pārbaudām, vai fails ir .pdf vai .docx
        if file_extension not in ('.pdf', '.docx'):
            bot.reply_to(message, "⚠️ Saņemts fails nav atbalstīts. Lūdzu, sūtiet `.pdf` vai `.docx` dokumentus.", reply_markup=main_menu_keyboard())
            return

        bot.send_message(chat_id, "⏳ Dokuments saņemts! Tiek veikts juridiskais audits...")
        
        # Lejupielādējam failu un saglabājam to 'storage/inbox' mapē
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Izveidojam unikālu faila nosaukumu, lai izvairītos no pārrakstīšanas
        unique_filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file_name}"
        save_path = INBOX_DIR / unique_filename # Izmantojam Path objektu ceļu savienošanai
        
        with open(save_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        bot.send_message(chat_id, f"✅ Dokuments `{file_name}` saglabāts, sāku analīzi.")

        # Izsaucam doc_digest analīzes moduli šim failam
        report_data, docx_file_path = doc_digest_analysis(save_path) # Tagad atgriež divas vērtības
        
        # Map risk level to color emoji for presentation
        risk_map = {
            "Zems": "🟢 Zaļš",
            "Vidējs": "🟡 Dzeltens",
            "Augsts": "🟠 Oranžs", 
            "Kritisks": "🔴 Sarkans",
            "Nav definēts": "⚪ Nav definēts",
            "Neskaidrs": "⚪ Nezināms",
            "⚠️ Kritiskas Kļūda": "🔴 Kļūda"
        }
        
        display_risk = risk_map.get(report_data['risk_level'], f"⚪ {report_data['risk_level']}")

        # Build report caption for Telegram message (short summary)
        caption_text = (
            f"📄 *Juridiskais Audits* (DIGEST)\n"
            f"──────────────────────────\n"
            f"🚦 Riska līmenis: *{display_risk}*\n"
            f"📝 Kopsavilkums: {report_data['summary']}\n"
            f"──────────────────────────\n"
            f"Pilnu slēdzienu skatīt pievienotajā Word failā."
        )

        # Send the generated Word document
        if docx_file_path and docx_file_path.exists():
            with open(docx_file_path, 'rb') as doc_file:
                bot.send_document(chat_id, doc_file, caption=caption_text, parse_mode="Markdown")
            bot.send_message(chat_id, "✅ Word slēdziens nosūtīts!")
            # Optionally delete the generated docx file from outbox after sending
            # docx_file_path.unlink(missing_ok=True)
        else:
            bot.send_message(chat_id, "⚠️ Kļūda: Neizdevās ģenerēt vai atrast Word slēdziena failu. Sazinieties ar Komandantu.", parse_mode="Markdown")


        # Send the on-screen textual report.
        report_parts = []
        report_parts.append("📄 *DOKUMENTA JURIDISKAIS AUDITS (AQ-OS)*")
        report_parts.append("──────────────────────────")
        report_parts.append(f"🚦 Riska līmenis: *{display_risk}*")
        report_parts.append(f"📌 Līguma priekšmets un Puses: `{report_data['contract_subject_parties']}`")
        
        main_risks_text = ""
        if report_data['main_risks_pitfalls']:
            main_risks_text = "\n".join([f"- {cp}" for cp in report_data['main_risks_pitfalls']])
            report_parts.append(f"⚠️ Galvenie riski un 'zemūdens akmeņi':\n{main_risks_text}")
        
        date_warnings_text = ""
        if report_data['date_warnings']:
            date_warnings_text = "\n".join([f"- {dw}" for dw in report_data['date_warnings']])
            report_parts.append(f"📅 Datumu brīdinājumi:\n{date_warnings_text}")
            
        report_parts.append(f"📝 Kopsavilkums: {report_data['summary']}")
        report_parts.append("──────────────────────────")
        report_parts.append("Status: Audits pabeigts. Komandants Kvarks gaida turpmākās instrukcijas.")
        
        final_report = '\n\n'.join(report_parts) 

        MAX_TELEGRAM_MESSAGE_LENGTH = 4096
        
        if len(final_report) > MAX_TELEGRAM_MESSAGE_LENGTH:
            part1_sections = [
                "📄 *DOKUMENTA JURIDISKAIS AUDITS (AQ-OS)*",
                "──────────────────────────",
                f"🚦 Riska līmenis: *{display_risk}*",
                f"📌 Līguma priekšmets un Puses: `{report_data['contract_subject_parties']}`",
                f"📝 Kopsavilkums: {report_data['summary']}",
                "──────────────────────────",
                "Status: Audits pabeigts. (1/2)"
            ]
            part1 = '\n\n'.join(part1_sections)
            bot.send_message(chat_id, part1, parse_mode="Markdown")
            
            part2_sections = []
            if main_risks_text:
                part2_sections.append(f"⚠️ Galvenie riski un 'zemūdens akmeņi':\n{main_risks_text}")
            if date_warnings_text:
                part2_sections.append(f"📅 Datumu brīdinājumi:\n{date_warnings_text}")
            
            if part2_sections:
                part2_sections.insert(0, "📄 *AUDITA TURPINĀJUMS (2/2)*")
                part2_sections.append("──────────────────────────")
                part2_sections.append("Komandants Kvarks gaida turpmākās instrukcijas.")
                part2 = '\n\n'.join(part2_sections)
                
                if len(part2) > MAX_TELEGRAM_MESSAGE_LENGTH:
                    bot.send_message(chat_id, "⚠️ Papildu detaļas ir pārāk garas, lai nosūtītu visu. Sazinieties ar Komandanta štābu.", parse_mode="Markdown", reply_markup=main_menu_keyboard())
                else:
                    bot.send_message(chat_id, part2, parse_mode="Markdown", reply_markup=main_menu_keyboard())
            else:
                bot.send_message(chat_id, "ℹ️ Nav sīkāku risku vai brīdinājumu, lai turpinātu ziņojumu. Pilns slēdziens Word formātā nosūtīts.", parse_mode="Markdown", reply_markup=main_menu_keyboard())
        else:
            bot.send_message(chat_id, final_report, parse_mode="Markdown", reply_markup=main_menu_keyboard())

        # Pēc apstrādes varam izdzēst failu no inbox mapes, ja tas vairs nav nepieciešams
        # save_path.unlink(missing_ok=True) # Jāizlemj, vai dzēst. Drošības pēc atstājam iesūtnē.

    except Exception as e:
        error_message = (
            "⚠️ Kļūda apstrādājot dokumentu! "
            "Sazinieties ar Komandanta štābu par šo kritisko incidentu. "
            f"Tehniskā informācija: `{e}`"
        )
        bot.send_message(chat_id, error_message, parse_mode="Markdown", reply_markup=main_menu_keyboard())


if __name__ == "__main__":
    print("🚀 AQ Telegram Investor Bot veiksmīgi palaists!")
    print("Atver Telegram savā telefonā un nosūti savam botam /start vai /help")
    bot.infinity_polling()