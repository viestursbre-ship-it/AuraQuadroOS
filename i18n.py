# -*- coding: utf-8 -*-
import json
from pathlib import Path # Import Path for robust path handling

# --- Configuration Constants ---
CONFIG_PATH = Path(__file__).resolve().parent / "config.json" # Drošs ceļš uz config.json
DEFAULT_UI_LANGUAGE_CODE = 'lv' # Noklusējuma valoda lietotāja saskarnei
DEFAULT_OUTPUT_LANGUAGE_CODE = 'auto' # Noklusējuma valoda atskaitēm/ģenerētajiem failiem (vai 'lv', 'en', 'auto' u.c.')

# --- Supported Languages Definition ---
# Nodrošinātas tieši 5 atbalstītās valodas: "lv", "en", "ru", "uk", "it"
SUPPORTED_LANGUAGES = {
    'lv': {'name': 'Latviešu', 'prompt_lang': 'Latvian'},
    'en': {'name': 'English', 'prompt_lang': 'English'},
    'uk': {'name': 'Українська', 'prompt_lang': 'Ukrainian'},
    'ru': {'name': 'Русский', 'prompt_lang': 'Russian'},
    'it': {'name': 'Italiano', 'prompt_lang': 'Italian'}
}

# --- Language Code Mapping for Normalization ---
# Ja saņemts kods "ua", automātiski pārvērt to par "uk"
LANG_CODE_MAP = {"ua": "uk"}

def _normalize_lang_input(raw_lang: str) -> str:
    """
    Normalizē ievadīto valodas kodu, ieskaitot .strip().lower() un "ua" -> "uk" kartēšanu.
    Ja ievade nav virkne, atgriež noklusējuma UI valodu.
    """
    if not isinstance(raw_lang, str):
        return DEFAULT_UI_LANGUAGE_CODE # Atgriež noklusējumu, ja ievade nav virkne
    
    cleaned_lang = raw_lang.strip().lower()
    return LANG_CODE_MAP.get(cleaned_lang, cleaned_lang)

# --- Translations Dictionary ---
TRANSLATIONS = {
    "lv": {
        "tab_offers": "💼 Piedāvājumi",
        "tab_hp": "💻 HP Eksperts",
        "tab_voice": "🎙️ Balss Studija",
        "tab_dropzone": "📁 Dokumentu Drop-Zone",
        "tab_radar": "⚙️ Radara Iestatījumi",
        "btn_calc": "⚡ Aprēķināt & Noformēt",
        "btn_excel": "📊 Atvērt Excel Tāmēšanai",
        "btn_copy": "📋 Kopēt E-pastam",
        "lbl_client_req": "1. 📥 Klienta Pieprasījums:",
        "lbl_dist_costs": "2. 🏷️ Distributora Pašizmaksa / Cenas:",
        "lbl_result": "⚡ Gatavais Piedāvājums Klientam (Drošs):",
        "lbl_margin": "Marža (%):",
        "lbl_lang": "Valoda:",
        "status_ready": "Gatavs!",
        "status_calculating": "Gemini tāmē & ģenerē Excel...",
        # Offer Studio atslēgas, nodrošinātas visām valodām
        "btn_calculate": "⚡ Aprēķināt & Noformēt",
        "btn_open_excel": "📊 Atvērt Excel Tāmēšanai",
        "btn_copy_email": "📋 Kopēt E-pastam",
        "lbl_language": "Valoda:",
        "lbl_client_rfq": "1. 📥 Klienta Pieprasījums:",
        "btn_load_file": "📁 Ielādēt",
        "lbl_supplier_costs": "2. 🏷️ Distributora Pašizmaksa / Cenas:",
        "lbl_commercial_offer": "⚡ Gatavais Piedāvājums Klientam (Drošs):",
        # Bank Studio atslēgas
        "btn_load_bank_statement": "📁 Ielādēt Bankas Izrakstu (CSV/XLSX)",
        "btn_analyze_cashflow": "⚡ Analizēt Naudas Plūsmu",
        "lbl_bank_summary": "Finanšu Kopsavilkums & Rādītāji:",
        "lbl_incoming_funds": "Ienākošie līdzekļi:",
        "lbl_outgoing_funds": "Izejošie maksājumi:",
        "lbl_net_cashflow": "Neto naudas plūsma:",
        "lbl_bank_title": "🏦 Citadele Bank Studio (ISO 20022)",
        "btn_load_pdf_invoice": "📄 Ielādēt PDF Rēķinu",
        "lbl_bank_ready_status": "Gatavs darbam. Ielādējiet PDF rēķinu no e-pasta vai datora.",
        "btn_add_partner": "⭐ Pievienot partneri bāzei",
        "btn_generate_xml": "🚀 Ģenerēt Citadele XML",
        # Dokumentu Drop-Zone atslēgas
        "btn_drop_browse": "📄 Izvēlēties Dokumentu (PDF, DOCX, TXT)",
        "btn_summarize_doc": "⚡ Izvilkt Būtību & Kopsavilkumu",
        "lbl_doc_preview": "Dokumenta satura priekšskatījums:",
        "lbl_ai_summary": "AI Analīze & Kopsavilkums:",
        # Document Digest atslēgas
        "btn_select_doc": "📥 Izvēlēties Dokumentu",
        "btn_open_contracts": "📊 Atvērt Līgumu Reģistru",
        "btn_open_conclusion": "📄 Atvērt Slēdzienu",
        "lbl_doc_waiting_status": "0% - Gaida dokumentu",
        # Investor Studio papildu atslēgas
        "btn_audit_ticker": "⚡ Analizēt Tīkeri",
        "lbl_ticker_input": "Ievadiet Tīkeri (piem., AAPL, NVDA):",
    },
    "en": {
        "tab_offers": "💼 Quote Studio",
        "tab_hp": "💻 HP Expert",
        "tab_voice": "🎙️ Voice Studio",
        "tab_dropzone": "📁 Document Drop-Zone",
        "tab_radar": "⚙️ Radar Settings",
        "btn_calc": "⚡ Calculate & Generate",
        "btn_excel": "📊 Open Excel Sheet",
        "btn_copy": "📋 Copy for Email",
        "lbl_client_req": "1. 📥 Client RFQ / Requirements:",
        "lbl_dist_costs": "2. 🏷️ Supplier / Disti Costs:",
        "lbl_result": "⚡ Generated Commercial Quote (Client-Safe):",
        "lbl_margin": "Margin (%):",
        "lbl_lang": "Language:",
        "status_ready": "Ready!",
        "status_calculating": "Gemini calculating & generating Excel...",
        # Offer Studio atslēgas, nodrošinātas visām valodām
        "btn_calculate": "⚡ Calculate & Generate",
        "btn_open_excel": "📊 Open Excel Sheet",
        "btn_copy_email": "📋 Copy for Email",
        "lbl_language": "Language:",
        "lbl_client_rfq": "1. 📥 Client RFQ / Requirements:",
        "btn_load_file": "📁 Load",
        "lbl_supplier_costs": "2. 🏷️ Supplier / Disti Costs:",
        "lbl_commercial_offer": "⚡ Commercial Offer (Client-Safe):",
        # Bank Studio atslēgas
        "btn_load_bank_statement": "📁 Load Bank Statement (CSV/XLSX)",
        "btn_analyze_cashflow": "⚡ Analyze Cash Flow",
        "lbl_bank_summary": "Financial Summary & Metrics:",
        "lbl_incoming_funds": "Incoming funds:",
        "lbl_outgoing_funds": "Outgoing payments:",
        "lbl_net_cashflow": "Net cash flow:",
        "lbl_bank_title": "🏦 Citadele Bank Studio (ISO 20022)",
        "btn_load_pdf_invoice": "📄 Load PDF Invoice",
        "lbl_bank_ready_status": "Ready. Load a PDF invoice from email or PC.",
        "btn_add_partner": "⭐ Add Partner to Database",
        "btn_generate_xml": "🚀 Generate Citadele XML",
        # Dokumentu Drop-Zone atslēgas
        "btn_drop_browse": "📄 Select Document (PDF, DOCX, TXT)",
        "btn_summarize_doc": "⚡ Extract Digest & Summary",
        "lbl_doc_preview": "Document content preview:",
        "lbl_ai_summary": "AI Analysis & Summary:",
        # Document Digest atslēgas
        "btn_select_doc": "📥 Select Document",
        "btn_open_contracts": "📊 Open Contracts Register",
        "btn_open_conclusion": "📄 Open Report",
        "lbl_doc_waiting_status": "0% - Waiting for document",
        # Investor Studio papildu atslēgas
        "btn_audit_ticker": "⚡ Analyze Ticker",
        "lbl_ticker_input": "Enter Ticker (e.g., AAPL, NVDA):",
    },
    "uk": {
        "tab_offers": "💼 Студія пропозицій",
        "tab_hp": "💻 HP Експерт",
        "tab_voice": "🎙️ Голосова Студія",
        "tab_dropzone": "📁 Зона для документів",
        "tab_radar": "⚙️ Налаштування радару",
        "btn_calc": "⚡ Розрахувати та оформити",
        "btn_excel": "📊 Відкрити Excel для кошторису",
        "btn_copy": "📋 Копіювати для Email",
        "lbl_client_req": "1. 📥 Запит клієнта / Вимоги:",
        "lbl_dist_costs": "2. 🏷️ Собівартість дистриб'ютора:",
        "lbl_result": "⚡ Згенерована комерційна пропозиція (безпечна для клієнта):",
        "lbl_margin": "Маржа (%):",
        "lbl_lang": "Мова:",
        "status_ready": "Готово!",
        "status_calculating": "Gemini розраховує та генерує Excel...",
        # Offer Studio atslēgas, nodrošinātas visām valodām
        "btn_calculate": "⚡ Розрахувати та оформити",
        "btn_open_excel": "📊 Відкрити Excel для кошторису",
        "btn_copy_email": "📋 Копіювати для Email",
        "lbl_language": "Мова:",
        "lbl_client_rfq": "1. 📥 Запит клієнта:",
        "btn_load_file": "📁 Завантажити",
        "lbl_supplier_costs": "2. 🏷️ Собівартість дистри'ютора:",
        "lbl_commercial_offer": "⚡ Готова пропозиція (Client-Safe):",
        # Bank Studio atslēgas
        "btn_load_bank_statement": "📁 Завантажити виписку банку (CSV/XLSX)",
        "btn_analyze_cashflow": "⚡ Аналізувати грошовий потік",
        "lbl_bank_summary": "Фінансовий підсумок та показники:",
        "lbl_incoming_funds": "Надходження:",
        "lbl_outgoing_funds": "Вихідні платежі:",
        "lbl_net_cashflow": "Чистий грошовий потік:",
        "lbl_bank_title": "🏦 Citadele Bank Studio (ISO 20022)",
        "btn_load_pdf_invoice": "📄 Завантажити рахунок PDF",
        "lbl_bank_ready_status": "Готовий до роботи. Завантажте PDF рахунок з пошти або ПК.",
        "btn_add_partner": "⭐ Додати партнера до бази",
        "btn_generate_xml": "🚀 Створити Citadele XML",
        # Dokumentu Drop-Zone atslēgas
        "btn_drop_browse": "📄 Вибрати документ (PDF, DOCX, TXT)",
        "btn_summarize_doc": "⚡ Витягти суть та резюме",
        "lbl_doc_preview": "Попередній перегляд документа:",
        "lbl_ai_summary": "AI Аналіз та резюме:",
        # Document Digest atslēgas
        "btn_select_doc": "📥 Вибрати документ",
        "btn_open_contracts": "📊 Відкрити реєстр договорів",
        "btn_open_conclusion": "📄 Відкрити висновок",
        "lbl_doc_waiting_status": "0% - Очікування документа",
        # Investor Studio papildu atslēgas
        "btn_audit_ticker": "⚡ Аналізувати тикер",
        "lbl_ticker_input": "Введіть тикер (напр., AAPL, NVDA):",
    },
    "ru": {
        "tab_offers": "💼 Студия предложений",
        "tab_hp": "💻 HP Эксперт",
        "tab_voice": "🎙️ Голосовая Студия",
        "tab_dropzone": "📁 Зона для документов",
        "tab_radar": "⚙️ Настройки радара",
        "btn_calc": "⚡ Рассчитать и оформить",
        "btn_excel": "📊 Открыть Excel для сметы",
        "btn_copy": "📋 Копировать для Email",
        "lbl_client_req": "1. 📥 Запрос клиента / Требования:",
        "lbl_dist_costs": "2. 🏷️ Себестоимость дистрибьютора:",
        "lbl_result": "⚡ Сгенерированное коммерческое предложение (безопасно для клиента):",
        "lbl_margin": "Маржа (%):",
        "lbl_lang": "Язык:",
        "status_ready": "Готово!",
        "status_calculating": "Gemini рассчитывает и генерирует Excel...",
        # Offer Studio atslēgas, nodrošinātas visām valodām
        "btn_calculate": "⚡ Рассчитать и оформить",
        "btn_open_excel": "📊 Открыть Excel для сметы",
        "btn_copy_email": "📋 Копировать для Email",
        "lbl_language": "Язык:",
        "lbl_client_rfq": "1. 📥 Запрос клиента:",
        "btn_load_file": "📁 Загрузить",
        "lbl_supplier_costs": "2. 🏷️ Себестоимость дистрибьютора:",
        "lbl_commercial_offer": "⚡ Готовое предложение (Client-Safe):",
        # Bank Studio atslēgas
        "btn_load_bank_statement": "📁 Загрузить выписку банка (CSV/XLSX)",
        "btn_analyze_cashflow": "⚡ Анализировать денежный поток",
        "lbl_bank_summary": "Финансовая сводка и показатели:",
        "lbl_incoming_funds": "Поступления:",
        "lbl_outgoing_funds": "Исходящие платежи:",
        "lbl_net_cashflow": "Чистый денежный поток:",
        "lbl_bank_title": "🏦 Citadele Bank Studio (ISO 20022)",
        "btn_load_pdf_invoice": "📄 Загрузить счет PDF",
        "lbl_bank_ready_status": "Готов к работе. Загрузите PDF счет из почты или ПК.",
        "btn_add_partner": "⭐ Добавить партнера в базу",
        "btn_generate_xml": "🚀 Создать Citadele XML",
        # Dokumentu Drop-Zone atslēgas
        "btn_drop_browse": "📄 Выбрать документ (PDF, DOCX, TXT)",
        "btn_summarize_doc": "⚡ Извлечь суть и резюме",
        "lbl_doc_preview": "Предпросмотр документа:",
        "lbl_ai_summary": "AI Анализ и резюме:",
        # Document Digest atslēgas
        "btn_select_doc": "📥 Выбрать документ",
        "btn_open_contracts": "📊 Открыть реестр договоров",
        "btn_open_conclusion": "📄 Открыть заключение",
        "lbl_doc_waiting_status": "0% - Ожидание документа",
        # Investor Studio papildu atslēgas
        "btn_audit_ticker": "⚡ Анализировать тикер",
        "lbl_ticker_input": "Введите тикер (напр., AAPL, NVDA):",
    },
    "it": {
        "tab_offers": "💼 Studio Offerte",
        "tab_hp": "💻 Esperto HP",
        "tab_voice": "🎙️ Studio Vocale",
        "tab_dropzone": "📁 Zona Documenti",
        "tab_radar": "⚙️ Impostazioni Radar",
        "btn_calc": "⚡ Calcola e genera",
        "btn_excel": "📊 Apri foglio Excel",
        "btn_copy": "📋 Copia per email",
        "lbl_client_req": "1. 📥 Richiesta del cliente / Requisiti:",
        "lbl_dist_costs": "2. 🏷️ Costi fornitore / distributore:",
        "lbl_result": "⚡ Preventivo commerciale generato (sicuro per il cliente):",
        "lbl_margin": "Margine (%):",
        "lbl_lang": "Lingua:",
        "status_ready": "Pronto!",
        "status_calculating": "Gemini calcola e genera Excel...",
        # Offer Studio atslēgas, nodrošinātas visām valodām
        "btn_calculate": "⚡ Calcola e genera",
        "btn_open_excel": "📊 Apri foglio Excel",
        "btn_copy_email": "📋 Copia per email",
        "lbl_language": "Lingua:",
        "lbl_client_rfq": "1. 📥 Richiesta del cliente:",
        "btn_load_file": "📁 Carica",
        "lbl_supplier_costs": "2. 🏷️ Costi fornitore:",
        "lbl_commercial_offer": "⚡ Offerta commerciale (Client-Safe):",
        # Bank Studio atslēgas
        "btn_load_bank_statement": "📁 Carica estratto conto (CSV/XLSX)",
        "btn_analyze_cashflow": "⚡ Analizza flusso di cassa",
        "lbl_bank_summary": "Riepilogo finanziario e metriche:",
        "lbl_incoming_funds": "Entrate:",
        "lbl_outgoing_funds": "Uscite:",
        "lbl_net_cashflow": "Flusso di cassa netto:",
        "lbl_bank_title": "🏦 Citadele Bank Studio (ISO 20022)",
        "btn_load_pdf_invoice": "📄 Carica fattura PDF",
        "lbl_bank_ready_status": "Pronto. Carica una fattura PDF da email o PC.",
        "btn_add_partner": "⭐ Aggiungi partner al database",
        "btn_generate_xml": "🚀 Genera Citadele XML",
        # Dokumentu Drop-Zone atslēgas
        "btn_drop_browse": "📄 Seleziona documento (PDF, DOCX, TXT)",
        "btn_summarize_doc": "⚡ Estrai sommario e sintesi",
        "lbl_doc_preview": "Anteprima documento:",
        "lbl_ai_summary": "Analisi AI e sintesi:",
        # Document Digest atslēgas
        "btn_select_doc": "📥 Seleziona documento",
        "btn_open_contracts": "📊 Apri registro contratti",
        "btn_open_conclusion": "📄 Apri relazione",
        "lbl_doc_waiting_status": "0% - In attesa del documento",
        # Investor Studio papildu atslēgas
        "btn_audit_ticker": "⚡ Analizza ticker",
        "lbl_ticker_input": "Inserisci ticker (es., AAPL, NVDA):",
    }
}

# --- Internal cache for current configuration ---
_cached_config = None

def _load_config() -> dict:
    """Helper to load config.json safely. Returns an empty dict if file is not found or corrupted."""
    global _cached_config
    if _cached_config is not None:
        return _cached_config

    if not CONFIG_PATH.exists():
        _cached_config = {}
        return _cached_config
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            _cached_config = json.load(f)
            return _cached_config
    except json.JSONDecodeError:
        # Komandants Kvarks nesveic JSON formāta pārkāpumus.
        print(f"Brīdinājums: '{CONFIG_PATH}' ir bojāts vai nav derīgs JSON. Tiek izmantoti noklusējuma iestatījumi.")
        _cached_config = {}
        return _cached_config
    except Exception as e:
        # Vispārīga kļūda faila lasīšanā.
        print(f"Kļūda lasot '{CONFIG_PATH}': {e}. Tiek izmantoti noklusējuma iestatījumi.")
        _cached_config = {}
        return _cached_config

def _save_config(config_data: dict):
    """Helper to save config.json safely and update the cache."""
    global _cached_config
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
        _cached_config = config_data # Atjauno kešu pēc veiksmīgas saglabāšanas
    except Exception as e:
        # Bruno brīdina par faila rakstīšanas kļūdām.
        print(f"Kļūda saglabājot '{CONFIG_PATH}': {e}")


def get_ui_language() -> str:
    """
    Atgriež aktīvo UI valodas kodu (lv, en, uk, ru, it).
    Nolasot "ui_language" (vai "language") no config.json, obligāti veic normalizāciju.
    Ja config.json ir ierakstīts "RU" vai "UA", tam OBLIGĀTI jāatgriež "ru" vai "uk", nevis jākrīt atpakaļ uz "lv".
    """
    config = _load_config()
    raw_lang_code = config.get("ui_language") or config.get("language")

    # Normalizē valodas kodu, izmantojot palīgfunkciju
    lang_code = _normalize_lang_input(raw_lang_code)

    if lang_code in SUPPORTED_LANGUAGES:
        return lang_code
    else:
        # Ja nav atbalstīto valodu sarakstā pēc normalizācijas, atgriež noklusējumu.
        return DEFAULT_UI_LANGUAGE_CODE

def set_ui_language(lang_code: str):
    """Saglabā ui_language iekš config.json."""
    # Nodrošina, ka iestatāmais valodas kods tiek normalizēts un atbalstīts
    normalized_lang_code = _normalize_lang_input(lang_code)
    if normalized_lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: UI valoda '{lang_code}' (normalizēta '{normalized_lang_code}') netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    config["ui_language"] = normalized_lang_code # Saglabā normalizēto kodu
    _save_config(config)

def get_output_language() -> str:
    """
    Atgriež reālo atskaišu valodas kodu.
    Ja konfigurācijā ir "auto" vai tukša, atgriež get_ui_language().
    Obligāti normalizē ar _normalize_lang_input().
    """
    config = _load_config()
    raw_output_lang_setting = config.get("output_language")

    # Ja sākotnējais iestatījums ir None, tukšs vai "auto", izmanto UI valodu.
    if not raw_output_lang_setting or raw_output_lang_setting.strip().lower() == "auto":
        return get_ui_language()
    
    # Citādi, normalizē konfigurēto valodas kodu (ieskaitot 'ua'->'uk' kartēšanu)
    normalized_lang_code = _normalize_lang_input(raw_output_lang_setting)

    # Ja normalizētais kods tiek atbalstīts, atgriež to. Citādi, atgriežas pie UI valodas.
    if normalized_lang_code in SUPPORTED_LANGUAGES:
        return normalized_lang_code
    else:
        return get_ui_language()

def set_output_language(lang_code: str):
    """
    Saglabā output_language iekš config.json.
    Atbalsta "auto", "lv", "en", "uk", "ru", "it" (pēc normalizācijas).
    """
    normalized_lang_code = _normalize_lang_input(lang_code) # Normalizē pirms pārbaudes un saglabāšanas

    if normalized_lang_code != "auto" and normalized_lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: Atskaišu valoda '{lang_code}' (normalizēta '{normalized_lang_code}') netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    config["output_language"] = normalized_lang_code # Saglabā normalizēto kodu
    _save_config(config)

def get_llm_language_instruction(lang_code: str = None) -> str:
    """
    Atgriež precīzu sistēmas uzvednes teikumu modeļiem par izmantojamo valodu.
    Ja lang_code nav padots, izmanto get_output_language().
    """
    if lang_code is None:
        lang_code = get_output_language()

    # Nodrošina, ka lang_code ir normalizēts pirms uzmeklēšanas
    normalized_lang_code = _normalize_lang_input(lang_code)

    # Atgriežas pie noklusējuma, ja pieprasītā valoda netiek atbalstīta
    lang_info = SUPPORTED_LANGUAGES.get(normalized_lang_code, SUPPORTED_LANGUAGES[DEFAULT_UI_LANGUAGE_CODE])

    prompt_lang = lang_info['prompt_lang']
    return f"IMPORTANT: You must generate the entire analysis, headings, and report strictly in {prompt_lang} language."

def get_text(key: str, default: str = "") -> str:
    """
    Atgriež tulkotu tekstu pašreizējā UI valodā.
    Ja tulkojums nav atrasts, atgriež noklusējuma vērtību vai pašu atslēgu.
    Nodrošina fallback uz 'lv' un tad 'en', ja atslēga nav atrasta aktīvajā valodā.
    """
    lang_code = get_ui_language() # Šī funkcija jau atgriež normalizētu, atbalstītu valodas kodu

    # Mēģina iegūt tulkojumu aktīvajā valodā
    if key in TRANSLATIONS.get(lang_code, {}):
        return TRANSLATIONS[lang_code][key]

    # Ja nav atrasts aktīvajā valodā, mēģina "lv" kā atkāpes variantu
    if key in TRANSLATIONS.get('lv', {}):
        return TRANSLATIONS['lv'][key]

    # Ja nav atrasts "lv", mēģina "en" kā atkāpes variantu
    if key in TRANSLATIONS.get('en', {}):
        return TRANSLATIONS['en'][key]

    # Ja joprojām nav atrasts, atgriež norādīto noklusējumu vai pašu atslēgu (nav tukšums)
    return default or key

def t(key: str, lang: str = None) -> str:
    """
    Atgriež tulkotu tekstu norādītajā valodā vai pašreizējā UI valodā, ja valoda nav norādīta.
    Nodrošina fallback uz 'lv' un tad 'en', ja atslēga nav atrasta norādītajā valodā.
    """
    # Normalizē norādīto 'lang' parametru vai izmanto normalizēto UI valodu
    if lang:
        target_lang = _normalize_lang_input(lang)
    else:
        target_lang = get_ui_language() # Šī funkcija jau atgriež normalizētu, atbalstītu valodas kodu

    # Mēģina iegūt tulkojumu mērķa valodā
    if key in TRANSLATIONS.get(target_lang, {}):
        return TRANSLATIONS[target_lang][key]

    # Ja nav atrasts mērķa valodā, mēģina "lv" kā atkāpes variantu
    if key in TRANSLATIONS.get('lv', {}):
        return TRANSLATIONS['lv'][key]

    # Ja nav atrasts "lv", mēģina "en" kā atkāpes variantu
    if key in TRANSLATIONS.get('en', {}):
        return TRANSLATIONS['en'][key]

    # Ja joprojām nav atrasts, atgriež pašu atslēgu (nav tukšums)
    return key

# Šī funkcija nodrošināta atpakaļejošai savietojamībai, ja kāds izsauc `set_language`.
# Tā tagad iestata UI valodu.
def set_language(lang_code: str):
    """
    Iestata pašreizējo UI valodu sistēmai, atjaunina config.json un iekšējo kešu.
    Šī funkcija nodrošina Komandanta Kvarka pavēles izpildi.
    """
    set_ui_language(lang_code)