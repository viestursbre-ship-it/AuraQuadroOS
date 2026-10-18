# -*- coding: utf-8 -*-
import json
from pathlib import Path # Import Path for robust path handling

# --- Configuration Constants ---
CONFIG_PATH = Path(__file__).resolve().parent / "config.json" # Drošs ceļš uz config.json
DEFAULT_UI_LANGUAGE_CODE = 'lv' # Noklusējuma valoda lietotāja saskarnei
DEFAULT_OUTPUT_LANGUAGE_CODE = 'auto' # Noklusējuma valoda atskaitēm/ģenerētajiem failiem (vai 'lv', 'en', 'auto' u.c.)

# --- Supported Languages Definition ---
SUPPORTED_LANGUAGES = {
    'lv': {'name': 'Latviešu', 'prompt_lang': 'Latvian'},
    'en': {'name': 'English', 'prompt_lang': 'English'},
    'uk': {'name': 'Українська', 'prompt_lang': 'Ukrainian'},
    'ru': {'name': 'Русский', 'prompt_lang': 'Russian'},
    'it': {'name': 'Italiano', 'prompt_lang': 'Italian'}
}

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
        # New translations for 'lv'
        "btn_calculate": "⚡ Aprēķināt & Noformēt",
        "btn_open_excel": "📊 Atvērt Excel Tāmēšanai",
        "btn_copy_email": "📋 Kopēt E-pastam",
        "lbl_language": "Valoda:",
        "lbl_client_rfq": "1. 📥 Klienta Pieprasījums:",
        "btn_load_file": "📁 Ielādēt",
        "lbl_supplier_costs": "2. 🏷️ Distributora Pašizmaksa / Cenas:",
        "lbl_commercial_offer": "⚡ Gatavais Piedāvājums Klientam (Drošs):",
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
        # New translations for 'en'
        "btn_calculate": "⚡ Calculate & Generate",
        "btn_open_excel": "📊 Open Excel Sheet",
        "btn_copy_email": "📋 Copy for Email",
        "lbl_language": "Language:",
        "lbl_client_rfq": "1. 📥 Client RFQ / Requirements:",
        "btn_load_file": "📁 Load",
        "lbl_supplier_costs": "2. 🏷️ Supplier / Disti Costs:",
        "lbl_commercial_offer": "⚡ Commercial Offer (Client-Safe):",
    },
    "uk": {
        # New translations for 'uk'
        "btn_calculate": "⚡ Розрахувати та оформити",
        "btn_open_excel": "📊 Відкрити Excel для кошторису",
        "btn_copy_email": "📋 Копіювати для Email",
        "lbl_margin": "Маржа (%):",
        "lbl_language": "Мова:",
        "lbl_client_rfq": "1. 📥 Запит клієнта:",
        "btn_load_file": "📁 Завантажити",
        "lbl_supplier_costs": "2. 🏷️ Собівартість дистриб'ютора:",
        "lbl_commercial_offer": "⚡ Готова пропозиція (Client-Safe):",
    },
    "ru": {
        # New translations for 'ru'
        "btn_calculate": "⚡ Рассчитать и оформить",
        "btn_open_excel": "📊 Открыть Excel для сметы",
        "btn_copy_email": "📋 Копировать для Email",
        "lbl_margin": "Маржа (%):",
        "lbl_language": "Язык:",
        "lbl_client_rfq": "1. 📥 Запрос клиента:",
        "btn_load_file": "📁 Загрузить",
        "lbl_supplier_costs": "2. 🏷️ Себестоимость дистрибьютора:",
        "lbl_commercial_offer": "⚡ Готовое предложение (Client-Safe):",
    },
    "it": {
        # New translations for 'it'
        "btn_calculate": "⚡ Calcola e genera",
        "btn_open_excel": "📊 Apri foglio Excel",
        "btn_copy_email": "📋 Copia per email",
        "lbl_margin": "Margine (%):",
        "lbl_language": "Lingua:",
        "lbl_client_rfq": "1. 📥 Richiesta del cliente:",
        "btn_load_file": "📁 Carica",
        "lbl_supplier_costs": "2. 🏷️ Costi fornitore:",
        "lbl_commercial_offer": "⚡ Offerta commerciale (Client-Safe):",
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
    Nolasot "ui_language" (vai "language") no config.json, obligāti veic .strip().lower().
    Ja config.json ir ierakstīts "RU", tam OBLIGĀTI jāatgriež "ru", nevis jākrīt atpakaļ uz "lv".
    """
    config = _load_config()
    # Nolasot "ui_language" (vai "language") no config.json, obligāti veic .strip().lower().
    # Fallback to "language" key for backward compatibility if it exists, though "ui_language" is primary.
    raw_lang_code = config.get("ui_language") or config.get("language")

    # Normalize the language code immediately
    if isinstance(raw_lang_code, str):
        lang_code = raw_lang_code.strip().lower()
    else:
        lang_code = DEFAULT_UI_LANGUAGE_CODE # If it's not a string (e.g., None), use default

    # Ja iegūtā vērtība ir "ru", "uk", "en", "it", "lv", atgriez to.
    # Ja config.json ir ierakstīts "RU", tam OBLIGĀTI jāatgriež "ru", nevis jākrīt atpakaļ uz "lv".
    if lang_code in SUPPORTED_LANGUAGES:
        return lang_code
    else:
        # Ja nav atbalstīto valodu sarakstā pēc normalizācijas, atgriež noklusējumu.
        return DEFAULT_UI_LANGUAGE_CODE

def set_ui_language(lang_code: str):
    """Saglabā ui_language iekš config.json."""
    # Ensure the lang_code being set is normalized and supported
    normalized_lang_code = lang_code.strip().lower()
    if normalized_lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: UI valoda '{lang_code}' (normalizēta '{normalized_lang_code}') netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    config["ui_language"] = normalized_lang_code # Save the normalized code
    _save_config(config)

def get_output_language() -> str:
    """
    Atgriež reālo atskaišu valodas kodu.
    Ja konfigurācijā ir "auto" vai tukša, atgriež get_ui_language().
    Obligāti normalizē ar .strip().lower().
    """
    config = _load_config()
    raw_output_lang_setting = config.get("output_language")

    # Obligāti normalizē ar .strip().lower().
    if isinstance(raw_output_lang_setting, str):
        output_lang_setting = raw_output_lang_setting.strip().lower()
    else:
        output_lang_setting = DEFAULT_OUTPUT_LANGUAGE_CODE # If not a string, use default (which is 'auto')

    # Ja tā ir "auto" vai tukša, atgriez get_ui_language().
    if not output_lang_setting or output_lang_setting == "auto":
        return get_ui_language()
    elif output_lang_setting in SUPPORTED_LANGUAGES:
        return output_lang_setting
    else:
        # Atgriež UI valodu, ja konfigurētā atskaišu valoda ir nederīga/nav atbalstīta
        return get_ui_language()

def set_output_language(lang_code: str):
    """
    Saglabā output_language iekš config.json.
    Atbalsta "auto", "lv", "en", "uk", "ru", "it".
    """
    normalized_lang_code = lang_code.strip().lower() # Normalize before checking and saving

    if normalized_lang_code != "auto" and normalized_lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: Atskaišu valoda '{lang_code}' (normalizēta '{normalized_lang_code}') netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    config["output_language"] = normalized_lang_code # Save the normalized code
    _save_config(config)

def get_llm_language_instruction(lang_code: str = None) -> str:
    """
    Atgriež precīzu sistēmas uzvednes teikumu modeļiem par izmantojamo valodu.
    Ja lang_code nav padots, izmanto get_output_language().
    """
    if lang_code is None:
        lang_code = get_output_language()

    # Ensure lang_code is normalized before lookup
    normalized_lang_code = lang_code.strip().lower() if lang_code else DEFAULT_UI_LANGUAGE_CODE

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
    lang_code = get_ui_language() # This already returns a normalized, supported language code

    # Try to get translation in the active language
    if key in TRANSLATIONS.get(lang_code, {}):
        return TRANSLATIONS[lang_code][key]

    # If not found in active language, try 'lv' as fallback
    if key in TRANSLATIONS.get('lv', {}):
        return TRANSLATIONS['lv'][key]

    # If not found in 'lv', try 'en' as fallback
    if key in TRANSLATIONS.get('en', {}):
        return TRANSLATIONS['en'][key]

    # If still not found, return the provided default or the key itself
    return default or key

# Šī funkcija nodrošināta atpakaļejošai savietojamībai, ja kāds izsauc `set_language`.
# Tā tagad iestata UI valodu.
def set_language(lang_code: str):
    """
    Iestata pašreizējo UI valodu sistēmai, atjaunina config.json un iekšējo kešu.
    Šī funkcija nodrošina Komandanta Kvarka pavēles izpildi.
    """
    set_ui_language(lang_code)