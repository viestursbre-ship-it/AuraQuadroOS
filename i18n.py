# -*- coding: utf-8 -*-
import json
import os

# --- Configuration Constants ---
CONFIG_FILE = "config.json"
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
    },
    # Empty stubs for new languages, allowing get_text to safely access them
    "uk": {},
    "ru": {},
    "it": {}
}

# --- Internal cache for current configuration ---
_cached_config = None 

def _load_config() -> dict:
    """Helper to load config.json safely. Returns an empty dict if file is not found or corrupted."""
    global _cached_config
    if _cached_config is not None:
        return _cached_config

    if not os.path.exists(CONFIG_FILE):
        _cached_config = {}
        return _cached_config
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            _cached_config = json.load(f)
            return _cached_config
    except json.JSONDecodeError:
        # Komandants Kvarks nesveic JSON formāta pārkāpumus.
        print(f"Brīdinājums: '{CONFIG_FILE}' ir bojāts vai nav derīgs JSON. Tiek izmantoti noklusējuma iestatījumi.")
        _cached_config = {}
        return _cached_config
    except Exception as e:
        # Vispārīga kļūda faila lasīšanā.
        print(f"Kļūda lasot '{CONFIG_FILE}': {e}. Tiek izmantoti noklusējuma iestatījumi.")
        _cached_config = {}
        return _cached_config

def _save_config(config_data: dict):
    """Helper to save config.json safely and update the cache."""
    global _cached_config
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
        _cached_config = config_data # Atjauno kešu pēc veiksmīgas saglabāšanas
    except Exception as e:
        # Bruno brīdina par faila rakstīšanas kļūdām.
        print(f"Kļūda saglabājot '{CONFIG_FILE}': {e}")


def get_ui_language() -> str:
    """Atgriež aktīvo UI valodas kodu (lv, en, uk, ru, it)."""
    config = _load_config()
    lang_code = config.get("ui_language", DEFAULT_UI_LANGUAGE_CODE)
    if lang_code not in SUPPORTED_LANGUAGES:
        lang_code = DEFAULT_UI_LANGUAGE_CODE
    return lang_code

def set_ui_language(lang_code: str):
    """Saglabā ui_language iekš config.json."""
    if lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: UI valoda '{lang_code}' netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    config["ui_language"] = lang_code
    _save_config(config)

def get_output_language() -> str:
    """
    Atgriež reālo atskaišu valodas kodu.
    Ja konfigurācijā ir "auto", atgriež get_ui_language(), citādi konkrēto norādīto kodu.
    """
    config = _load_config()
    output_lang_setting = config.get("output_language", DEFAULT_OUTPUT_LANGUAGE_CODE)

    if output_lang_setting == "auto":
        return get_ui_language()
    elif output_lang_setting in SUPPORTED_LANGUAGES:
        return output_lang_setting
    else:
        # Atgriež UI valodu, ja konfigurētā atskaišu valoda ir nederīga
        return get_ui_language()

def set_output_language(lang_code: str):
    """
    Saglabā output_language iekš config.json.
    Atbalsta "auto", "lv", "en", "uk", "ru", "it".
    """
    if lang_code != "auto" and lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: Atskaišu valoda '{lang_code}' netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    config["output_language"] = lang_code
    _save_config(config)

def get_llm_language_instruction(lang_code: str = None) -> str:
    """
    Atgriež precīzu sistēmas uzvednes teikumu modeļiem par izmantojamo valodu.
    Ja lang_code nav padots, izmanto get_output_language().
    """
    if lang_code is None:
        lang_code = get_output_language()

    # Atgriežas pie noklusējuma, ja pieprasītā valoda netiek atbalstīta
    lang_info = SUPPORTED_LANGUAGES.get(lang_code, SUPPORTED_LANGUAGES[DEFAULT_UI_LANGUAGE_CODE])
    
    prompt_lang = lang_info['prompt_lang']
    return f"IMPORTANT: You must generate the entire analysis, headings, and report strictly in {prompt_lang} language."

def get_text(key: str, default: str = "") -> str:
    """
    Atgriež tulkotu tekstu pašreizējā UI valodā.
    Ja tulkojums nav atrasts, atgriež noklusējuma vērtību vai pašu atslēgu.
    """
    lang_code = get_ui_language()
    return TRANSLATIONS.get(lang_code, {}).get(key, default or key)

# Šī funkcija nodrošināta atpakaļejošai savietojamībai, ja kāds izsauc `set_language`.
# Tā tagad iestata UI valodu.
def set_language(lang_code: str):
    """
    Iestata pašreizējo UI valodu sistēmai, atjaunina config.json un iekšējo kešu.
    Šī funkcija nodrošina Komandanta Kvarka pavēles izpildi.
    """
    set_ui_language(lang_code)