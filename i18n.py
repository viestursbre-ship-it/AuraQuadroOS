# -*- coding: utf-8 -*-
import json
import os

# --- Configuration Constants ---
CONFIG_FILE = "config.json"
DEFAULT_LANGUAGE_CODE = 'lv'
# Prioritize "report_language", then "language" for reading configuration
CONFIG_LANGUAGE_KEYS = ["report_language", "language"] 

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

# --- Internal cache for current language, initialized on first access ---
_current_lang_in_memory = None

def _load_config() -> dict:
    """Helper to load config.json safely. Returns an empty dict if file is not found or corrupted."""
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError:
        # Komandants Kvarks nesveic JSON formāta pārkāpumus.
        print(f"Brīdinājums: '{CONFIG_FILE}' ir bojāts vai nav derīgs JSON. Tiek izmantoti noklusējuma iestatījumi.")
        return {}
    except Exception as e:
        # Vispārīga kļūda faila lasīšanā.
        print(f"Kļūda lasot '{CONFIG_FILE}': {e}. Tiek izmantoti noklusējuma iestatījumi.")
        return {}

def _save_config(config_data: dict):
    """Helper to save config.json safely."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        # Bruno brīdina par faila rakstīšanas kļūdām.
        print(f"Kļūda saglabājot '{CONFIG_FILE}': {e}")


def get_current_language() -> str:
    """
    Nolasa pašreizējo valodas izvēli no config.json vai atgriež noklusējuma valodu.
    Valodas izvēle tiek kešota, lai izvairītos no liekas failu lasīšanas.
    """
    global _current_lang_in_memory
    if _current_lang_in_memory is None:
        config = _load_config()
        lang_code = None
        for key in CONFIG_LANGUAGE_KEYS:
            lang_code = config.get(key)
            if lang_code is not None:
                break
        
        if lang_code is None or lang_code not in SUPPORTED_LANGUAGES:
            lang_code = DEFAULT_LANGUAGE_CODE
        _current_lang_in_memory = lang_code
    return _current_lang_in_memory


def set_current_language(lang_code: str):
    """
    Saglabā lietotāja valodas izvēli config.json un atjaunina iekšējo kešu.
    """
    global _current_lang_in_memory
    if lang_code not in SUPPORTED_LANGUAGES:
        # Komandants Kvarks nepieļauj neatbalstītas operācijas.
        print(f"Brīdinājums: Valoda '{lang_code}' netiek atbalstīta. Izvēle netiek saglabāta.")
        return

    config = _load_config()
    # Always write to the primary language key
    config[CONFIG_LANGUAGE_KEYS[0]] = lang_code 
    _save_config(config)
    _current_lang_in_memory = lang_code


def get_llm_language_instruction(lang_code: str = None) -> str:
    """
    Atgriež skaidru sistēmas uzvednes frāzi modelim par izmantojamo valodu.
    Ja valoda nav norādīta, tiek izmantota pašreizējā sistēmas valoda.
    """
    if lang_code is None:
        lang_code = get_current_language()

    # Fallback to default if requested lang_code is not supported
    lang_info = SUPPORTED_LANGUAGES.get(lang_code, SUPPORTED_LANGUAGES[DEFAULT_LANGUAGE_CODE])
    
    prompt_lang = lang_info['prompt_lang']
    return f"IMPORTANT: You must generate the entire analysis, headings, and report strictly in {prompt_lang} language."


def get_text(key: str, default: str = "") -> str:
    """
    Atgriež tulkotu tekstu pašreizējā sistēmas valodā.
    Ja tulkojums nav atrasts, atgriež noklusējuma vērtību vai pašu atslēgu.
    """
    lang_code = get_current_language()
    return TRANSLATIONS.get(lang_code, {}).get(key, default or key)


def set_language(lang_code: str):
    """
    Iestata pašreizējo valodu sistēmai, atjaunina config.json un iekšējo kešu.
    Šī funkcija nodrošina Komandanta Kvarka pavēles izpildi.
    """
    set_current_language(lang_code)