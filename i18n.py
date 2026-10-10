# -*- coding: utf-8 -*-
TRANSLATIONS = {
    "LV": {
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
    "EN": {
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
    }
}

CURRENT_LANG = "LV"

def get_text(key: str, default: str = "") -> str:
    return TRANSLATIONS.get(CURRENT_LANG, {}).get(key, default or key)

def set_language(lang: str):
    global CURRENT_LANG
    if lang in TRANSLATIONS:
        CURRENT_LANG = lang
