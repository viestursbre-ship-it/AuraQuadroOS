# -*- coding: utf-8 -*-
import os
import json
import customtkinter as ctk
import i18n # JAUNUMS: Importējam i18n bibliotēku

from hp_expert import HPExpertFrame
from voice_studio import VoiceStudioFrame
from doc_digest import DocDigestFrame
from offer_studio import OfferStudioFrame
from warranty_hub import WarrantyHubFrame
from bank_studio import BankStudioFrame
from investor_studio import InvestorStudioFrame

# --- Valodu kartēšana / Language Mappings ---
_UI_LANG_DISPLAY_MAP = {
    "🇱🇻 LV": "LV", # Display text -> UI_TEXTS key (uppercase)
    "🇬🇧 EN": "EN",
    "🇺🇦 UA": "UA",
    "🌐 RU": "RU",
    "🇮🇹 IT": "IT"
}
_UI_LANG_KEY_MAP = {v: k for k, v in _UI_LANG_DISPLAY_MAP.items()} # UI_TEXTS key (uppercase) -> Display text

_OUTPUT_LANG_DISPLAY_MAP = {
    "Auto (tāda pati kā saskarnei)": "auto",
    "🇱🇻 Latviešu (LV)": "lv",
    "🇬🇧 English (EN)": "en",
    "🇺🇦 Українська (UA)": "uk", # i18n standard 'uk'
    "🌐 Русский (RU)": "ru",
    "🇮🇹 Italiano (IT)": "it"
}
_OUTPUT_LANG_CODE_MAP = {v: k for k, v in _OUTPUT_LANG_DISPLAY_MAP.items()}

# JAUNUMS: Radara režīma valodu un koda kartējumi
_RADAR_MODE_DISPLAY_MAP = {
    "Offline / Manual": "manual",
    "Gmail (IMAP)": "imap",
    "Outlook (Desktop)": "outlook"
}
_RADAR_MODE_CODE_MAP = {v: k for k, v in _RADAR_MODE_DISPLAY_MAP.items()}


UI_TEXTS = {
    "LV": {
        "title": "⚡ AQ-OS Cockpit — Vadības Centrs",
        "tab_offers": "💼 Piedāvājumi",
        "tab_hp": "🧠 HP Eksperts",
        "tab_voice": "🎙️ Balss Studija",
        "tab_docs": "📄 Dokumentu Drop-Zone",
        "tab_settings": "⚙️ AQ Iestatījumi",
        "tab_warranty": "🛡️ Garantija",
        "tab_bank": "🏦 Bank Studio",
        "tab_invest": "📈 Investor Studio",
        "lbl_bank_header": "🏦 Bankas Rekvizīti (Citadele / ISO 20022):",
        "lbl_my_iban": "Maksātāja IBAN konts:",
        "lbl_my_name": "Maksātāja Uzņēmums / Vārds:",
        "lbl_time": "🕒 Vakara Komandanta Atskaites laiks (HH:MM):",
        "lbl_kw": "📡 Radara Atslēgvārdi (atdalīti ar komatiem):",
        "lbl_desc": "E-pasti, kuru temats vai teksts satur šos vārdus, automātiski pārtop par uzdevumiem.",
        "btn_save": "💾 Saglabāt iestatījumus",
        "saved_msg": "✅ Iestatījumi saglabāti! (Atskaite: {time})",
        "lbl_output_lang_header": "📄 Atskaišu un dokumentu ģenerēšanas valoda:",
        "lbl_radar_mode": "📡 Radara režīms / Radar Mode:",
        "lbl_gemini_api_key": "🔑 Gemini API Key:",
        "lbl_finnhub_api_key": "📈 Finnhub API Key:",
        "lbl_alpha_api_key": "📊 Alpha Vantage API Key:", # JAUNUMS
        "lbl_fmp_api_key": "🏛️ FMP API Key:", # JAUNUMS
        # JAUNUMS: IMAP konfigurācijas sadaļa
        "lbl_imap_header": "📧 E-pasta IMAP Pieslēgums (Gmail / M365):",
        "lbl_imap_server": "Serveris:",
        "lbl_imap_port": "Ports:",
        "lbl_imap_user": "Lietotājs / E-pasts:",
        "lbl_imap_pass": "Parole (Gmail: App Password):",
    },
    "EN": {
        "title": "⚡ AQ-OS Cockpit — Control Center",
        "tab_offers": "💼 Quote Studio",
        "tab_hp": "🧠 HP Expert",
        "tab_voice": "🎙️ Voice Studio",
        "tab_docs": "📄 Document Drop-Zone",
        "tab_settings": "⚙️ AQ Settings",
        "tab_warranty": "🛡️ Warranty",
        "tab_bank": "🏦 Bank Studio",
        "tab_invest": "📈 Investor Studio",
        "lbl_bank_header": "🏦 Banking Details (Citadele / ISO 20022):",
        "lbl_my_iban": "Payer IBAN Account:",
        "lbl_my_name": "Payer Company / Name:",
        "lbl_time": "🕒 Evening Commandant Briefing Time (HH:MM):",
        "lbl_kw": "📡 Radar Keywords (comma separated):",
        "lbl_desc": "Emails containing these words in subject or body automatically become tasks.",
        "btn_save": "💾 Save Settings",
        "saved_msg": "✅ Settings saved! (Briefing: {time})",
        "lbl_output_lang_header": "📄 Report and Document Generation Language:",
        "lbl_radar_mode": "📡 Radar Mode:",
        "lbl_gemini_api_key": "🔑 Gemini API Key:",
        "lbl_finnhub_api_key": "📈 Finnhub API Key:",
        "lbl_alpha_api_key": "📊 Alpha Vantage API Key:", # JAUNUMS
        "lbl_fmp_api_key": "🏛️ FMP API Key:", # JAUNUMS
        # JAUNUMS: IMAP konfigurācijas sadaļa
        "lbl_imap_header": "📧 Email IMAP Connection (Gmail / M365):",
        "lbl_imap_server": "Server:",
        "lbl_imap_port": "Port:",
        "lbl_imap_user": "User / Email:",
        "lbl_imap_pass": "Password (Gmail: App Password):",
    },
    "UA": { # Placeholder translations (UI_TEXTS expects "UA")
        "title": "⚡ AQ-OS Cockpit — Центр Управління",
        "tab_offers": "💼 Пропозиції",
        "tab_hp": "🧠 HP Експерт",
        "tab_voice": "🎙️ Голосова Студія",
        "tab_docs": "📄 Зона Документів",
        "tab_settings": "⚙️ Налаштування AQ",
        "tab_warranty": "🛡️ Гарантія",
        "tab_bank": "🏦 Банк Студія",
        "tab_invest": "📈 Студія Інвестора",
        "lbl_bank_header": "🏦 Банківські Реквізити (Citadele / ISO 20022):",
        "lbl_my_iban": "IBAN рахунок платника:",
        "lbl_my_name": "Компанія / Ім'я платника:",
        "lbl_time": "🕒 Час вечірнього брифінгу Коменданта (HH:MM):",
        "lbl_kw": "📡 Ключові слова радара (через кому):",
        "lbl_desc": "Електронні листи, що містять ці слова у темі або тексте, автоматично стають завданнями.",
        "btn_save": "💾 Зберегти налаштування",
        "saved_msg": "✅ Налаштування збережено! (Брифінг: {time})",
        "lbl_output_lang_header": "📄 Мова генерування звітів та документів:",
        "lbl_radar_mode": "📡 Режим Радару:",
        "lbl_gemini_api_key": "🔑 Gemini API Key:",
        "lbl_finnhub_api_key": "📈 Finnhub API Key:",
        "lbl_alpha_api_key": "📊 Ключ API Alpha Vantage:", # JAUNUMS
        "lbl_fmp_api_key": "🏛️ Ключ API FMP:", # JAUNUMS
        # JAUNUMS: IMAP konfigurācijas sadaļa
        "lbl_imap_header": "📧 Підключення IMAP електронної пошти (Gmail / M365):",
        "lbl_imap_server": "Сервер:",
        "lbl_imap_port": "Порт:",
        "lbl_imap_user": "Користувач / Електронна пошта:",
        "lbl_imap_pass": "Пароль (Gmail: Пароль додатка):",
    },
    "RU": { # Placeholder translations
        "title": "⚡ AQ-OS Cockpit — Центр Управления",
        "tab_offers": "💼 Предложения",
        "tab_hp": "🧠 HP Эксперт",
        "tab_voice": "🎙️ Голосовая Студия",
        "tab_docs": "📄 Зона Документов",
        "tab_settings": "⚙️ Настройки AQ",
        "tab_warranty": "🛡️ Гарантия",
        "tab_bank": "🏦 Банк Студия",
        "tab_invest": "📈 Студия Инвестора",
        "lbl_bank_header": "🏦 Банковские Реквизиты (Citadele / ISO 20022):",
        "lbl_my_iban": "IBAN счет плательщика:",
        "lbl_my_name": "Компания / Имя плательщика:",
        "lbl_time": "🕒 Время вечернего отчета Коменданта (ЧЧ:ММ):",
        "lbl_kw": "📡 Ключевые слова радара (через запятую):",
        "lbl_desc": "Электронные письма, содержащие эти слова в теме или тексте, автоматически становятся задачами.",
        "btn_save": "💾 Сохранить настройки",
        "saved_msg": "✅ Настройки сохранены! (Отчет: {time})",
        "lbl_output_lang_header": "📄 Язык генерации отчетов и документов:",
        "lbl_radar_mode": "📡 Режим Радара:",
        "lbl_gemini_api_key": "🔑 Gemini API Key:",
        "lbl_finnhub_api_key": "📈 Finnhub API Key:",
        "lbl_alpha_api_key": "📊 Ключ API Alpha Vantage:", # JAUNUMS
        "lbl_fmp_api_key": "🏛️ Ключ API FMP:", # JAUNUMS
        # JAUNUMS: IMAP konfigurācijas sadaļa
        "lbl_imap_header": "📧 Подключение IMAP электронной почты (Gmail / M365):",
        "lbl_imap_server": "Сервер:",
        "lbl_imap_port": "Порт:",
        "lbl_imap_user": "Пользователь / Электронная почта:",
        "lbl_imap_pass": "Пароль (Gmail: Пароль приложения):",
    },
    "IT": { # Placeholder translations
        "title": "⚡ AQ-OS Cockpit — Centro di Controllo",
        "tab_offers": "💼 Offerte",
        "tab_hp": "🧠 Esperto HP",
        "tab_voice": "🎙️ Studio Vocale",
        "tab_docs": "📄 Zona Documenti",
        "tab_settings": "⚙️ Impostazioni AQ",
        "tab_warranty": "🛡️ Garanzia",
        "tab_bank": "🏦 Studio Bancario",
        "tab_invest": "📈 Studio Investitore",
        "lbl_bank_header": "🏦 Dettagli Bancari (Citadele / ISO 20022):",
        "lbl_my_iban": "Conto IBAN del pagatore:",
        "lbl_my_name": "Azienda / Nome del pagatore:",
        "lbl_time": "🕒 Orario del briefing serale del Comandante (HH:MM):",
        "lbl_kw": "📡 Parole chiave radar (separate da virgola):",
        "lbl_desc": "Le email contenenti queste parole nell'oggetto o nel corpo diventano automaticamente attività.",
        "btn_save": "💾 Salva impostazioni",
        "saved_msg": "✅ Impostazioni salvate! (Briefing: {time})",
        "lbl_output_lang_header": "📄 Lingua di generazione rapporti e documenti:",
        "lbl_radar_mode": "📡 Modalità Radar:",
        "lbl_gemini_api_key": "🔑 Chiave API Gemini:",
        "lbl_finnhub_api_key": "📈 Chiave API Finnhub:",
        "lbl_alpha_api_key": "📊 Chiave API Alpha Vantage:", # JAUNUMS
        "lbl_fmp_api_key": "🏛️ Chiave API FMP:", # JAUNUMS
        # JAUNUMS: IMAP konfigurācijas sadaļa
        "lbl_imap_header": "📧 Connessione IMAP Email (Gmail / M365):",
        "lbl_imap_server": "Server:",
        "lbl_imap_port": "Porta:",
        "lbl_imap_user": "Utente / Email:",
        "lbl_imap_pass": "Password (Gmail: Password App):",
    }
}

def get_gemini_api_key():
    # 1. Vispirms skatāmies lokālajā config.json
    if os.path.exists("config.json"):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                key = cfg.get("gemini_api_key") or cfg.get("aq_ai_api_key")
                if key:
                    return key
        except Exception:
            pass
            
    # 2. Tikai tad, ja failā nav, skatāmies sistēmas mainīgajos
    return os.environ.get("GEMINI_API_KEY") or os.environ.get("AQ_AI_API_KEY", "")

# JAUNUMS: Funkcija sākotnējās UI valodas ielādei no config.json
# Atgriež valodas kodu, kas atbilst UI_TEXTS atslēgām (piem., "LV", "EN", "UA")
def _load_initial_ui_language_from_config():
    if os.path.exists("config.json"):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                stored_lang_code = cfg.get("ui_language", "lv").lower() # Noklusējam uz 'lv' un pārveidojam uz mazo burtu
                if stored_lang_code == "uk": # I18n 'uk' -> UI_TEXTS 'UA'
                    return "UA"
                return stored_lang_code.upper() # "lv" -> "LV", "en" -> "EN"
        except Exception:
            pass
    return "LV" # Noklusējuma valodas kods UI_TEXTS meklēšanai

def get_initial_output_language():
    if os.path.exists("config.json"):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                return cfg.get("output_language", "auto")
        except Exception:
            pass
    return "auto"

def _load_config_data():
    """Helper function to load config.json data safely."""
    if os.path.exists("config.json"):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

class CockpitWindow(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        
        # Sākotnējā valoda no config.json, kas paredzēta UI_TEXTS atslēgām
        self.current_lang = _load_initial_ui_language_from_config()
        
        self.title(UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])["title"])
        self.geometry("1020x740")
        self.minsize(850, 600)

        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        header_frame = ctk.CTkFrame(self, fg_color="#0f172a", height=60, corner_radius=0)
        header_frame.pack(fill="x", side="top")

        ctk.CTkLabel(
            header_frame, 
            text="⚡ AQ-OS MISSION CONTROL", 
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#38bdf8"
        ).pack(side="left", padx=20, pady=15)

        ctk.CTkLabel(
            header_frame, 
            text="v2.5 Modular Enterprise Hub", 
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        ).pack(side="left", padx=5, pady=15)

        # Globālais valodas pārslēgs labajā pusē (UI) - aizvietots ar CTkOptionMenu
        self.ui_lang_optionmenu = ctk.CTkOptionMenu(
            header_frame,
            values=list(_UI_LANG_DISPLAY_MAP.keys()),
            width=120,
            height=28,
            command=self.on_ui_language_change
        )
        # Iestata sākuma vērtību valodas izvēlnē, balstoties uz config.json
        initial_ui_display_text = _UI_LANG_KEY_MAP.get(self.current_lang, "🇱🇻 LV")
        self.ui_lang_optionmenu.set(initial_ui_display_text)
        self.ui_lang_optionmenu.pack(side="right", padx=20, pady=15)

        self.tabview = ctk.CTkTabview(self, fg_color="#1e293b", segmented_button_fg_color="#0f172a")
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)

        # Sagatavojam cilnes ar sākotnējiem nosaukumiem
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        self.tab_offers = self.tabview.add(t["tab_offers"])
        self.tab_hp = self.tabview.add(t["tab_hp"])
        self.tab_bank = self.tabview.add(t.get("tab_bank", "🏦 Bank Studio"))
        self.tab_voice = self.tabview.add(t["tab_voice"])
        self.tab_docs = self.tabview.add(t["tab_docs"])
        self.tab_warranty = self.tabview.add(t["tab_warranty"])
        self.tab_settings = self.tabview.add(t["tab_settings"])
        # Investor Studio ir pievienots pēdējais, atbilstoši esošajam kodam
        self.tab_invest = self.tabview.add(t.get("tab_invest", "📈 Investor Studio"))

        # Moduļu izvietošana
        self.offer_module = OfferStudioFrame(self.tab_offers, get_gemini_key_fn=get_gemini_api_key)
        self.offer_module.pack(fill="both", expand=True)

        self.hp_module = HPExpertFrame(self.tab_hp, get_gemini_key_fn=get_gemini_api_key)
        self.hp_module.pack(fill="both", expand=True)

        self.bank_module = BankStudioFrame(self.tab_bank, get_gemini_key_fn=get_gemini_api_key)
        self.bank_module.pack(fill="both", expand=True)

        self.voice_module = VoiceStudioFrame(self.tab_voice, get_gemini_key_fn=get_gemini_api_key)
        self.voice_module.pack(fill="both", expand=True)

        self.docs_module = DocDigestFrame(self.tab_docs, get_gemini_key_fn=get_gemini_api_key)
        self.docs_module.pack(fill="both", expand=True)

        self.invest_module = InvestorStudioFrame(self.tab_invest, get_gemini_key_fn=get_gemini_api_key)
        self.invest_module.pack(fill="both", expand=True)

        self.warranty_module = WarrantyHubFrame(
            self.tab_warranty, 
            get_gemini_key_fn=get_gemini_api_key,
            on_send_to_offer_fn=self.transfer_warranty_to_offer
        )
        self.warranty_module.pack(fill="both", expand=True)

        # JAUNUMS: Izveido kartēšanu no tab rāmja uz moduļa instanci
        self.frame_to_module_map = {
            self.tab_offers: self.offer_module,
            self.tab_hp: self.hp_module,
            self.tab_bank: self.bank_module,
            self.tab_voice: self.voice_module,
            self.tab_docs: self.docs_module,
            self.tab_warranty: self.warranty_module,
            self.tab_invest: self.invest_module,
            self.tab_settings: self # 'self' (CockpitWindow) apstrādā iestatījumu cilni
        }

        self.setup_settings_tab()
        # Izsauc valodu maiņas funkciju, lai moduļi un UI elementi tiktu atsvaidzināti ar ielādēto valodu
        # Tas pārliecinās, ka viss startējas pareizā valodā, nevis noklusējuma latviešu
        self.on_ui_language_change(initial_ui_display_text)


    def on_ui_language_change(self, display_text):
        """
        Pārslēdz UI valodu, atjaunina ciļņu pogas un sinhronizē moduļus.
        """
        
        # 1. Tīra valodas koda iegūšana un tūlītēja atjaunošana
        # Iegūstam UI_TEXTS atslēgu (piem., "LV", "UA") no display_text, kas tiek izmantota UI tekstu meklēšanai.
        selected_ui_texts_key = _UI_LANG_DISPLAY_MAP.get(display_text, "LV")
        
        # Atjaunojam instances mainīgo, kas tieši atbilst UI_TEXTS atslēgām.
        self.current_lang = selected_ui_texts_key 
        
        # Iegūstam i18n standarta valodas kodu (piem., "lv", "uk"), kas jāglabā config.json un jānodod moduļiem.
        config_lang_code = selected_ui_texts_key.lower() # "LV" -> "lv", "UA" -> "ua"
        if config_lang_code == "ua":
            config_lang_code = "uk" # Pārvēršam "ua" uz "uk" atbilstoši i18n standartam

        # JAUNUMS: Izsauc i18n.set_ui_language ar jauniegūto kodu
        i18n.set_ui_language(config_lang_code)

        # JAUNUMS: Saglabā izvēli config.json (i18n standarta valodas kods)
        cfg = _load_config_data()
        cfg["ui_language"] = config_lang_code # Saglabājam i18n standarta kodu
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        
        # Atjaunina galvenā loga nosaukumu un ciļņu pogu tekstus
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        self.title(t["title"])

        tab_key_order = ["tab_offers", "tab_hp", "tab_bank", "tab_voice", "tab_docs", "tab_warranty", "tab_settings", "tab_invest"]
        
        button_list = list(self.tabview._segmented_button._buttons_dict.values())
        for idx, key in enumerate(tab_key_order):
            if idx < len(button_list):
                button_list[idx].configure(text=t.get(key, key))

        # 2. Moduļu tūlītēja atsvaidzināšana ar JAUNO valodu
        all_modules = [
            self.offer_module, self.hp_module, self.bank_module,
            self.voice_module, self.docs_module, self.warranty_module,
            self.invest_module
        ]

        for module in all_modules:
            if module:
                if hasattr(module, "set_ui_language"):
                    try: # JAUNUMS: Pievienots try-except bloks
                        module.set_ui_language(config_lang_code) 
                    except Exception as e:
                        print(f"Kļūda set_ui_language modulī {type(module).__name__}: {e}")
                if hasattr(module, "refresh_ui"):
                    try:
                        module.refresh_ui()
                    except Exception as e:
                        print(f"Kļūda refresh_ui modulī {type(module).__name__}: {e}")

        # 3. Paziņo Radaram (master logam), ja tam ir valodas metode
        if self.master and hasattr(self.master, "set_ui_language"):
            try:
                # Nododam i18n standarta valodas kodu master logam
                self.master.set_ui_language(config_lang_code)
            except Exception:
                pass

        # 4. Atjaunina Iestatījumu cilnes uzrakstus
        self.update_settings_labels()

    def update_settings_labels(self):
        """Atjaunina visus teksta uzrakstus Iestatījumu cilnē."""
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        if hasattr(self, "lbl_bank_sec"): # Pārbauda, vai logrīki ir inicializēti
            self.lbl_bank_sec.configure(text=t["lbl_bank_header"])
            self.lbl_iban.configure(text=t["lbl_my_iban"])
            self.lbl_name.configure(text=t["lbl_my_name"])
            self.lbl_time_header.configure(text=t["lbl_time"])
            self.lbl_kw_header.configure(text=t["lbl_kw"])
            self.lbl_desc_text.configure(text=t["lbl_desc"])
            self.btn_save_settings.configure(text=t["btn_save"])
            self.lbl_output_lang_header.configure(text=t["lbl_output_lang_header"]) 
            
            if hasattr(self, "lbl_radar_mode_header"):
                self.lbl_radar_mode_header.configure(text=t["lbl_radar_mode"])
            if hasattr(self, "lbl_gemini_api_key"):
                self.lbl_gemini_api_key.configure(text=t["lbl_gemini_api_key"])
            if hasattr(self, "lbl_finnhub_api_key"):
                self.lbl_finnhub_api_key.configure(text=t["lbl_finnhub_api_key"])
            # JAUNUMS: Atjaunina jauno API atslēgu uzrakstus
            if hasattr(self, "lbl_alpha_api_key"):
                self.lbl_alpha_api_key.configure(text=t["lbl_alpha_api_key"])
            if hasattr(self, "lbl_fmp_api_key"):
                self.lbl_fmp_api_key.configure(text=t["lbl_fmp_api_key"])
            
            # JAUNUMS: Atjaunina IMAP konfigurācijas uzrakstus
            if hasattr(self, "lbl_imap_header"):
                self.lbl_imap_header.configure(text=t["lbl_imap_header"])
            if hasattr(self, "lbl_imap_server"):
                self.lbl_imap_server.configure(text=t["lbl_imap_server"])
            if hasattr(self, "lbl_imap_port"):
                self.lbl_imap_port.configure(text=t["lbl_imap_port"])
            if hasattr(self, "lbl_imap_user"):
                self.lbl_imap_user.configure(text=t["lbl_imap_user"])
            if hasattr(self, "lbl_imap_pass"):
                self.lbl_imap_pass.configure(text=t["lbl_imap_pass"])

            # Atjaunina saglabāšanas ziņojumu, ja tas ir redzams
            if self.lbl_saved.cget("text"):
                briefing_time = self.entry_briefing_time.get().strip() or "18:00"
                msg_fmt = t["saved_msg"]
                self.lbl_saved.configure(text=msg_fmt.format(time=briefing_time))

    def on_output_language_change(self, display_text):
        """Saglabā izvēlēto atskaišu un dokumentu ģenerēšanas valodu."""
        selected_output_code = _OUTPUT_LANG_DISPLAY_MAP.get(display_text, "auto")

        cfg = _load_config_data()
        
        cfg["output_language"] = selected_output_code
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
            
    # JAUNUMS: Radar režīma maiņas funkcija (nekavējoties saglabā config.json)
    def on_radar_mode_change(self, display_text):
        """Saglabā izvēlēto radara režīmu."""
        selected_radar_code = _RADAR_MODE_DISPLAY_MAP.get(display_text, "manual")

        cfg = _load_config_data()
        
        cfg["radar_mode"] = selected_radar_code
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


    def open_for_offer(self, subject, sender, body):
        self.deiconify()
        self.lift()
        self.focus_force()
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        self.tabview.set(t["tab_offers"])
        self.offer_module.load_request(subject, sender, body)

    def open_for_voice(self, audio_path):
        self.deiconify()
        self.lift()
        self.focus_force()
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        self.tabview.set(t["tab_voice"])
        
        if hasattr(self, "voice_module"):
            if hasattr(self.voice_module, "load_audio"):
                self.voice_module.load_audio(audio_path)
            elif hasattr(self.voice_module, "set_audio_file"):
                self.voice_module.set_audio_file(audio_path)
            elif hasattr(self.voice_module, "current_audio_path"):
                self.voice_module.current_audio_path = audio_path

    def setup_settings_tab(self):
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        cfg = _load_config_data() # Ielādējam config.json datus

        # 1. Ritināms ietvars
        self.settings_scroll = ctk.CTkScrollableFrame(self.tab_settings, fg_color="transparent")
        self.settings_scroll.pack(fill="both", expand=True, padx=5, pady=5)

        # 1. Bankas rekvizīti
        self.lbl_bank_sec = ctk.CTkLabel(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["lbl_bank_header"], 
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_bank_sec.pack(anchor="w", padx=25, pady=(15, 6))

        bank_row = ctk.CTkFrame(self.settings_scroll, fg_color="transparent") # Pārvietots uz scrollable frame
        bank_row.pack(fill="x", padx=25, pady=(0, 10))

        # IBAN lauks
        f_iban = ctk.CTkFrame(bank_row, fg_color="transparent")
        f_iban.pack(side="left", padx=(0, 20))
        self.lbl_iban = ctk.CTkLabel(f_iban, text=t["lbl_my_iban"], font=ctk.CTkFont(size=12))
        self.lbl_iban.pack(anchor="w", pady=(0, 2))
        self.entry_my_iban = ctk.CTkEntry(f_iban, width=240, height=32, font=ctk.CTkFont(size=13), placeholder_text="LV..PARX...")
        self.entry_my_iban.pack(anchor="w")

        # Uzņēmuma nosaukuma lauks
        f_name = ctk.CTkFrame(bank_row, fg_color="transparent")
        f_name.pack(side="left")
        self.lbl_name = ctk.CTkLabel(f_name, text=t["lbl_my_name"], font=ctk.CTkFont(size=12))
        self.lbl_name.pack(anchor="w", pady=(0, 2))
        self.entry_my_name = ctk.CTkEntry(f_name, width=280, height=32, font=ctk.CTkFont(size=13), placeholder_text="Mans Uzņēmums SIA")
        self.entry_my_name.pack(anchor="w")
        
        # JAUNUMS: Radara režīms
        self.lbl_radar_mode_header = ctk.CTkLabel(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["lbl_radar_mode"],
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_radar_mode_header.pack(anchor="w", padx=25, pady=(15, 6))

        self.radar_mode_optionmenu = ctk.CTkOptionMenu(
            self.settings_scroll, # Pārvietots uz scrollable frame
            values=list(_RADAR_MODE_DISPLAY_MAP.keys()),
            width=280,
            height=32,
            command=self.on_radar_mode_change
        )
        self.radar_mode_optionmenu.pack(anchor="w", padx=25, pady=(0, 15))
        
        # Ielādējam radara režīma vērtību un iestatām to
        initial_radar_mode_code = cfg.get("radar_mode", "manual")
        initial_radar_mode_display = _RADAR_MODE_CODE_MAP.get(initial_radar_mode_code, "Offline / Manual")
        self.radar_mode_optionmenu.set(initial_radar_mode_display)


        # JAUNUMS: IMAP konfigurācijas sadaļa (tiek pievienota PĒC radara režīma)
        self.lbl_imap_header = ctk.CTkLabel(
            self.settings_scroll,
            text=t["lbl_imap_header"],
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_imap_header.pack(anchor="w", padx=25, pady=(15, 6))

        imap_frame_server_port = ctk.CTkFrame(self.settings_scroll, fg_color="transparent")
        imap_frame_server_port.pack(fill="x", padx=25, pady=(0, 10))

        # IMAP Server
        f_imap_server = ctk.CTkFrame(imap_frame_server_port, fg_color="transparent")
        f_imap_server.pack(side="left", padx=(0, 20))
        self.lbl_imap_server = ctk.CTkLabel(f_imap_server, text=t["lbl_imap_server"], font=ctk.CTkFont(size=12))
        self.lbl_imap_server.pack(anchor="w", pady=(0, 2))
        self.entry_imap_server = ctk.CTkEntry(f_imap_server, width=240, height=32, font=ctk.CTkFont(size=13), placeholder_text="imap.gmail.com")
        self.entry_imap_server.pack(anchor="w")

        # IMAP Port
        f_imap_port = ctk.CTkFrame(imap_frame_server_port, fg_color="transparent")
        f_imap_port.pack(side="left")
        self.lbl_imap_port = ctk.CTkLabel(f_imap_port, text=t["lbl_imap_port"], font=ctk.CTkFont(size=12))
        self.lbl_imap_port.pack(anchor="w", pady=(0, 2))
        self.entry_imap_port = ctk.CTkEntry(f_imap_port, width=100, height=32, font=ctk.CTkFont(size=13), placeholder_text="993")
        self.entry_imap_port.pack(anchor="w")

        imap_frame_user_pass = ctk.CTkFrame(self.settings_scroll, fg_color="transparent")
        imap_frame_user_pass.pack(fill="x", padx=25, pady=(0, 15)) # Pievienots pady (0, 15), lai atdalītu no nākamās sadaļas

        # IMAP User
        f_imap_user = ctk.CTkFrame(imap_frame_user_pass, fg_color="transparent")
        f_imap_user.pack(side="left", padx=(0, 20))
        self.lbl_imap_user = ctk.CTkLabel(f_imap_user, text=t["lbl_imap_user"], font=ctk.CTkFont(size=12))
        self.lbl_imap_user.pack(anchor="w", pady=(0, 2))
        self.entry_imap_user = ctk.CTkEntry(f_imap_user, width=240, height=32, font=ctk.CTkFont(size=13), placeholder_text="vards@gmail.com")
        self.entry_imap_user.pack(anchor="w")

        # IMAP Password
        f_imap_pass = ctk.CTkFrame(imap_frame_user_pass, fg_color="transparent")
        f_imap_pass.pack(side="left")
        self.lbl_imap_pass = ctk.CTkLabel(f_imap_pass, text=t["lbl_imap_pass"], font=ctk.CTkFont(size=12))
        self.lbl_imap_pass.pack(anchor="w", pady=(0, 2))
        self.entry_imap_pass = ctk.CTkEntry(f_imap_pass, width=240, height=32, font=ctk.CTkFont(size=13), show="*", placeholder_text="****")
        self.entry_imap_pass.pack(anchor="w")

        # Ielādējam IMAP iestatījumus
        imap_settings = cfg.get("imap", {})
        self.entry_imap_server.insert(0, imap_settings.get("server", ""))
        self.entry_imap_port.insert(0, str(imap_settings.get("port", "")))
        self.entry_imap_user.insert(0, imap_settings.get("user", ""))
        self.entry_imap_pass.insert(0, imap_settings.get("password", ""))

        # JAUNUMS: API Atslēgu sadaļa (Pārkārtoju, lai IMAP būtu pirms API atslēgām)
        api_keys_frame = ctk.CTkFrame(self.settings_scroll, fg_color="transparent")
        api_keys_frame.pack(fill="x", padx=25, pady=(10, 15))
        
        # Gemini API Key
        f_gemini = ctk.CTkFrame(api_keys_frame, fg_color="transparent")
        f_gemini.pack(side="top", anchor="w", fill="x", pady=(0, 10))
        self.lbl_gemini_api_key = ctk.CTkLabel(f_gemini, text=t["lbl_gemini_api_key"], font=ctk.CTkFont(size=12))
        self.lbl_gemini_api_key.pack(anchor="w", pady=(0, 2))
        self.entry_gemini_api_key = ctk.CTkEntry(f_gemini, width=350, height=32, font=ctk.CTkFont(size=13), show="*")
        self.entry_gemini_api_key.pack(anchor="w")

        # Finnhub API Key
        f_finnhub = ctk.CTkFrame(api_keys_frame, fg_color="transparent")
        f_finnhub.pack(side="top", anchor="w", fill="x", pady=(0, 10)) # Pievienots pady
        self.lbl_finnhub_api_key = ctk.CTkLabel(f_finnhub, text=t["lbl_finnhub_api_key"], font=ctk.CTkFont(size=12))
        self.lbl_finnhub_api_key.pack(anchor="w", pady=(0, 2))
        self.entry_finnhub_api_key = ctk.CTkEntry(f_finnhub, width=350, height=32, font=ctk.CTkFont(size=13), show="*")
        self.entry_finnhub_api_key.pack(anchor="w")
        
        # JAUNUMS: Alpha Vantage API Key
        f_alpha = ctk.CTkFrame(api_keys_frame, fg_color="transparent")
        f_alpha.pack(side="top", anchor="w", fill="x", pady=(0, 10))
        self.lbl_alpha_api_key = ctk.CTkLabel(f_alpha, text=t["lbl_alpha_api_key"], font=ctk.CTkFont(size=12))
        self.lbl_alpha_api_key.pack(anchor="w", pady=(0, 2))
        self.entry_alpha_api_key = ctk.CTkEntry(f_alpha, width=350, height=32, font=ctk.CTkFont(size=13), show="*")
        self.entry_alpha_api_key.pack(anchor="w")

        # JAUNUMS: FMP API Key
        f_fmp = ctk.CTkFrame(api_keys_frame, fg_color="transparent")
        f_fmp.pack(side="top", anchor="w", fill="x") # Bez papildu pady, ja šis ir pēdējais
        self.lbl_fmp_api_key = ctk.CTkLabel(f_fmp, text=t["lbl_fmp_api_key"], font=ctk.CTkFont(size=12))
        self.lbl_fmp_api_key.pack(anchor="w", pady=(0, 2))
        self.entry_fmp_api_key = ctk.CTkEntry(f_fmp, width=350, height=32, font=ctk.CTkFont(size=13), show="*")
        self.entry_fmp_api_key.pack(anchor="w")


        # Ielādējam API atslēgas no config.json
        gemini_api_key = cfg.get("gemini_api_key") or cfg.get("aq_ai_api_key", "")
        finnhub_api_key = cfg.get("finnhub_api_key", "")
        alpha_vantage_api_key = cfg.get("alpha_vantage_api_key", "") # JAUNUMS
        fmp_api_key = cfg.get("fmp_api_key", "") # JAUNUMS

        self.entry_gemini_api_key.insert(0, gemini_api_key)
        self.entry_finnhub_api_key.insert(0, finnhub_api_key)
        self.entry_alpha_api_key.insert(0, alpha_vantage_api_key) # JAUNUMS
        self.entry_fmp_api_key.insert(0, fmp_api_key) # JAUNUMS


        # 2. Atskaites laiks (saglabā esošo pozīciju vai pielāgo, ja nepieciešams)
        self.lbl_time_header = ctk.CTkLabel(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["lbl_time"], 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_time_header.pack(anchor="w", padx=25, pady=(10, 4)) # Adjusted padding to fit new elements

        self.entry_briefing_time = ctk.CTkEntry(self.settings_scroll, width=120, height=32, font=ctk.CTkFont(size=13)) # Pārvietots uz scrollable frame
        self.entry_briefing_time.pack(anchor="w", padx=25, pady=(0, 10))

        # 3. Atslēgvārdi
        self.lbl_kw_header = ctk.CTkLabel(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["lbl_kw"], 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_kw_header.pack(anchor="w", padx=25, pady=(5, 4))

        self.lbl_desc_text = ctk.CTkLabel(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["lbl_desc"],
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        self.lbl_desc_text.pack(anchor="w", padx=25, pady=(0, 6))

        self.txt_keywords = ctk.CTkTextbox(self.settings_scroll, height=90, fg_color="#0f172a", font=ctk.CTkFont(size=13)) # Pārvietots uz scrollable frame
        self.txt_keywords.pack(fill="x", padx=25, pady=(0, 15))

        kw_list = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin"]
        briefing_time = "18:00"
        my_iban = "LV00PARX0000000000000"
        my_name = "Mans Uzņēmums SIA"
        # output_lang tiek ielādēta no config.json (piem., "auto", "lv", "uk")
        output_lang = get_initial_output_language() 

        # Ielādējam datus no config.json
        # cfg ir jau ielādēta sākumā `setup_settings_tab` metodē
        kw_list = cfg.get("keywords", kw_list)
        briefing_time = cfg.get("briefing_time", "18:00")
        my_iban = cfg.get("my_iban", my_iban)
        my_name = cfg.get("my_name", my_name)


        self.entry_my_iban.insert(0, my_iban)
        self.entry_my_name.insert(0, my_name)
        self.entry_briefing_time.insert(0, briefing_time)
        self.txt_keywords.insert("end", ", ".join(kw_list))

        self.btn_save_settings = ctk.CTkButton(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["btn_save"], 
            fg_color="#059669", 
            hover_color="#10b981", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=36,
            command=self.save_settings
        )
        self.btn_save_settings.pack(anchor="w", padx=25, pady=10)

        self.lbl_saved = ctk.CTkLabel(self.settings_scroll, text="", font=ctk.CTkFont(size=12), text_color="#10b981") # Pārvietots uz scrollable frame
        self.lbl_saved.pack(anchor="w", padx=25)

        # Atskaišu valodas iestatījumu sadaļa
        self.lbl_output_lang_header = ctk.CTkLabel(
            self.settings_scroll, # Pārvietots uz scrollable frame
            text=t["lbl_output_lang_header"],
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_output_lang_header.pack(anchor="w", padx=25, pady=(15, 6))

        self.output_lang_optionmenu = ctk.CTkOptionMenu(
            self.settings_scroll, # Pārvietots uz scrollable frame
            values=list(_OUTPUT_LANG_DISPLAY_MAP.keys()),
            width=280,
            height=32,
            command=self.on_output_language_change
        )
        self.output_lang_optionmenu.pack(anchor="w", padx=25, pady=(0, 15))
        
        # Iestata sākotnējo vērtību atskaišu valodas nolaižamajā izvēlnē
        # _OUTPUT_LANG_CODE_MAP kartē i18n kodu (piem., "auto", "lv", "uk") uz attēlojamo tekstu
        initial_output_display_text = _OUTPUT_LANG_CODE_MAP.get(output_lang, _OUTPUT_LANG_CODE_MAP["auto"])
        self.output_lang_optionmenu.set(initial_output_display_text)

    def save_settings(self):
        raw_kw = self.txt_keywords.get("1.0", "end").strip()
        kw = [x.strip() for x in raw_kw.split(",") if x.strip()]
        b_time = self.entry_briefing_time.get().strip() or "18:00"
        my_iban = self.entry_my_iban.get().strip()
        my_name = self.entry_my_name.get().strip()
        
        # JAUNUMS: Nolasa radara režīmu un API atslēgas
        selected_radar_display = self.radar_mode_optionmenu.get()
        radar_mode_code = _RADAR_MODE_DISPLAY_MAP.get(selected_radar_display, "manual")

        gemini_api_key = self.entry_gemini_api_key.get().strip()
        finnhub_api_key = self.entry_finnhub_api_key.get().strip()
        alpha_vantage_api_key = self.entry_alpha_api_key.get().strip() # JAUNUMS
        fmp_api_key = self.entry_fmp_api_key.get().strip() # JAUNUMS

        cfg = _load_config_data() # Ielādējam config.json datus drošā veidā

        cfg["keywords"] = kw
        cfg["briefing_time"] = b_time
        
        # Saglabā UI valodu config.json i18n standarta formātā
        current_config_lang_code = self.current_lang.lower()
        if current_config_lang_code == "ua":
            current_config_lang_code = "uk"
        cfg["ui_language"] = current_config_lang_code

        cfg["my_iban"] = my_iban
        cfg["my_name"] = my_name
        # output_language tiek saglabāta, izmantojot on_output_language_change, tāpēc šeit nav jāmaina

        # JAUNUMS: Saglabā radara režīmu un API atslēgas
        cfg["radar_mode"] = radar_mode_code
        cfg["gemini_api_key"] = gemini_api_key
        cfg["aq_ai_api_key"] = gemini_api_key  # Sinhronizācijai
        cfg["finnhub_api_key"] = finnhub_api_key
        cfg["alpha_vantage_api_key"] = alpha_vantage_api_key # JAUNUMS
        cfg["fmp_api_key"] = fmp_api_key # JAUNUMS

        # JAUNUMS: Saglabāt IMAP konfigurāciju
        imap_server = self.entry_imap_server.get().strip()
        imap_port_str = self.entry_imap_port.get().strip()
        imap_port = int(imap_port_str) if imap_port_str.isdigit() else 993
        imap_user = self.entry_imap_user.get().strip()
        imap_pass = self.entry_imap_pass.get().strip()

        cfg["imap"] = {
            "server": imap_server,
            "port": imap_port,
            "user": imap_user,
            "password": imap_pass
        }

        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        msg_fmt = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])["saved_msg"]
        self.lbl_saved.configure(text=msg_fmt.format(time=b_time))

    def transfer_warranty_to_offer(self, req_text):
        """Pārmet nolasītos iekārtas datus uz Piedāvājumu Studiju."""
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        self.tabview.set(t["tab_offers"])
        self.offer_module.load_request("Iekārtas apkope / garantija", "Tehniķis", req_text)

    def open_for_warranty(self, img_path):
        """Atver Garantijas cilni un sāk bildes analīzi."""
        self.deiconify()
        self.lift()
        self.focus_force()
        t = UI_TEXTS.get(self.current_lang, UI_TEXTS["LV"])
        self.tabview.set(t["tab_warranty"])
        if hasattr(self, "warranty_module"):
            self.warranty_module.load_and_analyze(img_path)