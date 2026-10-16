# -*- coding: utf-8 -*-
import os
import json
import customtkinter as ctk

from hp_expert import HPExpertFrame
from voice_studio import VoiceStudioFrame
from doc_digest import DocDigestFrame
from offer_studio import OfferStudioFrame
from warranty_hub import WarrantyHubFrame
from bank_studio import BankStudioFrame

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
        "lbl_bank_header": "🏦 Bankas Rekvizīti (Citadele / ISO 20022):",
        "lbl_my_iban": "Maksātāja IBAN konts:",
        "lbl_my_name": "Maksātāja Uzņēmums / Vārds:",
        "lbl_time": "🕒 Vakara Komandanta Atskaites laiks (HH:MM):",
        "lbl_kw": "📡 Radara Atslēgvārdi (atdalīti ar komatiem):",
        "lbl_desc": "E-pasti, kuru temats vai teksts satur šos vārdus, automātiski pārtop par uzdevumiem.",
        "btn_save": "💾 Saglabāt iestatījumus",
        "saved_msg": "✅ Iestatījumi saglabāti! (Atskaite: {time})"
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
        "lbl_bank_header": "🏦 Banking Details (Citadele / ISO 20022):",
        "lbl_my_iban": "Payer IBAN Account:",
        "lbl_my_name": "Payer Company / Name:",
        "lbl_time": "🕒 Evening Commandant Briefing Time (HH:MM):",
        "lbl_kw": "📡 Radar Keywords (comma separated):",
        "lbl_desc": "Emails containing these words in subject or body automatically become tasks.",
        "btn_save": "💾 Save Settings",
        "saved_msg": "✅ Settings saved! (Briefing: {time})"
    }
}

def get_gemini_api_key():
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("AQ_AI_API_KEY", "")
    if not key and os.path.exists("config.json"):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                key = cfg.get("gemini_api_key", "")
        except Exception:
            pass
    return key

def get_initial_language():
    if os.path.exists("config.json"):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                cfg = json.load(f)
                return cfg.get("ui_language", "LV")
        except Exception:
            pass
    return "LV"

class CockpitWindow(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.current_lang = get_initial_language()
        
        self.title(UI_TEXTS[self.current_lang]["title"])
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

        # Globālais valodas pārslēgs labajā pusē
        self.lang_toggle = ctk.CTkSegmentedButton(
            header_frame,
            values=["LV", "EN"],
            width=90,
            height=28,
            command=self.on_language_change
        )
        self.lang_toggle.set(self.current_lang)
        self.lang_toggle.pack(side="right", padx=20, pady=15)

        self.tabview = ctk.CTkTabview(self, fg_color="#1e293b", segmented_button_fg_color="#0f172a")
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)

        # Sagatavojam cilnes ar sākotnējiem nosaukumiem
        t = UI_TEXTS[self.current_lang]
        self.tab_offers = self.tabview.add(t["tab_offers"])
        self.tab_hp = self.tabview.add(t["tab_hp"])
        self.tab_bank = self.tabview.add(t.get("tab_bank", "🏦 Bank Studio")) # <--- JAUNĀ CILNE ŠEIT!
        self.tab_voice = self.tabview.add(t["tab_voice"])
        self.tab_docs = self.tabview.add(t["tab_docs"])
        self.tab_warranty = self.tabview.add(t["tab_warranty"])
        self.tab_settings = self.tabview.add(t["tab_settings"])

        # Moduļu izvietošana
        self.offer_module = OfferStudioFrame(self.tab_offers, get_gemini_key_fn=get_gemini_api_key)
        self.offer_module.pack(fill="both", expand=True)

        self.hp_module = HPExpertFrame(self.tab_hp, get_gemini_key_fn=get_gemini_api_key)
        self.hp_module.pack(fill="both", expand=True)

        # <--- JAUNAIS BANKAS MODULIS ŠEIT!
        self.bank_module = BankStudioFrame(self.tab_bank, get_gemini_key_fn=get_gemini_api_key)
        self.bank_module.pack(fill="both", expand=True)

        self.voice_module = VoiceStudioFrame(self.tab_voice, get_gemini_key_fn=get_gemini_api_key)
        self.voice_module.pack(fill="both", expand=True)

        self.docs_module = DocDigestFrame(self.tab_docs, get_gemini_key_fn=get_gemini_api_key)
        self.docs_module.pack(fill="both", expand=True)

        # Jaunais Garantijas modulis ar piesaisti Piedāvājumiem
        self.warranty_module = WarrantyHubFrame(
            self.tab_warranty, 
            get_gemini_key_fn=get_gemini_api_key,
            on_send_to_offer_fn=self.transfer_warranty_to_offer
        )
        self.warranty_module.pack(fill="both", expand=True)

        self.setup_settings_tab()

    def on_language_change(self, lang):
        """Pārslēdz valodu, atjaunina ciļņu pogas un sinhronizē moduļus."""
        self.current_lang = lang
        t = UI_TEXTS[lang]
        self.title(t["title"])

        # 1. Atjaunina ciļņu pogu tekstus
        tab_keys = ["tab_offers", "tab_hp", "tab_voice", "tab_docs", "tab_warranty", "tab_settings"]
        button_list = list(self.tabview._segmented_button._buttons_dict.values())
        for idx, key in enumerate(tab_keys):
            if idx < len(button_list):
                button_list[idx].configure(text=t[key])

        # 2. Atjaunina Offer Studio moduli, ja tam ir pieejama metode
        if hasattr(self, "offer_module"):
            if hasattr(self.offer_module, "set_ui_language"):
                self.offer_module.set_ui_language(lang)
            elif hasattr(self.offer_module, "seg_lang"):
                self.offer_module.seg_lang.set(lang)

        # 3. Paziņo Radaram (master logam), ja tam ir valodas metode
        if self.master and hasattr(self.master, "set_ui_language"):
            try:
                self.master.set_ui_language(lang)
            except Exception:
                pass

        # 4. Saglabā izvēli config.json
        cfg = {}
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass
        cfg["ui_language"] = lang
        try:
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        # 5. Atjaunina Iestatījumu cilnes uzrakstus
        if hasattr(self, "lbl_time_header"):
            self.lbl_time_header.configure(text=t["lbl_time"])
            self.lbl_kw_header.configure(text=t["lbl_kw"])
            self.lbl_desc_text.configure(text=t["lbl_desc"])
            self.btn_save_settings.configure(text=t["btn_save"])

    def open_for_offer(self, subject, sender, body):
        self.deiconify()
        self.lift()
        self.focus_force()
        t = UI_TEXTS[self.current_lang]
        self.tabview.set(t["tab_offers"])
        self.offer_module.load_request(subject, sender, body)

    def open_for_voice(self, audio_path):
        self.deiconify()
        self.lift()
        self.focus_force()
        t = UI_TEXTS[self.current_lang]
        self.tabview.set(t["tab_voice"])
        
        if hasattr(self, "voice_module"):
            if hasattr(self.voice_module, "load_audio"):
                self.voice_module.load_audio(audio_path)
            elif hasattr(self.voice_module, "set_audio_file"):
                self.voice_module.set_audio_file(audio_path)
            elif hasattr(self.voice_module, "current_audio_path"):
                self.voice_module.current_audio_path = audio_path

    def setup_settings_tab(self):
        t = UI_TEXTS[self.current_lang]

        # 1. Bankas rekvizīti
        self.lbl_bank_sec = ctk.CTkLabel(
            self.tab_settings, 
            text=t["lbl_bank_header"], 
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_bank_sec.pack(anchor="w", padx=25, pady=(15, 6))

        bank_row = ctk.CTkFrame(self.tab_settings, fg_color="transparent")
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

        # 2. Atskaites laiks
        self.lbl_time_header = ctk.CTkLabel(
            self.tab_settings, 
            text=t["lbl_time"], 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_time_header.pack(anchor="w", padx=25, pady=(10, 4))

        self.entry_briefing_time = ctk.CTkEntry(self.tab_settings, width=120, height=32, font=ctk.CTkFont(size=13))
        self.entry_briefing_time.pack(anchor="w", padx=25, pady=(0, 10))

        # 3. Atslēgvārdi
        self.lbl_kw_header = ctk.CTkLabel(
            self.tab_settings, 
            text=t["lbl_kw"], 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.lbl_kw_header.pack(anchor="w", padx=25, pady=(5, 4))

        self.lbl_desc_text = ctk.CTkLabel(
            self.tab_settings,
            text=t["lbl_desc"],
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        self.lbl_desc_text.pack(anchor="w", padx=25, pady=(0, 6))

        self.txt_keywords = ctk.CTkTextbox(self.tab_settings, height=90, fg_color="#0f172a", font=ctk.CTkFont(size=13))
        self.txt_keywords.pack(fill="x", padx=25, pady=(0, 15))

        kw_list = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin"]
        briefing_time = "18:00"
        my_iban = "LV00PARX0000000000000"
        my_name = "Mans Uzņēmums SIA"

        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    kw_list = cfg.get("keywords", kw_list)
                    briefing_time = cfg.get("briefing_time", "18:00")
                    my_iban = cfg.get("my_iban", my_iban)
                    my_name = cfg.get("my_name", my_name)
            except Exception:
                pass

        self.entry_my_iban.insert(0, my_iban)
        self.entry_my_name.insert(0, my_name)
        self.entry_briefing_time.insert(0, briefing_time)
        self.txt_keywords.insert("end", ", ".join(kw_list))

        self.btn_save_settings = ctk.CTkButton(
            self.tab_settings, 
            text=t["btn_save"], 
            fg_color="#059669", 
            hover_color="#10b981", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=36,
            command=self.save_settings
        )
        self.btn_save_settings.pack(anchor="w", padx=25, pady=10)

        self.lbl_saved = ctk.CTkLabel(self.tab_settings, text="", font=ctk.CTkFont(size=12), text_color="#10b981")
        self.lbl_saved.pack(anchor="w", padx=25)

    def save_settings(self):
        raw_kw = self.txt_keywords.get("1.0", "end").strip()
        kw = [x.strip() for x in raw_kw.split(",") if x.strip()]
        b_time = self.entry_briefing_time.get().strip() or "18:00"
        my_iban = self.entry_my_iban.get().strip()
        my_name = self.entry_my_name.get().strip()

        cfg = {}
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass

        cfg["keywords"] = kw
        cfg["briefing_time"] = b_time
        cfg["ui_language"] = self.current_lang
        cfg["my_iban"] = my_iban
        cfg["my_name"] = my_name

        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        msg_fmt = UI_TEXTS[self.current_lang]["saved_msg"]
        self.lbl_saved.configure(text=msg_fmt.format(time=b_time))

    def transfer_warranty_to_offer(self, req_text):
        """Pārmet nolasītos iekārtas datus uz Piedāvājumu Studiju."""
        t = UI_TEXTS[self.current_lang]
        self.tabview.set(t["tab_offers"])
        self.offer_module.load_request("Iekārtas apkope / garantija", "Tehniķis", req_text)

    def open_for_warranty(self, img_path):
        """Atver Garantijas cilni un sāk bildes analīzi."""
        self.deiconify()
        self.lift()
        self.focus_force()
        t = UI_TEXTS[self.current_lang]
        self.tabview.set(t["tab_warranty"])
        if hasattr(self, "warranty_module"):
            self.warranty_module.load_and_analyze(img_path)
