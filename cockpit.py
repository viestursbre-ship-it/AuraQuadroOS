# -*- coding: utf-8 -*-
import os
import json
import customtkinter as ctk

from hp_expert import HPExpertFrame
from voice_studio import VoiceStudioFrame
from doc_digest import DocDigestFrame

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

class CockpitWindow(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("⚡ AQ-OS Cockpit — Vadības Centrs")
        self.geometry("1000x720")
        self.minsize(850, 600)

        # X poga tikai paslēpj lielo logu
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

        self.tabview = ctk.CTkTabview(self, fg_color="#1e293b", segmented_button_fg_color="#0f172a")
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)

        self.tab_hp = self.tabview.add("🧠 HP Eksperts")
        self.tab_voice = self.tabview.add("🎙️ Balss Studija")
        self.tab_docs = self.tabview.add("📄 Dokumentu Drop-Zone")
        self.tab_settings = self.tabview.add("⚙️ Radara Iestatījumi")

        # Moduļu izvietošana
        self.hp_module = HPExpertFrame(self.tab_hp, get_gemini_key_fn=get_gemini_api_key)
        self.hp_module.pack(fill="both", expand=True)

        self.voice_module = VoiceStudioFrame(self.tab_voice, get_gemini_key_fn=get_gemini_api_key)
        self.voice_module.pack(fill="both", expand=True)

        self.docs_module = DocDigestFrame(self.tab_docs, get_gemini_key_fn=get_gemini_api_key)
        self.docs_module.pack(fill="both", expand=True)

        # Iestatījumu cilnes noformējums
        self.setup_settings_tab()

    def setup_settings_tab(self):
        lbl = ctk.CTkLabel(
            self.tab_settings, 
            text="Radara Atslēgvārdi (atdalīti ar komatiem):", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl.pack(anchor="w", padx=25, pady=(25, 8))

        lbl_desc = ctk.CTkLabel(
            self.tab_settings,
            text="E-pasti, kuru temats vai teksts satur šos vārdus, automātiski pārtop par Radara uzdevumu kartītēm.",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        lbl_desc.pack(anchor="w", padx=25, pady=(0, 10))

        self.txt_keywords = ctk.CTkTextbox(self.tab_settings, height=130, fg_color="#0f172a", font=ctk.CTkFont(size=13))
        self.txt_keywords.pack(fill="x", padx=25, pady=(0, 15))

        kw_list = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin"]
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    kw_list = cfg.get("keywords", kw_list)
            except Exception:
                pass
        self.txt_keywords.insert("end", ", ".join(kw_list))

        btn_save = ctk.CTkButton(
            self.tab_settings, 
            text="💾 Saglabāt iestatījumus", 
            fg_color="#059669", 
            hover_color="#10b981", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=38,
            command=self.save_settings
        )
        btn_save.pack(anchor="w", padx=25, pady=10)

        self.lbl_saved = ctk.CTkLabel(self.tab_settings, text="", font=ctk.CTkFont(size=12), text_color="#10b981")
        self.lbl_saved.pack(anchor="w", padx=25)

    def save_settings(self):
        raw = self.txt_keywords.get("1.0", "end").strip()
        kw = [x.strip() for x in raw.split(",") if x.strip()]
        cfg = {}
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass
        cfg["keywords"] = kw
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        self.lbl_saved.configure(text="✅ Iestatījumi saglabāti! Radars tos nolasīs nākamajā ciklā.")
