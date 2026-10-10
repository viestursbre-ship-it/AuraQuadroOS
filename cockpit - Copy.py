# -*- coding: utf-8 -*-
import os
import customtkinter as ctk

from hp_expert import HPExpertFrame
from voice_studio import VoiceStudioFrame
from doc_digest import DocDigestFrame
from offer_studio import OfferStudioFrame

def get_gemini_api_key():
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("AQ_AI_API_KEY", "")
    if not key and os.path.exists("config.json"):
        try:
            import json
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

        self.tabview = ctk.CTkTabview(self, fg_color="#1e293b", segmented_button_fg_color="#0f172a")
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)

        # Ciļņu saraksts ar jauno Piedāvājumu Studiju
        self.tab_offers = self.tabview.add("💼 Piedāvājumi")
        self.tab_hp = self.tabview.add("🧠 HP Eksperts")
        self.tab_voice = self.tabview.add("🎙️ Balss Studija")
        self.tab_docs = self.tabview.add("📄 Dokumentu Drop-Zone")
        self.tab_settings = self.tabview.add("⚙️ Radara Iestatījumi")

        # Moduļu izvietošana
        self.offer_module = OfferStudioFrame(self.tab_offers, get_gemini_key_fn=get_gemini_api_key)
        self.offer_module.pack(fill="both", expand=True)

        self.hp_module = HPExpertFrame(self.tab_hp, get_gemini_key_fn=get_gemini_api_key)
        self.hp_module.pack(fill="both", expand=True)

        self.voice_module = VoiceStudioFrame(self.tab_voice, get_gemini_key_fn=get_gemini_api_key)
        self.voice_module.pack(fill="both", expand=True)

        self.docs_module = DocDigestFrame(self.tab_docs, get_gemini_key_fn=get_gemini_api_key)
        self.docs_module.pack(fill="both", expand=True)

        self.setup_settings_tab()

    def open_for_offer(self, subject, sender, body):
        """Atver Cockpit, pārslēdz uz Piedāvājumu cilni un ielādē e-pasta datus."""
        self.deiconify()
        self.lift()
        self.focus_force()
        self.tabview.set("💼 Piedāvājumi")
        self.offer_module.load_request(subject, sender, body)

    def open_for_voice(self, audio_path):
        """Atver Cockpit, pārslēdz uz Balss Studiju un padod audio failu."""
        self.deiconify()
        self.lift()
        self.focus_force()
        self.tabview.set("🎙️ Balss Studija")
        
        # Mēģinām ielādēt failu atkarībā no metodes nosaukuma tavā voice_studio
        if hasattr(self, "voice_module"):
            if hasattr(self.voice_module, "load_audio"):
                self.voice_module.load_audio(audio_path)
            elif hasattr(self.voice_module, "set_audio_file"):
                self.voice_module.set_audio_file(audio_path)
            elif hasattr(self.voice_module, "current_audio_path"):
                self.voice_module.current_audio_path = audio_path

    def setup_settings_tab(self):
        # 1. Atskaites laiks
        lbl_time = ctk.CTkLabel(
            self.tab_settings, 
            text="🕒 Vakara Komandanta Atskaites laiks (HH:MM):", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_time.pack(anchor="w", padx=25, pady=(20, 4))

        self.entry_briefing_time = ctk.CTkEntry(self.tab_settings, width=120, height=32, font=ctk.CTkFont(size=13))
        self.entry_briefing_time.pack(anchor="w", padx=25, pady=(0, 15))

        # 2. Atslēgvārdi
        lbl_kw = ctk.CTkLabel(
            self.tab_settings, 
            text="📡 Radara Atslēgvārdi (atdalīti ar komatiem):", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_kw.pack(anchor="w", padx=25, pady=(10, 4))

        lbl_desc = ctk.CTkLabel(
            self.tab_settings,
            text="E-pasti, kuru temats vai teksts satur šos vārdus, automātiski pārtop par uzdevumiem.",
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        lbl_desc.pack(anchor="w", padx=25, pady=(0, 8))

        self.txt_keywords = ctk.CTkTextbox(self.tab_settings, height=110, fg_color="#0f172a", font=ctk.CTkFont(size=13))
        self.txt_keywords.pack(fill="x", padx=25, pady=(0, 15))

        # Ielādējam esošās vērtības no config.json
        kw_list = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin"]
        briefing_time = "18:00"

        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    kw_list = cfg.get("keywords", kw_list)
                    briefing_time = cfg.get("briefing_time", "18:00")
            except Exception:
                pass

        self.entry_briefing_time.insert(0, briefing_time)
        self.txt_keywords.insert("end", ", ".join(kw_list))

        # Saglabāšanas poga
        btn_save = ctk.CTkButton(
            self.tab_settings, 
            text="💾 Saglabāt iestatījumus", 
            fg_color="#059669", 
            hover_color="#10b981", 
            font=ctk.CTkFont(size=13, weight="bold"),
            height=36,
            command=self.save_settings
        )
        btn_save.pack(anchor="w", padx=25, pady=10)

        self.lbl_saved = ctk.CTkLabel(self.tab_settings, text="", font=ctk.CTkFont(size=12), text_color="#10b981")
        self.lbl_saved.pack(anchor="w", padx=25)

    def save_settings(self):
        raw_kw = self.txt_keywords.get("1.0", "end").strip()
        kw = [x.strip() for x in raw_kw.split(",") if x.strip()]
        b_time = self.entry_briefing_time.get().strip() or "18:00"

        cfg = {}
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass

        cfg["keywords"] = kw
        cfg["briefing_time"] = b_time

        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

        self.lbl_saved.configure(text=f"✅ Iestatījumi saglabāti! (Atskaite iestatīta uz {b_time})")

    def save_settings(self):
        raw = self.txt_keywords.get("1.0", "end").strip()
        kw = [x.strip() for x in raw.split(",") if x.strip()]
        cfg = {}
        if os.path.exists("config.json"):
            try:
                import json
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                pass
        cfg["keywords"] = kw
        with open("config.json", "w", encoding="utf-8") as f:
            import json
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        self.lbl_saved.configure(text="✅ Iestatījumi saglabāti! Radars tos nolasīs nākamajā ciklā.")
