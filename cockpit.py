# -*- coding: utf-8 -*-
import os
import json
import threading
from tkinter import filedialog
import customtkinter as ctk
import google.generativeai as genai
from voice_studio import VoiceStudioFrame

# ==========================================
# KONFIGURĀCIJA UN GEMINI INICIALIZĀCIJA
# ==========================================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

GEMINI_KEY = os.environ.get("GEMINI_API_KEY", "")
if not GEMINI_KEY and os.path.exists("config.json"):
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            cfg = json.load(f)
            GEMINI_KEY = cfg.get("gemini_api_key", "")
    except Exception:
        pass

if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)

# Vienotā HP Eksperta sistēmas instruktāža
HP_SYSTEM_PROMPT = """
Tu esi augstākā līmeņa HP Inc. tehniskais analītiķis un datu lapu eksperts.
TEV OBLIGĀTI JĀVEIC MEKLĒŠANA TĪMEKLĪ (site:hp.com, site:support.hp.com vai QuickSpecs).

NOTEIKUMI:
1. Ja norādīts detaļas kods (Part Number, piem., E11NXAT) vai modelis:
   - Atrodi konkrēto lapu oficiālajā HP vietnē un pārbaudi precīzo šasiju.
2. Atmiņa (RAM):
   - Pārbaudi, vai norādīts 'DDR5 SODIMM' (maināms, sloti) vai 'LPDDR5/LPDDR5x' (soldered / ielodēts).
   - EliteBook 600 sērijai un 6 Gxx modeļiem standarta konfigurācijās ir SODIMM sloti.
3. Doki, porti un barošana:
   - Norādi precīzu portu skaitu (Thunderbolt 4 / USB-C, barošanas jaudas) un saderīgos dokus.
4. Ja kods nav atrodams oficiālajos avotos, godīgi pasaki to un lūdz precizēt modeli.
Atbildi sniedz latviešu valodā, strukturēti un kodolīgi.
"""

# ==========================================
# GALVENAIS COCKPIT LOGS
# ==========================================
class AQCockpitApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("⚡ AQ-OS Cockpit — Vadības Centrs")
        self.geometry("950x700")
        self.minsize(800, 600)

        header_frame = ctk.CTkFrame(self, fg_color="#0f172a", height=60, corner_radius=0)
        header_frame.pack(fill="x", side="top")

        title_label = ctk.CTkLabel(
            header_frame, 
            text="⚡ AQ-OS MISSION CONTROL", 
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#38bdf8"
        )
        title_label.pack(side="left", padx=20, pady=15)

        subtitle_label = ctk.CTkLabel(
            header_frame, 
            text="v2.0 Universal Hub", 
            font=ctk.CTkFont(size=12),
            text_color="#94a3b8"
        )
        subtitle_label.pack(side="left", padx=5, pady=15)

        self.tabview = ctk.CTkTabview(self, fg_color="#1e293b", segmented_button_fg_color="#0f172a")
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)

        # Šādi beidzas __init__ metode:
        self.tab_hp = self.tabview.add("🧠 HP Eksperts")
        self.tab_voice = self.tabview.add("🎙️ Balss Studija")
        self.tab_docs = self.tabview.add("📄 Dokumentu Drop-Zone")
        self.tab_settings = self.tabview.add("⚙️ Radara Iestatījumi")

        self.setup_hp_tab()
        self.setup_voice_tab()
        self.setup_docs_tab()
        self.setup_settings_tab()

    # Un šī ir atsevišķa metode zem __init__ (vienā līmenī ar setup_hp_tab):
    def setup_voice_tab(self):
        self.voice_studio = VoiceStudioFrame(self.tab_voice, get_gemini_key_fn=lambda: GEMINI_KEY)
        self.voice_studio.pack(fill="both", expand=True)

    # ------------------------------------------
    # 1. HP EKSPERTA CILNE
    # ------------------------------------------
    def setup_hp_tab(self):
        self.chat_box = ctk.CTkTextbox(
            self.tab_hp, 
            font=ctk.CTkFont(size=13), 
            text_color="#f1f5f9", 
            fg_color="#0f172a",
            wrap="word"
        )
        self.chat_box.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        self.chat_box.insert("end", "HP Solution Architect gatavs darbam.\nUzdodiet jautājumu par EliteBook, ZBook, dokiem, RAM arhitektūru vai printeriem...\n\n")
        self.chat_box.configure(state="disabled")

        input_frame = ctk.CTkFrame(self.tab_hp, fg_color="transparent")
        input_frame.pack(fill="x", padx=10, pady=10)

        self.hp_input = ctk.CTkEntry(
            input_frame, 
            placeholder_text="Piemēram: Vai E11NXAT ir ar ielodētu RAM? Kāds doks der ZBook Fury?",
            font=ctk.CTkFont(size=13),
            height=40
        )
        self.hp_input.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.hp_input.bind("<Return>", lambda e: self.ask_hp_expert())

        self.btn_ask = ctk.CTkButton(
            input_frame, 
            text="Jautāt ➔", 
            width=100, 
            height=40,
            command=self.ask_hp_expert
        )
        self.btn_ask.pack(side="right")

    def ask_hp_expert(self):
        query = self.hp_input.get().strip()
        if not query:
            return

        self.hp_input.delete(0, "end")
        self.chat_box.configure(state="normal")
        self.chat_box.insert("end", f"\n👤 Jūs: {query}\n")
        self.chat_box.insert("end", "⏳ HP Eksperts meklē oficiālajos avotos...\n")
        self.chat_box.see("end")
        self.chat_box.configure(state="disabled")

        threading.Thread(target=self._query_gemini_hp, args=(query,), daemon=True).start()

    def _search_hp_official(self, query_term):
        """Meklē datus TIKAI oficiālajos HP/HPE domēnos bez trešo pušu e-veikaliem."""
        import urllib.request
        import urllib.parse
        import re

        # Stingrs domēnu filtrs
        site_filter = "(site:hp.com OR site:support.hp.com OR site:hpe.com OR site:partsurfer.hp.com)"
        full_query = f"{query_term} {site_filter}"
        
        url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(full_query)
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        )
        
        try:
            with urllib.request.urlopen(req, timeout=7) as response:
                html = response.read().decode('utf-8', errors='ignore')
                
                # Izvelkam oficiālo rezultātu teksta fragmentus
                snippets = re.findall(r'<a class="result__snippet[^>]*>(.*?)</a>', html, re.DOTALL)
                clean_snippets = []
                for s in snippets[:4]:
                    text = re.sub(r'<[^>]+>', '', s).strip()
                    if text:
                        clean_snippets.append(text)
                
                return "\n---\n".join(clean_snippets) if clean_snippets else ""
        except Exception:
            return ""

    def _fetch_hp_page(self, url):
        """Nolasa HP lapas tekstu tieši no saites."""
        import urllib.request
        import re
        try:
            req = urllib.request.Request(
                url, 
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            )
            with urllib.request.urlopen(req, timeout=8) as res:
                html = res.read().decode('utf-8', errors='ignore')
                text = re.sub(r'<script.*?</script>', '', html, flags=re.DOTALL)
                text = re.sub(r'<style.*?</style>', '', text, flags=re.DOTALL)
                text = re.sub(r'<[^>]+>', ' ', text)
                clean_text = ' '.join(text.split())
                return clean_text[:6000]
        except Exception:
            return ""

    def _search_hp_official(self, query_term):
        """Meklē HP oficiālajos avotos, izmantojot tīru detaļas kodu."""
        import urllib.request
        import urllib.parse
        import re

        # Ja tekstā jau ir HP saite, lasām to tieši
        url_match = re.search(r'https?://[^\s]+hp\.com[^\s]+', query_term)
        if url_match:
            return self._fetch_hp_page(url_match.group(0))

        # Izvelkam konkrēto produkta kodu (piem., DG0L7AT vai E11NXAT)
        code_match = re.search(r'\b([A-Z0-9]{6,8}(?:#[A-Z0-9]+)?)\b', query_term.upper())
        search_target = code_match.group(1) if code_match else query_term

        # Vaicājums tikai ar kodu un ražotāju
        search_url = f"https://html.duckduckgo.com/html/?q={urllib.parse.quote(search_target + ' site:hp.com')}"
        req = urllib.request.Request(
            search_url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0'}
        )
        
        try:
            with urllib.request.urlopen(req, timeout=8) as response:
                html = response.read().decode('utf-8', errors='ignore')
                snippets = re.findall(r'<a class="result__snippet[^>]*>(.*?)</a>', html, re.DOTALL)
                clean_snippets = [re.sub(r'<[^>]+>', '', s).strip() for s in snippets[:5]]
                return "\n---\n".join([s for s in clean_snippets if s])
        except Exception:
            return ""

    def _query_gemini_hp(self, query):
        response_text = ""
        try:
            if not GEMINI_KEY:
                response_text = "⚠️ Gemini API atslēga nav atrasta! Pārbaudiet config.json failu."
            else:
                # 1. solis: Meklējam oficiālajos avotos
                official_data = self._search_hp_official(query)

                hp_strict_prompt = (
                    "Tu esi sertificēts HP Inc. sistēmu inženieris un aparatūras arhitekts.\n\n"
                    "ZINĀŠANU BĀZE PAR JAUNAJIEM MODEĻIEM:\n"
                    "- HP EliteBook 8 G2i / 6 G2i Next Gen AI PC ir Intel Core Ultra bāzēti portatīvie datori.\n"
                    "- EliteBook 8 G2i izmanto LPDDR5x (onboard / ielodētu RAM).\n"
                    "- EliteBook 6 G2i izmanto DDR5 SODIMM (maināmu RAM).\n"
                    "- Savietojamība ar monitoriem: Šiem datoriem ir Thunderbolt 4 / USB-C ar DisplayPort 1.4 atbalstu un HDMI 2.1. "
                    "Tie ir savietojami ar jebkuru modernu monitoru (HP E-sērijas USB-C ar Power Delivery, HP Series 7 Pro, Z-sērijas 4K displejiem).\n\n"
                    "NOTEIKUMI:\n"
                    "1. Ja norādīts produkta kods (piemēram, DG0L7AT), tas pieder HP EliteBook 8 G2i 14 AI PC sērijai.\n"
                    "2. Atbildi uz lietotāja jautājumu konkrēti, strukturēti un latviešu valodā."
                )

                model = genai.GenerativeModel(
                    model_name="gemini-2.5-flash",
                    system_instruction=hp_strict_prompt
                )

                prompt = f"Lietotāja jautājums: {query}\n\nOficiālie atrastie dati:\n{official_data if official_data else 'Nav papildu tīmekļa fragmentu'}"
                res = model.generate_content(prompt)
                response_text = res.text

        except Exception as e:
            response_text = f"⚠️ Kļūda saziņā ar ekspertu: {e}"

        def update_ui():
            self.chat_box.configure(state="normal")
            content = self.chat_box.get("1.0", "end")
            idx = content.rfind("⏳ HP Eksperts meklē oficiālajos avotos...\n")
            if idx != -1:
                self.chat_box.delete(f"1.0 + {idx} chars", "end")
            self.chat_box.insert("end", f"🧠 HP Eksperts:\n{response_text.strip()}\n\n")
            self.chat_box.see("end")
            self.chat_box.configure(state="disabled")

        self.after(0, update_ui)

    # ------------------------------------------
    # 2. DOKUMENTU DROP-ZONE CILNE
    # ------------------------------------------
    def setup_docs_tab(self):
        drop_frame = ctk.CTkFrame(self.tab_docs, fg_color="#0f172a", corner_radius=12, border_width=2, border_color="#334155")
        drop_frame.pack(fill="both", expand=True, padx=40, pady=20)

        lbl_icon = ctk.CTkLabel(drop_frame, text="📥", font=ctk.CTkFont(size=42))
        lbl_icon.pack(pady=(20, 5))

        lbl_text = ctk.CTkLabel(
            drop_frame, 
            text="Izvēlieties līgumu, tehnisko specifikāciju vai tāmi (PDF / DOCX / TXT)", 
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#cbd5e1"
        )
        lbl_text.pack(pady=(0, 10))

        btn_select = ctk.CTkButton(
            drop_frame, 
            text="Izvēlēties failu no datora", 
            fg_color="#2563eb", 
            width=220, 
            height=36,
            command=self.select_doc_file
        )
        btn_select.pack(pady=(0, 15))

        self.lbl_selected_doc = ctk.CTkLabel(drop_frame, text="Nav izvēlēts neviens dokuments", text_color="#94a3b8")
        self.lbl_selected_doc.pack(pady=(0, 10))

        # Analīzes logs dokumentam
        self.txt_doc_output = ctk.CTkTextbox(self.tab_docs, font=ctk.CTkFont(size=13), fg_color="#0f172a", text_color="#f1f5f9", wrap="word")
        self.txt_doc_output.pack(fill="both", expand=True, padx=40, pady=(0, 20))

    def select_doc_file(self):
        f = filedialog.askopenfilename(
            filetypes=[
                ("Dokumenti", "*.pdf *.docx *.txt"),
                ("Visi faili", "*.*")
            ]
        )
        if f:
            filename = os.path.basename(f)
            self.lbl_selected_doc.configure(text=f"Atlasīts: {filename}", text_color="#38bdf8")
            self.txt_doc_output.delete("1.0", "end")
            self.txt_doc_output.insert("end", f"📄 Fails '{filename}' gatavs apstrādei.\nNākamajā solī šeit tiks ģenerēts kopsavilkums un prasību audits.")

    # ------------------------------------------
    # 3. RADARA IESTATĪJUMU CILNE
    # ------------------------------------------
    def setup_settings_tab(self):
        lbl = ctk.CTkLabel(self.tab_settings, text="Radara Atslēgvārdi (atdalīti ar komatiem):", font=ctk.CTkFont(size=14, weight="bold"))
        lbl.pack(anchor="w", padx=20, pady=(20, 5))

        self.txt_keywords = ctk.CTkTextbox(self.tab_settings, height=120, fg_color="#0f172a")
        self.txt_keywords.pack(fill="x", padx=20, pady=(0, 15))

        kw_list = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin"]
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    kw_list = cfg.get("keywords", kw_list)
            except Exception:
                pass
        self.txt_keywords.insert("end", ", ".join(kw_list))

        btn_save = ctk.CTkButton(self.tab_settings, text="💾 Saglabāt iestatījumus", fg_color="#059669", hover_color="#10b981", command=self.save_settings)
        btn_save.pack(anchor="w", padx=20, pady=10)

        self.lbl_saved = ctk.CTkLabel(self.tab_settings, text="", text_color="#10b981")
        self.lbl_saved.pack(anchor="w", padx=20)

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

if __name__ == "__main__":
    app = AQCockpitApp()
    app.mainloop()
