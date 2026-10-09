# -*- coding: utf-8 -*-
import os
import re
import urllib.request
import urllib.parse
import threading
import customtkinter as ctk
import google.generativeai as genai

# ==========================================
# HP MEKLĒŠANAS UN DATU IEGŪŠANAS FUNKCIJAS
# ==========================================
def fetch_hp_page(url):
    """Nolasa HP lapas tekstu tieši no saites."""
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

def search_hp_official(query_term):
    """Meklē HP oficiālajos avotos, izmantojot tīru detaļas kodu."""
    url_match = re.search(r'https?://[^\s]+hp\.com[^\s]+', query_term)
    if url_match:
        return fetch_hp_page(url_match.group(0))

    code_match = re.search(r'\b([A-Z0-9]{6,8}(?:#[A-Z0-9]+)?)\b', query_term.upper())
    search_target = code_match.group(1) if code_match else query_term

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

# ==========================================
# HP EKSPERTA UI CILNE (CTkFrame)
# ==========================================
class HPExpertFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None):
        super().__init__(master, fg_color="transparent")
        self.get_gemini_key_fn = get_gemini_key_fn

        self.chat_box = ctk.CTkTextbox(
            self, 
            font=ctk.CTkFont(size=13), 
            text_color="#f1f5f9", 
            fg_color="#0f172a",
            wrap="word"
        )
        self.chat_box.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        self.chat_box.insert("end", "HP Solution Architect gatavs darbam.\nUzdodiet jautājumu par EliteBook, ZBook, dokiem, RAM arhitektūru vai printeriem...\n\n")
        self.chat_box.configure(state="disabled")

        input_frame = ctk.CTkFrame(self, fg_color="transparent")
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

    def _get_api_key(self):
        if self.get_gemini_key_fn:
            key = self.get_gemini_key_fn()
            if key:
                return key
        return os.environ.get("GEMINI_API_KEY", "")

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

    def _query_gemini_hp(self, query):
        response_text = ""
        api_key = self._get_api_key()

        try:
            if not api_key:
                response_text = "⚠️ Gemini API atslēga nav atrasta! Pārbaudiet config.json failu."
            else:
                genai.configure(api_key=api_key)
                official_data = search_hp_official(query)

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
