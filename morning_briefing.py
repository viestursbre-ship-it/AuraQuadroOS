# -*- coding: utf-8 -*-
import os
import json
import sys
from pathlib import Path
import customtkinter as ctk
import google.generativeai as genai

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

def load_config():
    cfg_path = get_base_dir() / "config.json"
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def generate_kvark_briefing():
    base_dir = get_base_dir()
    cfg = load_config()
    api_key = cfg.get("gemini_api_key") or os.environ.get("GEMINI_API_KEY", "")
    
    state_file = base_dir / "radar_state.json"
    tasks = {}
    if state_file.exists():
        try:
            with open(state_file, "r", encoding="utf-8") as f:
                tasks = json.load(f)
        except Exception:
            tasks = {}

    active_tasks = [t for t in tasks.values() if t.get("status") == "ACTIVE"]

    if not active_tasks:
        return "🐾 Komandants Kvarks ziņo: Perimetrs tīrs! Nav neviena aktīva uzdevuma. Var mierīgi dzert espresso un uzraudzīt situāciju."

    task_summary = []
    for idx, t in enumerate(active_tasks[:10], 1):
        task_summary.append(
            f"{idx}. [{t.get('category', 'Uzdevums')}] No: {t.get('sender', '')} | Temats: {t.get('subject', '')} | Termiņš: {t.get('deadline', '')}"
        )
    task_text = "\n".join(task_summary)

    if not api_key:
        return f"🐾 Kvarka tiešā atskaite (bez AI):\nAktīvi uzdevumi ({len(active_tasks)}):\n" + task_text

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-2.5-flash")
        
        prompt = f"""
Tu esi Komandants Kvarks — melns, autoritatīvs, inteliģents un nedaudz prasīgs biroja patruļas vadītājs (kaķis-vadītājs no AuraQuadro štāba).
Tavs uzdevums ir dot komandai Rīta Espresso Brīfingu (Morning Briefing) par šodienas aktuālajiem uzdevumiem.
Lieto stingru, bet labsirdīgu un asprātīgu toni (ar nelieliem itāļu izteicieniem, piemēram: 'Avanti!', 'Attenzione!', 'Basta chiacchiere!').
Izveido 3 skaidrus punktus:
1. 🔥 Steidzamākie uzdevumi (rēķini apmaksai, garantijas, cenu pieprasījumi).
2. 🎯 Dienas galvenais fokuss.
3. 🐾 Komandanta Kvarka rīkojums dienai.

Lūk, uzdevumu saraksts no Radara:
{task_text}
"""
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"⚠️ Kvarks aizrijās ar putām: {e}\n\nAktīvie uzdevumi:\n{task_text}"


class KvarkBriefingDialog(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("☕ Komandanta Kvarka Brīfings")
        self.geometry("620x520")
        self.minsize(500, 400)
        self.configure(fg_color="#0f172a")

        header = ctk.CTkFrame(self, fg_color="#1e293b", corner_radius=0, height=50)
        header.pack(fill="x")

        ctk.CTkLabel(
            header,
            text="🐾 Komandanta Kvarka Operatīvā Atskaite",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#38bdf8"
        ).pack(side="left", padx=20, pady=12)

        self.txt_content = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(size=13),
            text_color="#f8fafc",
            fg_color="#1e293b",
            wrap="word"
        )
        self.txt_content.pack(fill="both", expand=True, padx=20, pady=15)
        self.txt_content.insert("1.0", "⏳ Komandants Kvarks apkopo situāciju pie trim monitoriem...")
        self.txt_content.configure(state="disabled")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkButton(
            btn_row,
            text="Paldies, Komandant! (Aizvērt)",
            fg_color="#059669",
            hover_color="#10b981",
            command=self.destroy
        ).pack(side="right")

        self.after(100, self.load_briefing_async)

    def load_briefing_async(self):
        import threading
        def _run():
            text = generate_kvark_briefing()
            self.after(0, lambda: self._update_text(text))
        threading.Thread(target=_run, daemon=True).start()

    def _update_text(self, text):
        self.txt_content.configure(state="normal")
        self.txt_content.delete("1.0", "end")
        self.txt_content.insert("1.0", text)
        self.txt_content.configure(state="disabled")
