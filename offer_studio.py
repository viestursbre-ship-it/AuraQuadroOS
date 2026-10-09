# -*- coding: utf-8 -*-
import os
import json
import threading
from pathlib import Path
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog
import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

try:
    from google import genai
    from google.genai import types
    HAS_GENAI_NEW = True
except ImportError:
    HAS_GENAI_NEW = False
    try:
        import google.generativeai as legacy_genai
        HAS_GENAI_OLD = True
    except ImportError:
        HAS_GENAI_OLD = False

BASE_DIR = Path(__file__).resolve().parent
STORAGE_DIR = BASE_DIR / "storage"
OUTBOX_DIR = STORAGE_DIR / "outbox"
OUTBOX_DIR.mkdir(parents=True, exist_ok=True)

def extract_offer_items(text_content, api_key):
    if not api_key:
        return {
            "client_name": "Klients",
            "delivery_days": "3-5 darba dienas",
            "items": [{"model": "Pieprasītā iekārta", "code": "", "qty": 1, "price": 0.0}]
        }

    prompt = f"""Analizē šo cenu pieprasījuma tekstu un izvelc strukturētu piedāvājuma saturu JSON formātā:
{{
  "client_name": "Klienta vārds vai uzņēmums (ja minēts, citādi Klients)",
  "delivery_days": "Paredzamais piegādes laiks (piem., 3-5 darba dienas)",
  "notes": "Piezīmes par garantiju vai specifikāciju",
  "items": [
    {{
      "model": "Iekārtas modelis/apraksts",
      "code": "HP vai ražotāja detaļas kods, ja atrasts",
      "qty": 1,
      "price": 0.0
    }}
  ]
}}
Ja konkrēta cena nav minēta, norādi 0.0.
Atgriez TIKAI derīgu JSON bez markdown blokiem.

Pieprasījuma teksts:
{text_content}"""

    try:
        if HAS_GENAI_NEW:
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.2)
            )
            raw = response.text.strip()
        elif HAS_GENAI_OLD:
            legacy_genai.configure(api_key=api_key)
            model = legacy_genai.GenerativeModel(
                model_name="gemini-2.5-flash",
                generation_config={"response_mime_type": "application/json", "temperature": 0.2}
            )
            res = model.generate_content(prompt)
            raw = res.text.strip()
        else:
            return {"error": "Nav pieejama Gemini bibliotēka"}

        return json.loads(raw)
    except Exception as e:
        return {"error": str(e), "items": [{"model": "Standarta pozīcija", "code": "", "qty": 1, "price": 0.0}]}

def fill_client_template(template_path: Path, data: dict, output_path: Path):
    """Mēģina aizpildīt klienta sagatavi (DOCX) vai ģenerē jaunu."""
    try:
        doc = docx.Document(template_path)
        # Ja klientam ir tabulas, meklējam pirmo un aizpildām rindas
        if doc.tables:
            table = doc.tables[0]
            for itm in data.get("items", []):
                row = table.add_row()
                cells = row.cells
                if len(cells) >= 3:
                    cells[0].text = str(itm.get("code") or itm.get("model", ""))
                    cells[1].text = str(itm.get("qty", 1))
                    cells[2].text = f"{float(itm.get('price', 0.0)):.2f}" if float(itm.get('price', 0.0)) > 0 else "—"
        doc.save(output_path)
        return output_path
    except Exception:
        return None

class OfferStudioFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None):
        super().__init__(master, fg_color="transparent")
        self.get_gemini_key_fn = get_gemini_key_fn
        self.current_analysis = {}
        self.client_template_path = None

        # Režīmu pārslēga panelis
        mode_bar = ctk.CTkFrame(self, fg_color="#0f172a", height=50)
        mode_bar.pack(fill="x", padx=10, pady=(10, 5))

        self.mode_var = ctk.StringVar(value="quick")
        self.seg_mode = ctk.CTkSegmentedButton(
            mode_bar,
            values=["⚡ Ātrais E-pasts", "📑 Klienta Forma / Veidne"],
            command=self.on_mode_change
        )
        self.seg_mode.set("⚡ Ātrais E-pasts")
        self.seg_mode.pack(side="left", padx=12, pady=8)

        self.btn_parse = ctk.CTkButton(
            mode_bar,
            text="🧠 Analizēt Pieprasījumu",
            width=160,
            height=32,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.analyze_input
        )
        self.btn_parse.pack(side="left", padx=6, pady=8)

        self.btn_action = ctk.CTkButton(
            mode_bar,
            text="📋 Kopēt E-pastam",
            width=150,
            height=32,
            fg_color="#059669",
            hover_color="#047857",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.execute_main_action
        )
        self.btn_action.pack(side="left", padx=6, pady=8)

        self.lbl_status = ctk.CTkLabel(mode_bar, text="", font=ctk.CTkFont(size=12), text_color="#38bdf8")
        self.lbl_status.pack(side="right", padx=15)

        # Klienta veidnes izvēles josla (redzama tikai 2. režīmā)
        self.tpl_bar = ctk.CTkFrame(self, fg_color="#1e293b", height=40)
        self.btn_load_tpl = ctk.CTkButton(
            self.tpl_bar,
            text="📁 Izvēlēties Klienta Formu (.docx)",
            width=200,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            command=self.select_template
        )
        self.btn_load_tpl.pack(side="left", padx=12, pady=6)
        self.lbl_tpl_name = ctk.CTkLabel(self.tpl_bar, text="Nav izvēlēts fails", font=ctk.CTkFont(size=12), text_color="#94a3b8")
        self.lbl_tpl_name.pack(side="left", padx=8)

        # Sadalījums 2 kolonnās
        panes = ctk.CTkFrame(self, fg_color="transparent")
        panes.pack(fill="both", expand=True, padx=10, pady=5)

        # Kreisā kolonna: Ieeja
        left_col = ctk.CTkFrame(panes, fg_color="#0b1329", corner_radius=8)
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 5))
        ctk.CTkLabel(left_col, text="📥 Pieprasījuma Saturs:", font=ctk.CTkFont(size=13, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=12, pady=(10, 4))
        self.txt_source = ctk.CTkTextbox(left_col, fg_color="#0f172a", font=ctk.CTkFont(size=12), wrap="word")
        self.txt_source.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Labā kolonna: Melnraksts
        right_col = ctk.CTkFrame(panes, fg_color="#0b1329", corner_radius=8)
        right_col.pack(side="right", fill="both", expand=True, padx=(5, 0))
        ctk.CTkLabel(right_col, text="⚡ Sagatavotā Atbilde / Dati:", font=ctk.CTkFont(size=13, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=12, pady=(10, 4))
        self.txt_preview = ctk.CTkTextbox(right_col, fg_color="#0f172a", font=ctk.CTkFont(size=12), wrap="word")
        self.txt_preview.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def on_mode_change(self, val):
        if "Klienta Forma" in val:
            self.mode_var.set("template")
            self.tpl_bar.pack(fill="x", padx=10, pady=(0, 5), before=self.txt_source.master.master)
            self.btn_action.configure(text="📄 Aizpildīt & Atvērt Formu", fg_color="#0284c7", hover_color="#0369a1")
        else:
            self.mode_var.set("quick")
            self.tpl_bar.pack_forget()
            self.btn_action.configure(text="📋 Kopēt E-pastam", fg_color="#059669", hover_color="#047857")

    def select_template(self):
        f = filedialog.askopenfilename(filetypes=[("Word Dokuments", "*.docx"), ("Visi faili", "*.*")])
        if f:
            self.client_template_path = Path(f)
            self.lbl_tpl_name.configure(text=self.client_template_path.name, text_color="#38bdf8")

    def load_request(self, subject, sender, body):
        self.txt_source.delete("1.0", "end")
        self.txt_source.insert("end", f"Temats: {subject}\nNo: {sender}\n\n{body}")
        self.lbl_status.configure(text="Pieprasījums ielādēts no Radara!")
        self.analyze_input()

    def analyze_input(self):
        text = self.txt_source.get("1.0", "end").strip()
        if not text:
            self.lbl_status.configure(text="Ievadiet tekstu!")
            return

        api_key = self.get_gemini_key_fn() if self.get_gemini_key_fn else os.environ.get("GEMINI_API_KEY", "")
        self.lbl_status.configure(text="Gemini analizē...")
        self.btn_parse.configure(state="disabled")

        def _worker():
            data = extract_offer_items(text, api_key)
            self.current_analysis = data
            def _done():
                self.btn_parse.configure(state="normal")
                self.lbl_status.configure(text="Gatavs rediģēšanai!")
                self._render_quick_email(data)
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    def _render_quick_email(self, data):
        self.txt_preview.delete("1.0", "end")
        lines = [
            f"Labdien, {data.get('client_name', 'klient')}!\n",
            "Paldies par pieprasījumu. Nosūtu piedāvājumu:\n",
            "-" * 65,
            f"{'Kods / Modelis':<28} | {'Skaits':<6} | {'Cena bez PVN':<12} | {'Kopā'}",
            "-" * 65
        ]
        for itm in data.get("items", []):
            name = itm.get("code") or itm.get("model", "")
            qty = itm.get("qty", 1)
            lines.append(f"{name:<28} | {qty:<6} | {'[ Cena ]':<12} | [ Summa ]")
        lines.extend([
            "-" * 65,
            f"\nPiegādes termiņš: {data.get('delivery_days', '3-5 darba dienas')}",
            "Cenas norādītas EUR bez PVN 21%.\n",
            "Ar cieņu,\nAuraQuadro"
        ])
        self.txt_preview.insert("end", "\n".join(lines))

    def execute_main_action(self):
        if self.mode_var.get() == "quick":
            txt = self.txt_preview.get("1.0", "end").strip()
            if txt:
                self.clipboard_clear()
                self.clipboard_append(txt)
                self.lbl_status.configure(text="📋 E-pasta tabula iekopēta!")
        else:
            if not self.client_template_path or not self.client_template_path.exists():
                self.lbl_status.configure(text="Izvēlieties klienta veidnes failu!")
                return
            out_file = OUTBOX_DIR / f"Aizpildits_{self.client_template_path.name}"
            res = fill_client_template(self.client_template_path, self.current_analysis, out_file)
            if res:
                self.lbl_status.configure(text="Forma aizpildīta!")
                os.startfile(str(res))
            else:
                self.lbl_status.configure(text="Kļūda formas aizpildē")
