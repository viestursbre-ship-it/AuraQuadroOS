import os
import sys
import time
import json
import shutil
import base64
import threading
from pathlib import Path
from datetime import datetime

# GUI un sistēmas joslas bibliotēkas
import customtkinter as ctk
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item

# Biroja un dokumentu bibliotēkas
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
import openpyxl
from openpyxl.styles import Font

# MI un Windows COM integrācija
from google import genai
from google.genai import types

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# ==========================================
# KONFIGURĀCIJA UN CEĻI
# ==========================================
STORAGE_DIR = Path("storage")
INBOX_DIR = STORAGE_DIR / "inbox"
OUTBOX_DIR = STORAGE_DIR / "outbox"
ARCHIVE_DIR = STORAGE_DIR / "archive"
EXCEL_REGISTRY_PATH = STORAGE_DIR / "Ligumu_Registrs.xlsx"
STATE_FILE = "radar_state.json"

for d in [INBOX_DIR, OUTBOX_DIR, ARCHIVE_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ==========================================
# 1. DOKUMENTU APSTRĀDES MODULIS
# ==========================================
def safe_archive_file(source_path: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    destination = target_dir / source_path.name
    if destination.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destination = target_dir / f"{source_path.stem}_{timestamp}{source_path.suffix}"
    shutil.move(str(source_path), str(destination))
    return destination

def append_to_excel_registry(file_name: str, analysis: dict, status: str = "Apstrādāts"):
    headers = ["Ieraksta Datums", "Fails", "Dokumenta tips", "Dokumenta Nr.", 
               "Dokumenta Datums", "Puses", "Summa", "Termiņš", "Statuss"]
    try:
        if not EXCEL_REGISTRY_PATH.exists():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Līgumu Reģistrs"
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True)
            wb.save(EXCEL_REGISTRY_PATH)

        wb = openpyxl.load_workbook(EXCEL_REGISTRY_PATH)
        ws = wb.active

        parties_data = analysis.get("parties", [])
        if isinstance(parties_data, list):
            parties_str = ", ".join([p.get("name", "").strip() if isinstance(p, dict) else str(p) for p in parties_data])
        else:
            parties_str = str(parties_data or "N/A")

        fin = analysis.get("financial_terms", {})
        sum_val = fin.get("total_amount", "N/A") if isinstance(fin, dict) else "N/A"
        term_val = fin.get("deadlines", "N/A") if isinstance(fin, dict) else "N/A"

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            file_name,
            str(analysis.get("document_type", "N/A")),
            str(analysis.get("document_number", "N/A")),
            str(analysis.get("document_date", "N/A")),
            parties_str,
            str(sum_val),
            str(term_val),
            status
        ]
        ws.append(row)
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
        wb.save(EXCEL_REGISTRY_PATH)
    except Exception as e:
        print(f"⚠️ Kļūda Excel reģistrācijā: {e}")

class DocDigestService:
    def _read_docx(self, path: Path) -> str:
        doc = docx.Document(path)
        return "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])

    def _read_txt(self, path: Path) -> str:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    def _read_pptx(self, path: Path) -> str:
        prs = Presentation(path)
        full_text = []
        for i, slide in enumerate(prs.slides, 1):
            full_text.append(f"\n--- Slaids {i} ---")
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            full_text.append(paragraph.text.strip())
        return "\n".join(full_text)

    def _analyze_with_gemini(self, content_payload, file_type: str = "docx") -> dict:
        api_key = os.getenv("AQ_AI_API_KEY")
        if not api_key:
            return {"error": "AQ_AI_API_KEY nav iestatīts."}

        client = genai.Client(api_key=api_key)
        prompt_text = """Analizējiet šo dokumentu un sagatavojiet precīzu JSON objektu:
{
  "document_type": "Līgums / Rēķins / Akts / Cits",
  "document_number": "Numurs vai N/A",
  "document_date": "Datums vai N/A",
  "parties": [{"name": "Puses nosaukums", "role": "Pasūtītājs/Piegādātājs"}],
  "subject": "Darījuma priekšmets",
  "financial_terms": {"total_amount": "Summa ar PVN", "deadlines": "Apmaksas/Izpildes termiņi"},
  "risks_warnings": ["Konstatētie riski vai brīdinājumi"]
}
Atgrieziet TIKAI derīgu JSON bez markdown blokiem."""

        try:
            if file_type == "pdf":
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=[prompt_text, content_payload],
                    config=types.GenerateContentConfig(temperature=0.2)
                )
            else:
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=f"{prompt_text}\n\nDokumenta saturs:\n{content_payload[:12000]}",
                    config=types.GenerateContentConfig(temperature=0.2)
                )
            if response and response.text:
                clean_json = response.text.replace("```json", "").replace("```", "").strip()
                return json.loads(clean_json)
            return {"error": "Tukša atbilde no MI"}
        except Exception as e:
            return {"error": str(e)}

    def _create_digest_docx(self, output_path: Path, data: dict, original_filename: str):
        doc = docx.Document()
        doc.add_heading(f"Kopsavilkums: {original_filename}", level=1)
        doc.add_paragraph(f"Izveidots: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        def add_sec(title, val):
            doc.add_heading(title, level=2)
            if isinstance(val, dict):
                for k, v in val.items():
                    doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: {v}")
            elif isinstance(val, list):
                for item in val:
                    doc.add_paragraph(f"• {item}")
            else:
                doc.add_paragraph(str(val or "N/A"))

        add_sec("Vispārīgā informācija", {
            "Tips": data.get("document_type", "N/A"),
            "Numurs": data.get("document_number", "N/A"),
            "Datums": data.get("document_date", "N/A"),
            "Priekšmets": data.get("subject", "N/A")
        })
        add_sec("Finanšu noteikumi", data.get("financial_terms", {}))
        add_sec("Riski un brīdinājumi", data.get("risks_warnings", []))
        doc.save(output_path)

    def process_file(self, file_path: Path):
        ext = file_path.suffix.lower()
        text_payload = None

        if ext == ".docx":
            text_payload = self._read_docx(file_path)
        elif ext == ".txt":
            text_payload = self._read_txt(file_path)
        elif ext == ".pptx":
            text_payload = self._read_pptx(file_path)
        elif ext == ".pdf":
            with open(file_path, "rb") as f:
                pdf_bytes = f.read()
            text_payload = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
        else:
            return

        if not text_payload:
            return

        analysis = self._analyze_with_gemini(text_payload, file_type=ext.replace(".", ""))
        if "error" in analysis:
            print(f"⚠️️ MI kļūda failam {file_path.name}: {analysis['error']}")
            return

        out_name = f"DIGEST_{file_path.stem}.docx"
        self._create_digest_docx(OUTBOX_DIR / out_name, analysis, file_path.name)
        append_to_excel_registry(file_path.name, analysis)
        safe_archive_file(file_path, ARCHIVE_DIR)
        print(f"✅ [DOKUMENTI] Apstrādāts: {file_path.name} -> {out_name}")

class FolderWatcherService:
    def __init__(self, watch_dir: Path, digest_service: DocDigestService):
        self.watch_dir = watch_dir
        self.digest_service = digest_service
        self.running = False

    def start_watching(self, interval_sec: int = 5):
        self.running = True
        print(f"👀 [DOKUMENTI] Mape {self.watch_dir} tiek uzraudzīta...")
        while self.running:
            try:
                for file_path in self.watch_dir.iterdir():
                    if file_path.is_file() and not file_path.name.startswith("~"):
                        self.digest_service.process_file(file_path)
            except Exception as e:
                print(f"⚠️ Uzrauga kļūda: {e}")
            time.sleep(interval_sec)

# ==========================================
# 2. OUTLOOK RADARA SKENERIS (DROŠS PRET AVĀRIJĀM)
# ==========================================
def scan_outlook_mailbox():
    """Pārbauda Outlook kasti tikai tad, ja Outlook ir pieejams"""
    if not HAS_WIN32COM:
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
    except Exception:
        # Outlook nav instalēts vai nav atvērts — klusa iziešana
        return

    # Skenējam pēdējos 10 e-pastus
    items_to_save = {}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                items_to_save = json.load(f)
        except Exception:
            pass

    for i in range(1, min(11, messages.Count + 1)):
        try:
            msg = messages.Item(i)
            entry_id = msg.EntryID
            if entry_id in items_to_save:
                continue

            subj = msg.Subject or "(Bez temata)"
            sender = msg.SenderName or "Nezināms"
            body = msg.Body or ""

            # Ātrā filtrācija
            subj_lower = subj.lower()
            if any(k in subj_lower for k in ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums"]):
                items_to_save[entry_id] = {
                    "subject": subj,
                    "sender": sender,
                    "core_request": body[:120].strip() + "...",
                    "deadline": "Pārbaudīt e-pastā",
                    "category": "Cenu pieprasījums" if "cenu" in subj_lower else "Kritisks",
                    "status": "ACTIVE",
                    "received": msg.ReceivedTime.strftime("%d.%m %H:%M") if hasattr(msg, "ReceivedTime") else ""
                }
        except Exception:
            continue

    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(items_to_save, f, ensure_ascii=False, indent=2)

class OutlookRadarWatcher(threading.Thread):
    def __init__(self, interval_sec: int = 180):
        super().__init__(daemon=True)
        self.interval_sec = interval_sec

    def run(self):
        print("🛰️ [RADARS] Fona pastkastes skeneris startēts...")
        while True:
            try:
                scan_outlook_mailbox()
            except Exception:
                pass
            time.sleep(self.interval_sec)

# ==========================================
# 3. GRAFISKĀ SASKAIRNE (RADARA LOGS & TRAY)
# ==========================================
def open_email_in_outlook(entry_id):
    if not HAS_WIN32COM:
        print("⚠️ Šajā datorā nav instalēts Outlook Classic.")
        return
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.GetItemFromID(entry_id)
        msg.Display()
    except Exception as e:
        print(f"⚠️ Nevarēja atvērt e-pastu: {e}")

class RadarApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AQ Termiņu Radars")
        self.geometry("420x600")
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"420x600+{sw - 440}+{sh - 680}")
        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 6))

        ctk.CTkLabel(header, text="⚡ AQ Operētājsistēma", font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkButton(header, text="🔄", width=30, height=28, command=self.load_cards, fg_color="#1e293b").pack(side="right")

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=390, height=520)
        self.scroll_frame.pack(padx=10, pady=5, fill="both", expand=True)
        self.load_cards()

    def show_window(self):
        self.deiconify()
        self.lift()
        self.load_cards()

    def mark_as_done(self, entry_id, card_widget):
        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if entry_id in data:
                    data[entry_id]["status"] = "DONE"
                    with open(STATE_FILE, "w", encoding="utf-8") as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        card_widget.destroy()

    def load_cards(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()

        if not os.path.exists(STATE_FILE):
            ctk.CTkLabel(self.scroll_frame, text="Nav aktīvu datu.").pack(pady=30)
            return

        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}

        active_items = {k: v for k, v in data.items() if v.get("status") == "ACTIVE"}
        if not active_items:
            ctk.CTkLabel(self.scroll_frame, text="✨ Visi uzdevumi nokārtoti!", font=ctk.CTkFont(size=14)).pack(pady=40)
            return

        for entry_id, item_data in active_items.items():
            card = ctk.CTkFrame(self.scroll_frame, corner_radius=8, fg_color="#131b2e")
            card.pack(pady=6, padx=4, fill="x")

            cat = item_data.get("category", "Cits")
            badge_bg = "#ef4444" if cat in ["Līgums", "Kritisks"] else "#f59e0b" if "pieprasījums" in cat.lower() else "#3b82f6"

            top_bar = ctk.CTkFrame(card, fg_color="transparent")
            top_bar.pack(fill="x", padx=10, pady=(8, 2))
            ctk.CTkLabel(top_bar, text=f" {cat} ", fg_color=badge_bg, corner_radius=4, text_color="white", font=ctk.CTkFont(size=10, weight="bold")).pack(side="left")
            if item_data.get("received"):
                ctk.CTkLabel(top_bar, text=item_data.get("received"), text_color="#64748b", font=ctk.CTkFont(size=10)).pack(side="right")

            ctk.CTkLabel(card, text=item_data.get("subject", ""), font=ctk.CTkFont(size=12, weight="bold"), wraplength=350, justify="left").pack(anchor="w", padx=10, pady=(2, 0))
            ctk.CTkLabel(card, text=f"No: {item_data.get('sender', '')}", font=ctk.CTkFont(size=11), text_color="#94a3b8", wraplength=350, justify="left").pack(anchor="w", padx=10)

            req = item_data.get("core_request")
            if req:
                ctk.CTkLabel(card, text=req, font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=350, justify="left").pack(anchor="w", padx=10, pady=4)

            btn_row = ctk.CTkFrame(card, fg_color="transparent")
            btn_row.pack(fill="x", padx=10, pady=(4, 8))
            ctk.CTkButton(btn_row, text="↗ Atvērt", width=75, height=24, fg_color="#2563eb", hover_color="#1d4ed8", font=ctk.CTkFont(size=11),
                          command=lambda eid=entry_id: open_email_in_outlook(eid)).pack(side="left", padx=(0, 6))
            ctk.CTkButton(btn_row, text="✓ Nokārtots", width=85, height=24, fg_color="#059669", hover_color="#047857", font=ctk.CTkFont(size=11),
                          command=lambda eid=entry_id, c=card: self.mark_as_done(eid, c)).pack(side="left")

def run_tray(app):
    img = Image.new("RGBA", (64, 64), color=(0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((6, 6, 58, 58), fill="#f59e0b")

    def on_open(icon, itm):
        app.after(0, app.show_window)

    def on_quit(icon, itm):
        icon.stop()
        app.after(0, app.destroy)

    menu = pystray.Menu(item("Atvērt Radaru", on_open, default=True), item("Iziet", on_quit))
    icon = pystray.Icon("AQRCS", img, "AQ Termiņu Radars", menu)
    icon.run()

# ==========================================
# 4. GALVENAIS STARTS
# ==========================================
if __name__ == "__main__":
    print("🚀 [AQ-OS] Dzinējs startēts...")

    # A. Palaižam dokumentu uzraudzību
    digest_service = DocDigestService()
    watcher = FolderWatcherService(INBOX_DIR, digest_service)
    doc_thread = threading.Thread(target=watcher.start_watching, daemon=True)
    doc_thread.start()

    # B. Palaižam fona Outlook radaru
    radar_thread = OutlookRadarWatcher(interval_sec=300)
    radar_thread.start()

    # C. Palaižam GUI un System Tray
    app = RadarApp()
    tray_thread = threading.Thread(target=run_tray, args=(app,), daemon=True)
    tray_thread.start()

    app.mainloop()