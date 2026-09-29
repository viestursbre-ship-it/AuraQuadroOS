import os
import time
import json
import threading
from pathlib import Path
from datetime import datetime

import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# --- Ceļi un mapes ---
STORAGE_DIR = Path("storage")
INBOX_DIR = STORAGE_DIR / "inbox"
OUTBOX_DIR = STORAGE_DIR / "outbox"
ARCHIVE_DIR = STORAGE_DIR / "archive"
EXCEL_REGISTRY_PATH = STORAGE_DIR / "Ligumu_Registrs.xlsx"

for _dir in [INBOX_DIR, OUTBOX_DIR, ARCHIVE_DIR]:
    _dir.mkdir(parents=True, exist_ok=True)

class DocDigestService:
    def __init__(self):
        print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Gatavs apstrādei!")

    def _read_docx(self, file_path: Path) -> str:
        doc = docx.Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text:
                    full_text.append(" | ".join(row_text))
        return "\n".join(full_text)

    def _read_txt(self, file_path: Path) -> str:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    def _analyze_with_gemini(self, content_payload) -> dict:
        api_key = os.getenv("AQ_AI_API_KEY")
        if not api_key:
            print("ERROR: AQ_AI_API_KEY nav iestatīts!")
            return {"error": "AQ_AI_API_KEY nav iestatīts."}

        prompt_text = """
        Analizējiet šo juridisko dokumentu un izvelciet galveno informāciju precīzā JSON formātā:
        {
          "document_type": "Līgums / Rēķins / Pielikums",
          "document_number": "Dokumenta numurs vai N/A",
          "document_date": "Dokumenta noslēgšanas datums vai N/A",
          "parties": [
            {"name": "Puses nosaukums", "role": "Pasūtītājs / Piegādātājs / Iznomātājs / Nomnieks"}
          ],
          "subject": "Darījuma priekšmeta īss, precīzs apraksts",
          "financial_terms": {
            "total_amount": "Kopējā summa ar valūtu un PVN statusu",
            "payment_schedule": "Apmaksas termiņš un kārtība",
            "deadlines": "Būtiskākie piegādes vai izpildes termiņi",
            "late_payment_penalties": "Kavējuma procenti vai līgumsods"
          },
          "ownership_termination": {
            "ownership_transfer": "Īpašumtiesību vai riska pārejas brīdis",
            "termination_conditions": "Līguma laušanas kārtība"
          },
          "risks_warnings": [
            "Būtiskākie riski, sankcijas, atbildības ierobežojumi vai brīdinājumi"
          ]
        }
        Atgrieziet TIKAI tīru JSON bez markdown blokiem.
        """

        try:
            client = genai.Client(api_key=api_key)

            if isinstance(content_payload, str):
                contents = [prompt_text, "\n\nDokumenta teksts:\n", content_payload]
            else:
                contents = [content_payload, prompt_text]

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=0.2,
                    response_mime_type="application/json"
                ),
            )

            if response and response.text:
                return json.loads(response.text.strip())
            return {"error": "Tukša AI atbilde"}

        except Exception as e:
            print(f"ERROR Gemini analīzē: {e}")
            return {"error": str(e)}

    def _create_digest_docx(self, output_path: Path, data: dict, original_filename: str):
        doc = docx.Document()

        doc.add_heading(f"Dokumenta Kopsavilkums: {original_filename}", level=1)
        doc.add_paragraph(f"Izveidots: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph("")

        def add_section(title: str, content):
            doc.add_heading(title, level=2)
            if isinstance(content, dict):
                for key, value in content.items():
                    doc.add_paragraph(f"• {key.replace('_', ' ').capitalize()}: {value}")
            elif isinstance(content, list):
                if content:
                    for item in content:
                        if isinstance(item, dict):
                            doc.add_paragraph(f"• {item.get('name', 'N/A')}: {item.get('role', 'N/A')}")
                        else:
                            doc.add_paragraph(f"• {item}")
                else:
                    doc.add_paragraph("Nav informācijas.")
            else:
                doc.add_paragraph(str(content))
            doc.add_paragraph("")

        add_section("Vispārīgā Informācija", {
            "Dokumenta tips": data.get("document_type", "N/A"),
            "Dokumenta numurs": data.get("document_number", "N/A"),
            "Dokumenta datums": data.get("document_date", "N/A")
        })

        add_section("Puses un Lomas", data.get("parties", []))
        add_section("Darījuma Priekšmets", data.get("subject", "N/A"))
        add_section("Finanšu Noteikumi", data.get("financial_terms", {}))
        add_section("Īpašumtiesības un Līguma Izbeigšana", data.get("ownership_termination", {}))
        add_section("Riski un Brīdinājumi", data.get("risks_warnings", []))

        doc.save(output_path)
        print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Izveidots kopsavilkums: {output_path.name}")

    def process_file(self, file_path: Path):
        print(f"[{datetime.now().strftime('%H:%M')}] Apstrādāju failu: {file_path.name}")
        ext = file_path.suffix.lower()

        try:
            if ext == ".docx":
                text = self._read_docx(file_path)
                if not text.strip():
                    return
                analysis = self._analyze_with_gemini(text)

            elif ext == ".txt":
                text = self._read_txt(file_path)
                if not text.strip():
                    return
                analysis = self._analyze_with_gemini(text)

            elif ext == ".pdf":
                with open(file_path, "rb") as f:
                    pdf_bytes = f.read()
                if not pdf_bytes:
                    return
                pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
                analysis = self._analyze_with_gemini(pdf_part)

            else:
                print(f"Neatbalstīts formāts: {file_path.name}")
                return

            if "error" in analysis:
                print(f"Kļūda AI analīzē: {analysis['error']}")
                return

            out_name = f"DIGEST_{file_path.stem}.docx"
            self._create_digest_docx(OUTBOX_DIR / out_name, analysis, file_path.name)
	    _append_to_excel_registry(file_path.name, analysis)

            archive_target = ARCHIVE_DIR / file_path.name
            file_path.rename(archive_target)
            print(f"Fails pārvietots uz arhīvu: {file_path.name}")

        except Exception as e:
            print(f"Neizdevās apstrādāt {file_path.name}: {e}")

# Šeit beidzas klases DocDigestService metode process_file:
        except Exception as e:
            print(f"Neizdevās apstrādāt {file_path.name}: {e}")

# ==========================================
# ŠEIT IELIEC 4. PUNKTA LEO FUNKCIJU:
# ==========================================
def _append_to_excel_registry(file_name, analysis, status="🟢 Apstrādāts"):
    headers = [
        "Ieraksta Datums", "Fails", "Dokumenta tips", "Dokumenta Nr.",
        "Dokumenta Datums", "Puses", "Summa", "Termiņš", "Statuss"
    ]

    if not EXCEL_REGISTRY_PATH.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Līgumu Reģistrs"
        ws.append(headers)
        for col_cell in ws[1]:
            col_cell.font = Font(bold=True)
        wb.save(EXCEL_REGISTRY_PATH)

    wb = load_workbook(EXCEL_REGISTRY_PATH)
    ws = wb.active

    parties_data = analysis.get("parties", [])
    if isinstance(parties_data, list):
        parties_str = ", ".join([p.get("name", str(p)) if isinstance(p, dict) else str(p) for p in parties_data])
    else:
        parties_str = str(parties_data)

    fin = analysis.get("financial_terms", {})
    sum_val = fin.get("total_amount", "N/A") if isinstance(fin, dict) else "N/A"
    term_val = fin.get("deadlines", "N/A") if isinstance(fin, dict) else "N/A"

    new_row_data = [
        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        file_name,
        analysis.get("document_type", "N/A"),
        analysis.get("document_number", "N/A"),
        analysis.get("document_date", "N/A"),
        parties_str,
        sum_val,
        term_val,
        status
    ]
    ws.append(new_row_data)
    wb.save(EXCEL_REGISTRY_PATH)
    print(f"[{datetime.now().strftime('%H:%M')}] Excel reģistrs atjaunināts: {file_name}")

# ==========================================
# Un tālāk turpinās klase FolderWatcherService:
# ==========================================
class FolderWatcherService:
    def __init__(self, watch_dir: Path, digest_service: DocDigestService):
        ...

class FolderWatcherService:
    def __init__(self, watch_dir: Path, digest_service: DocDigestService):
        self.watch_dir = watch_dir
        self.digest_service = digest_service
        self._running = False

    def start_watching(self, interval_sec: int = 4):
        self._running = True
        print(f"Uzraugu mapi: {self.watch_dir}")
        while self._running:
            for file_path in list(self.watch_dir.iterdir()):
                name = file_path.name
                if file_path.is_file() and not (name.startswith("~$") or name.startswith("DIGEST_")):
                    self.digest_service.process_file(file_path)
            time.sleep(interval_sec)

if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%H:%M')}] AQ-OS startēts! Andiamo!")
    digest = DocDigestService()
    watcher = FolderWatcherService(INBOX_DIR, digest)

    t = threading.Thread(target=watcher.start_watching, daemon=True)
    t.start()

    print(f"Sistēma gatava! Ievietojiet failus (docx, txt, pdf) mapē '{INBOX_DIR}'...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Sistēma apturēta.")