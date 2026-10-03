import os
import time
import json
import threading
from pathlib import Path
from datetime import datetime
import shutil # NEW: Import for file moving utilities

import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from pptx import Presentation

# --- Ceļi un mapes ---
STORAGE_DIR = Path("storage")
INBOX_DIR = STORAGE_DIR / "inbox"
OUTBOX_DIR = STORAGE_DIR / "outbox"
ARCHIVE_DIR = STORAGE_DIR / "archive"
EXCEL_REGISTRY_PATH = STORAGE_DIR / "Ligumu_Registrs.xlsx"

for _dir in [INBOX_DIR, OUTBOX_DIR, ARCHIVE_DIR]:
    _dir.mkdir(parents=True, exist_ok=True)

def _append_to_excel_registry(file_name, analysis, status="🟢 Apstrādāts"):
    headers = [
        "Ieraksta Datums", "Fails", "Dokumenta tips", "Dokumenta Nr.",
        "Dokumenta Datums", "Puses", "Summa", "Termiņš", "Statuss"
    ]

    try:
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
            parties_str = ", ".join([
                p.get("name", str(p)) if isinstance(p, dict) else str(p) 
                for p in parties_data
            ])
        else:
            parties_str = str(parties_data or "N/A")

        fin = analysis.get("financial_terms")
        if isinstance(fin, dict):
            sum_val = fin.get("total_amount", "N/A")
            term_val = fin.get("deadlines", "N/A")
        else:
            sum_val = "N/A"
            term_val = "N/A"

        new_row_data = [
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            file_name,
            str(analysis.get("document_type", "N/A")),
            str(analysis.get("document_number", "N/A")),
            str(analysis.get("document_date", "N/A")),
            parties_str,
            str(sum_val),
            str(term_val),
            status
        ]
        ws.append(new_row_data)

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(EXCEL_REGISTRY_PATH)
        print(f"[{datetime.now().strftime('%H:%M')}] Excel reģistrs atjaunināts: {file_name}")
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M')}] Kļūda rakstot Excel: {e}")

# NEW: Safe archive function to prevent duplicates and overwrites
def safe_archive_file(source_path, archive_dir):
    """
    Droši pārvieto failu uz arhīvu. Ja fails ar tādu nosaukumu jau eksistē,
    pievieno laika zīmogu (YYYYMMDD_HHMMSS), novēršot pārrakstīšanas un bloķēšanas cilpas.
    """
    os.makedirs(archive_dir, exist_ok=True)
    filename = os.path.basename(source_path)
    base_name, ext = os.path.splitext(filename)
    
    target_path = os.path.join(archive_dir, filename)
    
    # Ja fails jau eksistē mērķī, pieliekam laika zīmogu
    if os.path.exists(target_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_filename = f"{base_name}_{timestamp}{ext}"
        target_path = os.path.join(archive_dir, new_filename)
        
    try:
        shutil.move(source_path, target_path)
        print(f"[{datetime.now().strftime('%H:%M')}] 📦 Arhīvēts: {filename} -> {os.path.basename(target_path)}")
        return target_path
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M')}] ⚠️ Kļūda arhivējot failu {filename}: {e}")
        return None

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

    def _read_pptx(self, file_path: Path) -> str:
        prs = Presentation(file_path)
        full_text = []
        for i, slide in enumerate(prs.slides):
            full_text.append(f"\n--- Slaids {i+1} ---\n")
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.text.strip():
                        full_text.append(paragraph.text.strip())
            
            if slide.has_notes_slide:
                text_frame = slide.notes_slide.notes_text_frame
                if text_frame is not None:
                    for paragraph in text_frame.paragraphs:
                        if paragraph.text.strip():
                            full_text.append(f"Piezīmes: {paragraph.text.strip()}")
        return "\n".join(full_text)

    def _analyze_with_gemini(self, content_payload, file_type: str = ".docx") -> dict:
        api_key = os.getenv("AQ_AI_API_KEY")
        if not api_key:
            print("ERROR: AQ_AI_API_KEY nav iestatīts!")
            return {"error": "AQ_AI_API_KEY nav iestatīts."}

        if file_type == ".pptx":
            prompt_text = """
            Analizējiet šo prezentāciju un izvelciet galveno informāciju precīzā JSON formātā, pielāgojoties akadēmiskam kopsavilkumam:
            {
              "document_type": "Prezentācija / Lekcija",
              "document_number": "N/A",
              "document_date": "Prezentācijas datums vai N/A",
              "parties": [
                {"name": "Prezentācijas autors vai Pasniedzējs", "role": "Autors / Pasniedzējs"}
              ],
              "subject": "Prezentācijas galvenā tēma un mērķis",
              "financial_terms": {},
              "ownership_termination": {},
              "risks_warnings": [
                "Galvenās tēzes, kas izceltas prezentācijā",
                "Svarīgākie jēdzieni un definīcijas",
                "Kopsavilkums pa prezentācijas tēmām (katra tēma kā atsevišķs punkts ar īsu kopsavilkumu)",
                "Cita svarīga informācija vai secinājumi"
              ]
            }
            Centieties aizpildīt visus laukus, ja informācija ir pieejama.
            Atgrieziet TIKAI tīru JSON bez markdown blokiem.
            """
        else:
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
            print(f"[{datetime.now().strftime('%H:%M')}] ERROR Gemini analīzē: {e}")
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
                            if "name" in item and "role" in item:
                                doc.add_paragraph(f"• {item.get('name', 'N/A')}: {item.get('role', 'N/A')}")
                            elif "topic" in item and "summary" in item:
                                doc.add_paragraph(f"• Tēma: {item['topic']} - Kopsavilkums: {item['summary']}")
                            else:
                                doc.add_paragraph(f"• {str(item)}")
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
            analysis = None
            if ext == ".docx":
                text = self._read_docx(file_path)
                if not text.strip():
                    print(f"[{datetime.now().strftime('%H:%M')}] {file_path.name}: Tukšs dokuments, neapstrādāju.")
                    return
                analysis = self._analyze_with_gemini(text, file_type=ext)

            elif ext == ".txt":
                text = self._read_txt(file_path)
                if not text.strip():
                    print(f"[{datetime.now().strftime('%H:%M')}] {file_path.name}: Tukšs dokuments, neapstrādāju.")
                    return
                analysis = self._analyze_with_gemini(text, file_type=ext)

            elif ext == ".pdf":
                with open(file_path, "rb") as f:
                    pdf_bytes = f.read()
                if not pdf_bytes:
                    print(f"[{datetime.now().strftime('%H:%M')}] {file_path.name}: Tukšs dokuments, neapstrādāju.")
                    return
                pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
                analysis = self._analyze_with_gemini(pdf_part, file_type=ext)
            
            elif ext == ".pptx":
                text = self._read_pptx(file_path)
                if not text.strip():
                    print(f"[{datetime.now().strftime('%H:%M')}] {file_path.name}: Tukša prezentācija, neapstrādāju.")
                    return
                analysis = self._analyze_with_gemini(text, file_type=ext)

            else:
                print(f"[{datetime.now().strftime('%H:%M')}] Neatbalstīts faila formāts: {file_path.name}")
                return

            if analysis is None or "error" in analysis:
                print(f"[{datetime.now().strftime('%H:%M')}] Kļūda AI analīzē vai tukša atbilde {file_path.name}: {analysis.get('error', 'N/A')}")
                return

            out_name = f"DIGEST_{file_path.stem}.docx"
            self._create_digest_docx(OUTBOX_DIR / out_name, analysis, file_path.name)
            _append_to_excel_registry(file_path.name, analysis)

            # OLD: file_path.rename(ARCHIVE_DIR / file_path.name)
            safe_archive_file(file_path, ARCHIVE_DIR) # NEW: Use the safe archive function

        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M')}] Neizdevās apstrādāt {file_path.name}: {e}")

class FolderWatcherService:
    def __init__(self, watch_dir: Path, digest_service: DocDigestService):
        self.watch_dir = watch_dir
        self.digest_service = digest_service
        self._running = False

    def start_watching(self, interval_sec: int = 4):
        self._running = True
        print(f"[{datetime.now().strftime('%H:%M')}] Uzraugu mapi: {self.watch_dir}")
        while self._running:
            for file_path in list(self.watch_dir.iterdir()):
                name = file_path.name
                if file_path.is_file() and not (name.startswith("~$") or name.startswith("DIGEST_")):
                    self.digest_service.process_file(file_path)
            time.sleep(interval_sec)

if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%H:%M')}] AQ-OS startēts! Andiamo!")
    print('[AQ] Dzinējs pārbaudīts un gatavs darbam!')
    digest = DocDigestService()
    watcher = FolderWatcherService(INBOX_DIR, digest)

    t = threading.Thread(target=watcher.start_watching, daemon=True)
    t.start()

    print(f"Sistēma gatava! Ievietojiet failus (docx, txt, pdf, pptx) mapē '{INBOX_DIR}'...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print(f"[{datetime.now().strftime('%H:%M')}] Sistēma apturēta. Arrivederci!")