import os
import time
import json
import threading
from pathlib import Path
from datetime import datetime

import docx
from docx.shared import Inches
from google import genai # Jaunā bibliotēka

# --- AQ-OS Sistēmas Konfigurācija un Ceļi ---
STORAGE_DIR = Path("storage")
INBOX_DIR = STORAGE_DIR / "inbox"
OUTBOX_DIR = STORAGE_DIR / "outbox"
ARCHIVE_DIR = STORAGE_DIR / "archive" # Ja vajadzēs arhivēšanai

# Izveido nepieciešamās mapes, ja tās neeksistē
for _dir in [INBOX_DIR, OUTBOX_DIR, ARCHIVE_DIR]:
    _dir.mkdir(parents=True, exist_ok=True)

# --- Pakalpojumi (Modulārā Monolīta Iekšējās Dzinēja Daļas) ---

class DocDigestService:
    def __init__(self):
        # Šis simulē, ka DocDigestService ir "reģistrēts" ServiceRegistry
        # un izmanto EventEngine, bet vienā failā tas ir abstrakts.
        print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Maestro sagatavojies simfonijai!")

    def _read_docx(self, file_path: Path) -> str:
        """Nolasa .docx faila saturu, ieskaitot tabulas."""
        doc = docx.Document(file_path)
        full_text = []

        # Rindkopas
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())

        # Tabulas
        for table in doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_text.append(cell_text)
                if row_text:
                    full_text.append(" | ".join(row_text)) # Formats tabulas rindai

        return "\n".join(full_text)

    def _read_txt(self, file_path: Path) -> str:
        """Nolasa .txt faila saturu."""
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()

    def _analyze_with_ai(self, text: str) -> dict:
        """Analizē tekstu ar Google Gemini API un atgriež strukturētu JSON."""
        api_key = os.getenv("AQ_AI_API_KEY")
        if not api_key:
            print(f"[{datetime.now().strftime('%H:%M')}] Leo error: AQ_AI_API_KEY nav iestatīts!")
            return {"error": "AQ_AI_API_KEY nav iestatīts."}

        try:
            from google import genai
            from google.genai import types

            client = genai.Client(api_key=api_key)

            prompt_text = f"""
            Analizējiet šo dokumenta tekstu un izvelciet galveno informāciju strukturētā JSON formātā.
            JSON objektam ir jāiekļauj šādi lauki:
            "document_type": (piemēram, "Līgums", "Rēķins", "Protokols", "Pielikums", "Atskaite")
            "document_number": (piemēram, "LV2023/123", "N/A", ja nav atrasts)
            "document_date": (piemēram, "2023-10-26", "N/A", ja nav atrasts)
            "parties": [
                {{"name": "Puses A Nosaukums", "role": "Iznomātājs / Piegādātājs"}},
                {{"name": "Puses B Nosaukums", "role": "Nomnieks / Pasūtītājs"}}
            ]
            "subject": "Īss darījuma priekšmeta kopsavilkums."
            "financial_terms": {{
                "total_amount": "Kopējā summa ar valūtu un PVN",
                "payment_schedule": "Maksājumu grafika apraksts",
                "deadlines": "Galvenie termiņi",
                "late_payment_penalties": "Kavējuma procenti vai līgumsods"
            }}
            "ownership_termination": {{
                "ownership_transfer": "Īpašumtiesību pārejas nosacījumi",
                "termination_conditions": "Līguma izbeigšanas nosacījumi"
            }}
            "risks_warnings": [
                "Būtiskākie riski vai sankcijas"
            ]

            Dokumenta teksts:
            ---
            {text}
            ---

            Izvadei jābūt TIKAI tīram JSON objektam bez markdown blokiem.
            """

            # Pareizais jaunais izsaukums:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt_text,
                config=types.GenerateContentConfig(
                    temperature=0.2,
                    response_mime_type="application/json"
                ),
            )

            if response and response.text:
                return json.loads(response.text)
            else:
                return {"error": "Nav teksta AI atbildē"}

        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M')}] Leo error: AI analīze neizdevās: {e}")
            return {"error": str(e)}

    def _create_digest_docx(self, output_path: Path, data: dict, original_filename: str):
        """Izveido noformētu .docx kopsavilkuma failu."""
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
        print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Izveidots kopsavilkuma fails: {output_path.name}")

    def process_file(self, file_path: Path):
        """Apstrādā vienu failu: nolasa, analizē, izveido DOCX kopsavilkumu."""
        print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Sāku apstrādāt failu: {file_path.name}")
        try:
            if file_path.suffix == '.docx':
                text_content = self._read_docx(file_path)
            elif file_path.suffix == '.txt':
                text_content = self._read_txt(file_path)
            else:
                print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Neatbalstīts faila formāts: {file_path.name}")
                return

            if not text_content.strip():
                print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Fails {file_path.name} ir tukšs vai nesatur nolasāmu tekstu. Ignorēju.")
                return

            analysis_result = self._analyze_with_ai(text_content)

            if "error" in analysis_result:
                print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Kļūda AI analīzē failam {file_path.name}: {analysis_result['error']}")
                return

            original_filename_no_suffix = file_path.stem
            digest_filename = f"DIGEST_{original_filename_no_suffix}.docx"
            output_path = OUTBOX_DIR / digest_filename
            self._create_digest_docx(output_path, analysis_result, file_path.name)

            # Pēc apstrādes failu pārvieto uz arhīvu
            archive_path = ARCHIVE_DIR / file_path.name
            file_path.rename(archive_path)
            print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Fails {file_path.name} pārvietots uz arhīvu: {archive_path.name}")

        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M')}] DocDigestService: Neizdevās apstrādāt failu {file_path.name}: {e}")

class FolderWatcherService:
    def __init__(self, watch_dir: Path, digest_service: DocDigestService):
        self.watch_dir = watch_dir
        self.digest_service = digest_service
        self._running = False
        print(f"[{datetime.now().strftime('%H:%M')}] FolderWatcherService: Uzraugu mapi: {self.watch_dir}")

    def _filter_files(self, file_path: Path) -> bool:
        """Filtrē failus, kas jāignorē (Word slēdzenes, DIGEST_ faili)."""
        filename = file_path.name
        if filename.startswith("~$") or filename.startswith("DIGEST_"):
            return False
        return True

    def start_watching(self, interval_sec: int = 5):
        """Sāk mapes uzraudzību."""
        self._running = True
        while self._running:
            print(f"[{datetime.now().strftime('%H:%M')}] FolderWatcherService: Pārbaudu mapi {self.watch_dir}...")
            for file_path in self.watch_dir.iterdir():
                if file_path.is_file() and self._filter_files(file_path):
                    print(f"[{datetime.now().strftime('%H:%M')}] FolderWatcherService: Atrasts jauns fails: {file_path.name}")
                    self.digest_service.process_file(file_path)
            time.sleep(interval_sec)

    def stop_watching(self):
        """Pārtrauc mapes uzraudzību."""
        self._running = False
        print(f"[{datetime.now().strftime('%H:%M')}] FolderWatcherService: Mapes uzraudzība apturēta.")

# --- Galvenā Lietojumprogrammas Loģika ---
if __name__ == "__main__":
    print(f"[{datetime.now().strftime('%H:%M')}] AQ-OS Modulārais Monolīts startē! Andiamo!")

    # Šajā fāzē GatewayRouter, EventEngine un ServiceRegistry ir abstrakcijas,
    # kas nozīmē, ka pakalpojumi tiek tieši instancēti un saistīti,
    # simulējot iekšējo komunikāciju.
    doc_digest_service = DocDigestService()
    folder_watcher = FolderWatcherService(watch_dir=INBOX_DIR, digest_service=doc_digest_service)

    # Palaist FolderWatcherService atsevišķā pavedienā, lai galvenais pavediens nebūtu bloķēts
    watcher_thread = threading.Thread(target=folder_watcher.start_watching, daemon=True)
    watcher_thread.start()

    print(f"[{datetime.now().strftime('%H:%M')}] Sistēma ir gatava! Novietojiet .docx vai .txt failus mapē '{INBOX_DIR}' un skatieties maģiju mapē '{OUTBOX_DIR}'!")
    print(f"[{datetime.now().strftime('%H:%M')}] Lai apturētu, nospiediet Ctrl+C.")

    try:
        while True:
            time.sleep(1) # Turam galveno pavedienu dzīvu
    except KeyboardInterrupt:
        print(f"[{datetime.now().strftime('%H:%M')}] Sistēma tiek apturēta. Arrivederci!")
        folder_watcher.stop_watching()