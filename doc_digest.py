# -*- coding: utf-8 -*-
import os
import json
import shutil
import threading
from pathlib import Path
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog

import docx
from pptx import Presentation
import openpyxl
from openpyxl.styles import Font
import google.generativeai as genai

BASE_DIR = Path(__file__).resolve().parent
STORAGE_DIR = BASE_DIR / "storage"
INBOX_DIR = STORAGE_DIR / "inbox"
OUTBOX_DIR = STORAGE_DIR / "outbox"
ARCHIVE_DIR = STORAGE_DIR / "archive"
EXCEL_REGISTRY_PATH = STORAGE_DIR / "Ligumu_Registrs.xlsx"

for d in [INBOX_DIR, OUTBOX_DIR, ARCHIVE_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ==========================================
# PALĪGFUNKCIJAS UN EXCEL REĢISTRS
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
               "Dokumenta Datums", "Puses", "Summa", "Termiņš", "Statuss", "Riska Līmenis"] # Added Riska Līmenis
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
        
        risk_level = analysis.get("legal_analysis", {}).get("risk_level", "N/A") # Get risk level

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            file_name,
            str(analysis.get("document_type", "N/A")),
            str(analysis.get("document_number", "N/A")),
            str(analysis.get("document_date", "N/A")),
            parties_str,
            str(sum_val),
            str(term_val),
            status,
            risk_level # Added risk level to row
        ]
        ws.append(row)
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
        wb.save(EXCEL_REGISTRY_PATH)
    except Exception as e:
        print(f"⚠️ Kļūda Excel reģistrācijā: {e}")

# ==========================================
# DOKUMENTU APSTRĀDES PAKALPOJUMS
# ==========================================
class DocDigestService:
    def __init__(self, get_gemini_key_fn=None):
        self.get_gemini_key_fn = get_gemini_key_fn

    def _get_key(self):
        if self.get_gemini_key_fn:
            k = self.get_gemini_key_fn()
            if k:
                return k
        return os.environ.get("GEMINI_API_KEY") or os.environ.get("AQ_AI_API_KEY", "")

    def _read_docx(self, path: Path) -> str:
        doc = docx.Document(path)
        full_text = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if cells:
                    full_text.append(" | ".join(cells))
        return "\n".join(full_text)

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
        api_key = self._get_key()
        if not api_key:
            return {"error": "Nav norādīta Gemini API atslēga."}

        genai.configure(api_key=api_key)
        
        # Updated prompt text for Universal Legal Radar and Corrector
        prompt_text = """Analizējiet šo dokumentu rūpīgi. Jūsu uzdevums ir veikt "Universālo Juridisko Radaru un Korektoru" analīzi, identificējot "zemūdens akmeņus", juridiskos un gramatiskos trūkumus, kā arī sniegt skaidru, strukturētu atskaiti JSON formātā. Atbalstītās valodas analīzei ir latviešu, angļu un itāļu.

Īpaša uzmanība jāpievērš:
-   **Līgumu tipiem**: Īres līgumiem (īpaši Itālijas 'contratto di locazione' ar terminiem kā 'deposito cauzionale', 'spese condominiali', 'cedolare secca'), pakalpojumu un biznesa līgumiem. Sniedziet šo specifisko terminu īsu skaidrojumu, ja tie ir konstatēti.
-   **"Zemūdens akmeņiem"**: Automātiskiem pagarinājumiem, slēptajām izmaksām, sankcijām, līguma laušanas termiņiem un kārtību. Aprakstiet tos vienkāršā, cilvēcīgā valodā.
-   **Juridiskajiem un gramatiskajiem aspektiem**: Pārbaudiet visus datumus (brīdiniet par pagātnes datumiem, neloģiskiem termiņiem kā "beigas pirms sākuma"). Pārbaudiet gramatiku, pārrakstīšanās kļūdas, kā arī pušu rekvizītu vai vārdu neatbilstības tekstā.

Atgrieziet TIKAI derīgu JSON objektu, bez markdown blokiem vai papildu teksta, ievērojot šādu struktūru:

{
  "document_type": "Līgums / Rēķins / Akts / Specifikācija / Cits",
  "document_number": "Numurs vai N/A",
  "document_date": "Dokumenta galvenais datums YYYY-MM-DD vai N/A",
  "parties": [{"name": "Puses nosaukums", "role": "Izīrētājs/Īrnieks/Pasūtītājs/Piegādātājs/Cits"}],
  "subject": "Darījuma priekšmets (īss apraksts)",
  "financial_terms": {
    "total_amount": "Kopsumma ar valūtu vai N/A",
    "deadlines": "Būtiskākie apmaksas/izpildes termiņi vai N/A",
    "hidden_costs": [], // Saraksts ar slēptajām izmaksām vai negaidītiem maksājumiem, ja atrasti
    "penalties": []     // Saraksts ar soda sankcijām vai līgumsodiem, ja atrasti
  },
  "legal_analysis": {
    "risk_level": "Zaļš/Dzeltens/Sarkans", // Vispārējais riska līmenis, pamatojoties uz analīzi
    "critical_points": [], // Saraksts ar identificētiem kritiskiem punktiem vai "zemūdens akmeņiem"
    "automatic_extensions": [], // Saraksts ar automātiskajām pagarinājuma klauzulām
    "termination_clauses": [], // Saraksts ar līguma laušanas termiņiem un kārtību
    "italian_lease_specifics": { // Specifiski Itālijas īres līguma termini, ja attiecināms
      "deposito_cauzionale": "Paskaidrojums un summa, ja atrasts",
      "spese_condominiali": "Paskaidrojums un maksājumu kārtība, ja atrasts",
      "cedolare_secca": "Paskaidrojums, ja atrasts",
      "other_specifics": [] // Citi specifiski termini
    },
    "date_audit": {
      "all_dates_found": [], // Visi datumformāti, kas atrasti dokumentā
      "past_dates_warnings": [], // Brīdinājumi par pagātnes datumiem, kas varētu būt problemātiski
      "illogical_date_ranges": [] // Brīdinājumi par neloģiskiem datumu diapazoniem (piem., beigas pirms sākuma)
    },
    "grammatical_corrections": [], // Saraksts ar ieteiktajiem gramatikas/pārrakstīšanās labojumiem
    "party_name_inconsistencies": [] // Saraksts ar neatbilstībām pušu nosaukumos/rekvizītos
  },
  "suggested_corrections": [], // Kopējais saraksts ar ieteiktajiem labojumiem vai uzlabojumiem dokumentā
  "summary": "Īss un skaidrs kopsavilkums par dokumenta būtību un galvenajiem secinājumiem, cilvēcīgā valodā.",
  "risks_warnings": ["Šī sadaļa ir apkopojums, kas var pārklāties ar 'critical_points' un 'legal_analysis'. Tā kalpo kā ātra apkopojums par galvenajiem riskiem."]
}
"""

        try:
            model_name_to_use = "gemini-2.5-flash" # Primary model requested
            try:
                model = genai.GenerativeModel(
                    model_name=model_name_to_use,
                    generation_config={"response_mime_type": "application/json", "temperature": 0.2}
                )
            except Exception as e_primary:
                print(f"⚠️ Brīdinājums: Neizdevās ielādēt primāro modeli '{model_name_to_use}'. Mēģinu rezerves modeli 'gemini-2.0-flash'. Kļūda: {e_primary}")
                model_name_to_use = "gemini-2.0-flash" # Fallback model
                try:
                    model = genai.GenerativeModel(
                        model_name=model_name_to_use,
                        generation_config={"response_mime_type": "application/json", "temperature": 0.2}
                    )
                except Exception as e_fallback:
                    print(f"❌ Kļūda: Neizdevās ielādēt arī rezerves modeli '{model_name_to_use}'. Atgriežu kļūdu. Kļūda: {e_fallback}")
                    return {"error": f"Nav iespējams ielādēt nevienu no norādītajiem Gemini modeļiem. Pēdējā kļūda: {e_fallback}"}
            
            if file_type == "pdf":
                with open(str(content_payload), "rb") as f:
                    pdf_bytes = f.read()
                pdf_blob = {"mime_type": "application/pdf", "data": pdf_bytes}
                response = model.generate_content([prompt_text, pdf_blob])
            else:
                response = model.generate_content(f"{prompt_text}\n\nDokumenta saturs:\n{content_payload[:25000]}") # Limiting content for large files

            if response and response.text:
                return json.loads(response.text.strip())
            return {"error": "Tukša atbilde no MI"}
        except Exception as e:
            return {"error": str(e)}

    def _create_digest_docx(self, output_path: Path, data: dict, original_filename: str):
        doc = docx.Document()
        doc.add_heading(f"Kopsavilkums: {original_filename}", level=1)
        doc.add_paragraph(f"Izveidots: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        def add_sec(title, val, level=2):
            doc.add_heading(title, level=level)
            if isinstance(val, dict):
                if not val or all(not v for v in val.values()): # Check if dict is empty or all values are empty
                    doc.add_paragraph("N/A")
                    return
                for k, v in val.items():
                    if isinstance(v, list):
                        if v:
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}:")
                            for item in v:
                                doc.add_paragraph(f"  - {item}")
                        else:
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: N/A")
                    elif isinstance(v, dict):
                        if v:
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}:")
                            for sub_k, sub_v in v.items():
                                if sub_v:
                                    doc.add_paragraph(f"  - {sub_k.replace('_', ' ').capitalize()}: {sub_v}")
                                else:
                                    doc.add_paragraph(f"  - {sub_k.replace('_', ' ').capitalize()}: N/A")
                        else:
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: N/A")
                    else:
                        doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: {v if v else 'N/A'}")
            elif isinstance(val, list):
                if val and any(item for item in val if item and str(item).strip() != "N/A"): # Check if list is not empty and has meaningful content
                    for item in val:
                        doc.add_paragraph(f"• {item}")
                else:
                    doc.add_paragraph("N/A")
            else:
                doc.add_paragraph(str(val if val else "N/A"))
        
        # General Info
        add_sec("Vispārīgā informācija", {
            "Tips": data.get("document_type", "N/A"),
            "Numurs": data.get("document_number", "N/A"),
            "Datums": data.get("document_date", "N/A"),
            "Priekšmets": data.get("subject", "N/A")
        })

        # Parties
        parties_data = data.get("parties", [])
        if parties_data:
            doc.add_heading("Līguma Puses", level=2)
            for p in parties_data:
                doc.add_paragraph(f"• {p.get('name', 'N/A')} ({p.get('role', 'N/A')})")
        else:
            add_sec("Līguma Puses", "N/A")

        # Financial Terms
        fin_terms = data.get("financial_terms", {})
        add_sec("Finanšu noteikumi", {
            "Kopsumma": fin_terms.get("total_amount", "N/A"),
            "Termiņi": fin_terms.get("deadlines", "N/A")
        })
        add_sec("Slēptās izmaksas", fin_terms.get("hidden_costs", []))
        add_sec("Soda sankcijas", fin_terms.get("penalties", []))

        # Legal Analysis
        legal_analysis = data.get("legal_analysis", {})
        if legal_analysis and any(legal_analysis.values()):
            doc.add_heading("Juridiskā Analīze & Korektors", level=2)
            
            add_sec("Riska Līmenis", legal_analysis.get("risk_level", "N/A"), level=3)
            add_sec("Kritiskie Punkti", legal_analysis.get("critical_points", []), level=3)
            add_sec("Automātiskie pagarinājumi", legal_analysis.get("automatic_extensions", []), level=3)
            add_sec("Līguma laušanas noteikumi", legal_analysis.get("termination_clauses", []), level=3)

            italian_specs = legal_analysis.get("italian_lease_specifics", {})
            if italian_specs and any(v for v in italian_specs.values() if v):
                 add_sec("Itālijas Īres Līguma Specifika", italian_specs, level=3)
            
            date_audit = legal_analysis.get("date_audit", {})
            if date_audit and any(v for v in date_audit.values() if v):
                doc.add_heading("Datumu Audits", level=3)
                add_sec("Atrastie Datumi", date_audit.get("all_dates_found", []), level=4)
                add_sec("Brīdinājumi par pagātnes datumiem", date_audit.get("past_dates_warnings", []), level=4)
                add_sec("Neloģiski datumu diapazoni", date_audit.get("illogical_date_ranges", []), level=4)

            add_sec("Gramatikas labojumi", legal_analysis.get("grammatical_corrections", []), level=3)
            add_sec("Pušu nosaukumu neatbilstības", legal_analysis.get("party_name_inconsistencies", []), level=3)

        add_sec("Ieteiktie labojumi", data.get("suggested_corrections", []))
        add_sec("Kopsavilkums (cilvēcīgā valodā)", data.get("summary", "N/A"))

        # Existing risks_warnings (keeping for general summary of risks)
        add_sec("Vispārējie riski un brīdinājumi", data.get("risks_warnings", []))

        doc.save(output_path)

    def process_file(self, file_path: Path):
        ext = file_path.suffix.lower()
        if ext == ".docx":
            payload = self._read_docx(file_path)
        elif ext == ".txt":
            payload = self._read_txt(file_path)
        elif ext == ".pptx":
            payload = self._read_pptx(file_path)
        elif ext == ".pdf":
            payload = file_path
        else:
            return None, "Neatbalstīts formāts"

        if not payload:
            return None, "Neizdevās nolasīt faila saturu"

        analysis = self._analyze_with_gemini(payload, file_type=ext.replace(".", ""))
        if "error" in analysis:
            return None, analysis["error"]

        out_name = f"DIGEST_{file_path.stem}.docx"
        out_path = OUTBOX_DIR / out_name
        self._create_digest_docx(out_path, analysis, file_path.name)
        append_to_excel_registry(file_path.name, analysis)
        safe_archive_file(file_path, ARCHIVE_DIR)
        return analysis, str(out_path)

# ==========================================
# DOKUMENTU CILNE (CTkFrame)
# ==========================================
class DocDigestFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None):
        super().__init__(master, fg_color="transparent")
        self.service = DocDigestService(get_gemini_key_fn=get_gemini_key_fn)

        top_bar = ctk.CTkFrame(self, fg_color="#0f172a", height=50)
        top_bar.pack(fill="x", padx=10, pady=(10, 5))

        self.btn_select = ctk.CTkButton(
            top_bar,
            text="📥 Izvēlēties Dokumentu",
            width=170,
            height=34,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.select_and_process
        )
        self.btn_select.pack(side="left", padx=12, pady=8)

        self.btn_open_excel = ctk.CTkButton(
            top_bar,
            text="📊 Atvērt Līgumu Reģistru",
            width=170,
            height=34,
            fg_color="#059669",
            hover_color="#10b981",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.open_registry
        )
        self.btn_open_excel.pack(side="left", padx=6, pady=8)

        self.lbl_status = ctk.CTkLabel(top_bar, text="", font=ctk.CTkFont(size=12), text_color="#38bdf8")
        self.lbl_status.pack(side="right", padx=15)

        self.txt_output = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(size=13),
            fg_color="#0f172a",
            text_color="#f1f5f9",
            wrap="word"
        )
        self.txt_output.pack(fill="both", expand=True, padx=10, pady=10)
        self.txt_output.insert("end", "📄 Dokumentu Drop-Zone gatava.\nIelādējiet PDF, DOCX, PPTX vai TXT failu analīzei, Word kopsavilkuma izveidei un Excel reģistrēšanai.\n\n")

    def open_registry(self):
        if EXCEL_REGISTRY_PATH.exists():
            os.startfile(str(EXCEL_REGISTRY_PATH))
        else:
            self.lbl_status.configure(text="Reģistrs vēl nav izveidots")

    def _format_clean_report(self, filename: str, analysis: dict, docx_path: str) -> str:
        """Pārvērš JSON datus glītā, lasāmā cilvēka formātā."""
        lines = []
        lines.append("=" * 70)
        lines.append(f"⚡ AURAQUADROOS DOKUMENTA KOPSAVILKUMS: {filename}")
        lines.append("=" * 70)

        # General Info
        doc_type = analysis.get("document_type", "Dokuments")
        doc_nr = analysis.get("document_number", "Nav norādīts")
        doc_date = analysis.get("document_date", "Nav norādīts")
        lines.append(f"📌 Tips: {doc_type} | Nr: {doc_nr} | Datums: {doc_date}")

        subject = analysis.get("subject", "Nav norādīts")
        lines.append(f"🎯 Priekšmets: {subject}\n")

        # Parties
        parties = analysis.get("parties", [])
        if parties and isinstance(parties, list):
            lines.append("👥 LĪGUMA PUSES:")
            for p in parties:
                if isinstance(p, dict):
                    lines.append(f"   • {p.get('name', 'Nezināms')} ({p.get('role', 'Puse')})")
                else:
                    lines.append(f"   • {p}")
            lines.append("")

        # Financial Terms
        fin = analysis.get("financial_terms", {})
        if fin and any(v for v in fin.values() if v and (isinstance(v, list) and v or not isinstance(v, list))):
            lines.append("💵 FINANŠU NOSACĪJUMI & TERMIŅI:")
            if isinstance(fin, dict):
                if "total_amount" in fin and fin.get('total_amount'):
                    lines.append(f"   • Summa / Maksa: {fin.get('total_amount', 'N/A')}")
                if "deadlines" in fin and fin.get('deadlines'):
                    lines.append(f"   • Termiņi: {fin.get('deadlines', 'N/A')}")
                hidden_costs = fin.get("hidden_costs", [])
                if hidden_costs and any(hc.strip() != "N/A" for hc in hidden_costs):
                    lines.append("   ⚠️ Slēptās izmaksas:")
                    for cost in hidden_costs:
                        lines.append(f"     - {cost}")
                penalties = fin.get("penalties", [])
                if penalties and any(p.strip() != "N/A" for p in penalties):
                    lines.append("   🚫 Soda sankcijas:")
                    for penalty in penalties:
                        lines.append(f"     - {penalty}")
            lines.append("")

        # Legal Analysis
        legal_analysis = analysis.get("legal_analysis", {})
        if legal_analysis and any(v for v in legal_analysis.values() if v and (isinstance(v, dict) and any(sub_v for sub_v in v.values() if sub_v) or isinstance(v, list) and v or not isinstance(v, (dict, list)))):
            lines.append("🏛️ JURIDISKĀ ANALĪZE & KOREKTORS:")
            risk_level = legal_analysis.get("risk_level", "N/A").upper()
            if risk_level == "SARKANS":
                lines.append(f"   🔴 RISKA LĪMENIS: {risk_level} - Steidzama uzmanība!")
            elif risk_level == "DZELTENS":
                lines.append(f"   🟡 RISKA LĪMENIS: {risk_level} - Nepieciešama pārskatīšana.")
            elif risk_level == "ZAĻŠ":
                lines.append(f"   🟢 RISKA LĪMENIS: {risk_level} - Zems risks.")
            else:
                lines.append(f"   Riska līmenis: {risk_level}")

            critical_points = legal_analysis.get("critical_points", [])
            if critical_points and any(cp.strip() != "N/A" for cp in critical_points):
                lines.append("   🚨 Kritiskie Punkti & 'Zemūdens Akmeņi':")
                for cp in critical_points:
                    lines.append(f"     - {cp}")

            automatic_extensions = legal_analysis.get("automatic_extensions", [])
            if automatic_extensions and any(ae.strip() != "N/A" for ae in automatic_extensions):
                lines.append("   🔄 Automātiskie Pagarinājumi:")
                for ae in automatic_extensions:
                    lines.append(f"     - {ae}")

            termination_clauses = legal_analysis.get("termination_clauses", [])
            if termination_clauses and any(tc.strip() != "N/A" for tc in termination_clauses):
                lines.append("   ✂️ Līguma Laušanas Noteikumi:")
                for tc in termination_clauses:
                    lines.append(f"     - {tc}")

            italian_specs = legal_analysis.get("italian_lease_specifics", {})
            if italian_specs and any(v for v in italian_specs.values() if v):
                lines.append("   🇮🇹 Itālijas Īres Līguma Specifika:")
                for k, v in italian_specs.items():
                    if v and str(v).strip() != "N/A": lines.append(f"     • {k.replace('_', ' ').capitalize()}: {v}")

            date_audit = legal_analysis.get("date_audit", {})
            if date_audit and any(v for v in date_audit.values() if v):
                lines.append("   🗓️ Datumu Audits:")
                all_dates = date_audit.get("all_dates_found", [])
                if all_dates and any(d.strip() != "N/A" for d in all_dates):
                    lines.append(f"     • Atrastie datumi: {', '.join(all_dates)}")
                past_warnings = date_audit.get("past_dates_warnings", [])
                if past_warnings and any(w.strip() != "N/A" for w in past_warnings):
                    lines.append("     ❗ Pagātnes datumu brīdinājumi:")
                    for w in past_warnings:
                        lines.append(f"       - {w}")
                illogical_ranges = date_audit.get("illogical_date_ranges", [])
                if illogical_ranges and any(ir.strip() != "N/A" for ir in illogical_ranges):
                    lines.append("     ❌ Neloģiski datumu diapazoni:")
                    for ir in illogical_ranges:
                        lines.append(f"       - {ir}")

            gram_corrections = legal_analysis.get("grammatical_corrections", [])
            if gram_corrections and any(gc.strip() != "N/A" for gc in gram_corrections):
                lines.append("   📝 Gramatikas & Pārrakstīšanās Labojumi:")
                for gc in gram_corrections:
                    lines.append(f"     - {gc}")

            party_inconsistencies = legal_analysis.get("party_name_inconsistencies", [])
            if party_inconsistencies and any(pi.strip() != "N/A" for pi in party_inconsistencies):
                lines.append("   👤 Pušu Nosaukumu Neatbilstības:")
                for pi in party_inconsistencies:
                    lines.append(f"     - {pi}")
            lines.append("")

        # Suggested Corrections
        suggested_corrections = analysis.get("suggested_corrections", [])
        if suggested_corrections and any(sc.strip() != "N/A" for sc in suggested_corrections):
            lines.append("💡 IETEIKTIE LABOJUMI:")
            for sc in suggested_corrections:
                lines.append(f"   • {sc}")
            lines.append("")

        # Summary
        summary = analysis.get("summary", "Nav pieejams.")
        if summary.strip() != "N/A" and summary.strip() != "":
            lines.append("💬 KOPSAVILKUMS (cilvēcīgā valodā):")
            lines.append(f"   {summary}\n")

        # Existing risks_warnings (keeping for now, potentially can be deprecated if critical_points is sufficient)
        risks = analysis.get("risks_warnings", [])
        if risks and any(r for r in risks if r and r.strip() != "N/A" and r.strip() != ""): # Filter out "N/A"
            lines.append("⚠️ VISPĀRĒJI RISKI & BRĪDINĀJUMI:")
            if isinstance(risks, list):
                for r in risks:
                    if r and r.strip() != "N/A" and r.strip() != "":
                        lines.append(f"   🚩 {r}")
            else:
                if risks and risks.strip() != "N/A" and risks.strip() != "":
                    lines.append(f"   🚩 {risks}")
            lines.append("")


        lines.append("-" * 70)
        lines.append(f"💾 Word kopsavilkums saglabāts:\n   {docx_path}\n\n")
        return "\n".join(lines)

    def select_and_process(self):
        path = filedialog.askopenfilename(
            filetypes=[("Dokumenti", "*.pdf *.docx *.pptx *.txt"), ("Visi faili", "*.*")]
        )
        if not path:
            return

        p = Path(path)
        self.btn_select.configure(state="disabled")
        self.lbl_status.configure(text=f"Apstrādā {p.name}...")

        def _worker():
            analysis, res_msg = self.service.process_file(p)
            def _done():
                self.btn_select.configure(state="normal")
                if analysis:
                    self.lbl_status.configure(text="✅ Apstrāde pabeigta!")
                    clean_report = self._format_clean_report(p.name, analysis, res_msg)
                    self.txt_output.insert("end", clean_report)
                    self.txt_output.see("end")
                else:
                    self.lbl_status.configure(text="⚠️ Kļūda!")
                    self.txt_output.insert("end", f"\n❌ Kļūda failam {p.name}: {res_msg}\n\n")
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()