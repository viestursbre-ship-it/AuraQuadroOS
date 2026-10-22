# -*- coding: utf-8 -*-
import os
import json
import shutil
import threading
from pathlib import Path
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog
import subprocess 
import sys 
import time 

import docx
from pptx import Presentation
import openpyxl
from openpyxl.styles import Font
import google.generativeai as genai

import i18n # <--- ADDED: Import local i18n module

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

def append_to_excel_registry(file_name: str, analysis: dict, status: str = "Apstrādāts"):
    headers = [i18n.t("excel_header_record_date"), i18n.t("excel_header_file"), 
               i18n.t("excel_header_doc_type"), i18n.t("excel_header_doc_number"), 
               i18n.t("excel_header_doc_date"), i18n.t("excel_header_parties"), 
               i18n.t("excel_header_amount"), i18n.t("excel_header_deadline"), 
               i18n.t("excel_header_status"), i18n.t("excel_header_risk_level")] 
    try:
        if not EXCEL_REGISTRY_PATH.exists():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = i18n.t("excel_sheet_title")
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
        
        risk_level = analysis.get("legal_analysis", {}).get("risk_level", "N/A")

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
            risk_level
        ]
        ws.append(row)
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
        wb.save(EXCEL_REGISTRY_PATH)
    except Exception as e:
        print(f"⚠️ {i18n.t('error_excel_registry')}: {e}")

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
            full_text.append(f"\n--- {i18n.t('slide')} {i} ---")
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            full_text.append(paragraph.text.strip())
        return "\n".join(full_text)

    def _analyze_with_gemini(self, content_payload, file_type: str = "docx") -> dict:
        api_key = self._get_key()
        if not api_key:
            return {"error": i18n.t("error_gemini_api_key_not_provided")}

        genai.configure(api_key=api_key)
        
        prompt_text = """Analizējiet šo dokumentu rūpīgi. Jūsu uzdevums ir veikt "Universālo Juridisko Radaru un Korektoru" analīzi, identificējot "zemūdens akmeņus", juridiskos un gramatiskos trūkumus, kā arī sniegt skaidru, strukturētu atskaiti JSON formātā. Atbalstītās valodas analīzei ir latviešu, angļu, itāļu, ukraiņu un krievu.

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
    "italian_lease_specifics": { // This section should ONLY be populated if the document is explicitly identified as an Italian property lease. Otherwise, leave all its fields as 'N/A' or empty strings.
      "deposito_cauzionale": "Paskaidrojums un summa, ja atrasts (N/A ja nav attiecināms)",
      "spese_condominiali": "Paskaidrojums un maksājumu kārtība, ja atrasts (N/A ja nav attiecināms)",
      "cedolare_secca": "Paskaidrojums, ja atrasts (N/A ja nav attiecināms)",
      "other_specifics": [] // Citi specifiski termini (atstāt tukšu, ja nav attiecināms)
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
            model_name_to_use = "gemini-2.5-flash"
            try:
                model = genai.GenerativeModel(
                    model_name=model_name_to_use,
                    generation_config={"response_mime_type": "application/json", "temperature": 0.2}
                )
            except Exception as e_primary:
                print(f"⚠️ {i18n.t('warning_primary_model_failed', model=model_name_to_use)} 'gemini-2.0-flash'. {i18n.t('error_msg')}: {e_primary}")
                model_name_to_use = "gemini-2.0-flash"
                try:
                    model = genai.GenerativeModel(
                        model_name=model_name_to_use,
                        generation_config={"response_mime_type": "application/json", "temperature": 0.2}
                    )
                except Exception as e_fallback:
                    print(f"❌ {i18n.t('error_fallback_model_failed', model=model_name_to_use)}. {i18n.t('error_msg')}: {e_fallback}")
                    return {"error": f"{i18n.t('error_no_gemini_model_loaded')}. {i18n.t('error_last_error')}: {e_fallback}"}
            
            if file_type == "pdf":
                with open(str(content_payload), "rb") as f:
                    pdf_bytes = f.read()
                pdf_blob = {"mime_type": "application/pdf", "data": pdf_bytes}
                response = model.generate_content([prompt_text, pdf_blob])
            else:
                response = model.generate_content(f"{prompt_text}\n\n{i18n.t('document_content')}:\n{content_payload[:25000]}")

            if response and response.text:
                return json.loads(response.text.strip())
            return {"error": i18n.t("error_empty_ai_response")}
        except Exception as e:
            return {"error": str(e)}

    def _create_digest_docx(self, output_path_base: Path, data: dict, original_filename: str) -> Path:
        doc = docx.Document()
        doc.add_heading(f"{i18n.t('summary_title')}: {original_filename}", level=1)
        doc.add_paragraph(f"{i18n.t('created_at')}: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        def add_sec(title, val, level=2):
            doc.add_heading(title, level=level)
            if isinstance(val, dict):
                if not val or all(
                    (isinstance(v, (list, dict)) and not v) or
                    (isinstance(v, str) and v.strip().upper() == "N/A") or
                    (v is None) for v in val.values()
                ):
                    doc.add_paragraph(i18n.t("not_available_short"))
                    return
                for k, v in val.items():
                    if isinstance(v, list):
                        if v and any(item for item in v if item and str(item).strip().upper() != "N/A"):
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}:")
                            for item in v:
                                if item and str(item).strip().upper() != "N/A":
                                    doc.add_paragraph(f"  - {item}")
                        else:
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: {i18n.t('not_available_short')}")
                    elif isinstance(v, dict):
                        if v and any(sub_v for sub_v in v.values() if sub_v and str(sub_v).strip().upper() != "N/A"):
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}:")
                            for sub_k, sub_v in v.items():
                                if sub_v and str(sub_v).strip().upper() != "N/A":
                                    doc.add_paragraph(f"  - {sub_k.replace('_', ' ').capitalize()}: {sub_v}")
                                else:
                                    doc.add_paragraph(f"  - {sub_k.replace('_', ' ').capitalize()}: {i18n.t('not_available_short')}")
                        else:
                            doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: {i18n.t('not_available_short')}")
                    else:
                        doc.add_paragraph(f"• {k.replace('_', ' ').capitalize()}: {v if v and str(v).strip().upper() != 'N/A' else i18n.t('not_available_short')}")
            elif isinstance(val, list):
                if val and any(item for item in val if item and str(item).strip().upper() != "N/A"):
                    for item in val:
                        doc.add_paragraph(f"• {item}")
                else:
                    doc.add_paragraph(i18n.t("not_available_short"))
            else:
                doc.add_paragraph(str(val if val and str(val).strip().upper() != "N/A" else i18n.t("not_available_short")))
        
        add_sec(i18n.t("general_info_title"), {
            i18n.t("doc_type_label"): data.get("document_type", "N/A"),
            i18n.t("doc_number_label"): data.get("document_number", "N/A"),
            i18n.t("doc_date_label"): data.get("document_date", "N/A"),
            i18n.t("subject_label"): data.get("subject", "N/A")
        })

        parties_data = data.get("parties", [])
        if parties_data:
            doc.add_heading(i18n.t("parties_title"), level=2)
            for p_item in parties_data:
                doc.add_paragraph(f"• {p_item.get('name', i18n.t('unknown'))} ({p_item.get('role', i18n.t('party'))})")
        else:
            add_sec(i18n.t("parties_title"), i18n.t("not_available_short"))

        fin_terms = data.get("financial_terms", {})
        add_sec(i18n.t("financial_terms_title"), {
            i18n.t("total_amount_label"): fin_terms.get("total_amount", "N/A"),
            i18n.t("deadlines_label"): fin_terms.get("deadlines", "N/A")
        })
        add_sec(i18n.t("hidden_costs_title"), fin_terms.get("hidden_costs", []))
        add_sec(i18n.t("penalties_title"), fin_terms.get("penalties", []))

        legal_analysis = data.get("legal_analysis", {})
        if legal_analysis and any(
            (isinstance(v, (list, dict)) and v and any(sub_v for sub_v in (v.values() if isinstance(v, dict) else v) if sub_v and str(sub_v).strip().upper() != "N/A")) or
            (isinstance(v, str) and v.strip().upper() != "N/A") or
            (v is not None and not isinstance(v, (list, dict)))
            for v in legal_analysis.values()
        ):
            # Update: Replace "legal_analysis_corrector_title" with direct localized string
            doc.add_heading("Juridiskā risku analīze", level=2) # Replaced with direct localized string
            
            add_sec(i18n.t("risk_level_title"), legal_analysis.get("risk_level", "N/A"), level=3)
            add_sec(i18n.t("critical_points_title"), legal_analysis.get("critical_points", []), level=3)
            add_sec(i18n.t("automatic_extensions_title"), legal_analysis.get("automatic_extensions", []), level=3)
            add_sec(i18n.t("termination_clauses_title"), legal_analysis.get("termination_clauses", []), level=3)

            italian_specs = legal_analysis.get("italian_lease_specifics", {})
            if italian_specs and any(v for v in italian_specs.values() if v and str(v).strip().upper() != "N/A"):
                 add_sec(i18n.t("italian_lease_specifics_title"), italian_specs, level=3)
            
            date_audit = legal_analysis.get("date_audit", {})
            if date_audit and any(v for v in date_audit.values() if v and (isinstance(v, list) and v or not isinstance(v, list))):
                doc.add_heading(i18n.t("date_audit_title"), level=3)
                add_sec(i18n.t("found_dates_title"), date_audit.get("all_dates_found", []), level=4)
                add_sec(i18n.t("past_date_warnings_title"), date_audit.get("past_dates_warnings", []), level=4)
                add_sec(i18n.t("illogical_date_ranges_title"), date_audit.get("illogical_date_ranges", []), level=4)

            add_sec(i18n.t("grammatical_corrections_title"), legal_analysis.get("grammatical_corrections", []), level=3)
            add_sec(i18n.t("party_name_inconsistencies_title"), legal_analysis.get("party_name_inconsistencies", []), level=3)

        # Update: Replace with direct localized strings
        add_sec("Ieteiktie labojumi un korekcijas", data.get("suggested_corrections", []))
        add_sec("Dokumenta kopsavilkums saprotamā valodā", data.get("summary", "N/A"))

        # Update: Replace with direct localized strings
        add_sec("Galvenie riski un brīdinājumi", data.get("risks_warnings", []))

        current_output_path = output_path_base
        attempt_count = 0
        max_attempts = 10

        while attempt_count < max_attempts:
            try:
                if current_output_path.exists() and attempt_count == 0:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                    current_output_path = output_path_base.parent / f"{output_path_base.stem}_{timestamp}{output_path_base.suffix}"
                
                doc.save(str(current_output_path))
                return current_output_path
            except PermissionError:
                attempt_count += 1
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                current_output_path = output_path_base.parent / f"{output_path_base.stem}_CONFLICT_{timestamp}{output_path_base.suffix}"
                print(f"⚠️ {i18n.t('error_saving_word_permission')}. {i18n.t('retrying_with_unique_name')}: {current_output_path.name}")
                time.sleep(0.5)
            except Exception as e:
                raise RuntimeError(f"{i18n.t('error_saving_word_doc', filename=current_output_path.name)}: {e}")
        
        raise RuntimeError(f"{i18n.t('error_failed_to_save_word_summary', attempts=max_attempts)}")
    

# ==========================================
# DOKUMENTU CILNE (CTkFrame)
# ==========================================
class DocDigestFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None):
        super().__init__(master, fg_color="transparent")
        self.service = DocDigestService(get_gemini_key_fn=get_gemini_key_fn)
        self.last_summary_path = None

        top_bar = ctk.CTkFrame(self, fg_color="#0f172a", height=50)
        top_bar.pack(fill="x", padx=10, pady=(10, 5))

        self.btn_select = ctk.CTkButton(
            top_bar,
            text=i18n.t("btn_select_doc"),
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
            text=i18n.t("btn_open_contracts"),
            width=170,
            height=34,
            fg_color="#059669",
            hover_color="#10b981",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.open_registry
        )
        self.btn_open_excel.pack(side="left", padx=6, pady=8)

        self.btn_open_summary = ctk.CTkButton(
            top_bar,
            text=i18n.t("btn_open_conclusion"),
            width=170,
            height=34,
            fg_color="#8b5cf6",
            hover_color="#7c3aed",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.open_summary_docx,
            state="disabled"
        )
        self.btn_open_summary.pack(side="left", padx=6, pady=8)

        self.progress_bar = ctk.CTkProgressBar(
            top_bar,
            orientation="horizontal",
            mode="determinate",
            width=200,
            height=20,
            fg_color="#475569",
            progress_color="#60a5fa"
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=(10, 15), pady=8)

        self.lbl_status = ctk.CTkLabel(top_bar, text=f"0% - {i18n.t('status_waiting_doc')}", font=ctk.CTkFont(size=12), text_color="#38bdf8")
        self.lbl_status.pack(side="right", padx=(0, 10), pady=8)

        self.txt_output = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(size=13),
            fg_color="#0f172a",
            text_color="#f1f5f9",
            wrap="word"
        )
        self.txt_output.pack(fill="both", expand=True, padx=10, pady=10)
        self.txt_output.insert("end", i18n.t("txt_output_initial_message"))

    # <--- ADDED: set_ui_language and refresh_ui methods
    def set_ui_language(self, lang: str = None):
        """Sets the UI language and refreshes all translatable elements."""
        self.refresh_ui()

    def refresh_ui(self):
        """Refreshes the text of UI elements according to the current language."""
        if hasattr(self, "btn_select"):
            self.btn_select.configure(text=i18n.t("btn_select_doc"))
        if hasattr(self, "btn_open_excel"):
            self.btn_open_excel.configure(text=i18n.t("btn_open_contracts"))
        if hasattr(self, "btn_open_summary"):
            self.btn_open_summary.configure(text=i18n.t("btn_open_conclusion"))
        # lbl_status initial message part is updated in __init__ and then dynamically.
        # No specific instruction to configure lbl_status text in refresh_ui directly.
    # ADDED END ---

    def open_registry(self):
        if EXCEL_REGISTRY_PATH.exists():
            try:
                if sys.platform == "win32":
                    os.startfile(str(EXCEL_REGISTRY_PATH))
                elif sys.platform == "darwin":
                    subprocess.run(['open', str(EXCEL_REGISTRY_PATH)], check=True)
                else:
                    subprocess.run(['xdg-open', str(EXCEL_REGISTRY_PATH)], check=True)
                self._update_gui_progress(1.0, i18n.t("status_registry_opened"))
            except Exception as e:
                self._update_gui_progress(0.0, f"{i18n.t('error_opening_registry')}: {e}")
        else:
            self._update_gui_progress(0.0, i18n.t("status_registry_not_created"))

    def open_summary_docx(self):
        """Opens the last created Word summary document using the default application."""
        if self.last_summary_path and self.last_summary_path.exists():
            try:
                if sys.platform == "win32":
                    os.startfile(str(self.last_summary_path))
                elif sys.platform == "darwin":
                    subprocess.run(['open', str(self.last_summary_path)], check=True)
                else:
                    subprocess.run(['xdg-open', str(self.last_summary_path)], check=True)
                self._update_gui_progress(1.0, i18n.t("status_word_summary_opened"))
            except Exception as e:
                self._update_gui_progress(0.0, f"{i18n.t('error_opening_word_summary')}: {e}")
        else:
            self._update_gui_progress(0.0, i18n.t("status_no_word_summary_available"))

    def _update_gui_progress(self, progress_value: float, message: str):
        """Updates the progress bar and status label with percentage and message."""
        percentage = int(progress_value * 100)
        self.progress_bar.set(progress_value)
        self.lbl_status.configure(text=f"{percentage}% - {message}")

    def _format_clean_report(self, filename: str, analysis: dict, docx_path: str) -> str:
        """Pārvērš JSON datus glītā, lasāmā cilvēka formātā."""
        lines = []
        lines.append("=" * 70)
        lines.append(f"⚡ {i18n.t('report_title')}: {filename}")
        lines.append("=" * 70)

        doc_type = analysis.get("document_type", i18n.t("doc_type_default"))
        doc_nr = analysis.get("document_number", i18n.t("not_specified"))
        doc_date = analysis.get("document_date", i18n.t("not_specified"))
        lines.append(f"📌 {i18n.t('report_doc_type')}: {doc_type} | {i18n.t('report_doc_nr')}: {doc_nr} | {i18n.t('report_doc_date')}: {doc_date}")

        subject = analysis.get("subject", i18n.t("not_specified"))
        lines.append(f"🎯 {i18n.t('report_subject')}: {subject}\n")

        parties = analysis.get("parties", [])
        if parties and isinstance(parties, list):
            lines.append(i18n.t("report_parties_title").upper() + ":")
            for p_item in parties:
                if isinstance(p_item, dict):
                    lines.append(f"   • {p_item.get('name', i18n.t('unknown'))} ({p_item.get('role', i18n.t('party'))})")
                else:
                    lines.append(f"   • {p_item}")
            lines.append("")

        fin = analysis.get("financial_terms", {})
        if fin and any(v for v in fin.values() if v and (isinstance(v, list) and v or not isinstance(v, list))):
            lines.append(i18n.t("report_financial_terms_title").upper() + ":")
            if isinstance(fin, dict):
                if "total_amount" in fin and fin.get('total_amount') and str(fin.get('total_amount')).strip().upper() != "N/A":
                    lines.append(f"   • {i18n.t('report_total_amount')}: {fin.get('total_amount', i18n.t('not_available_short'))}")
                if "deadlines" in fin and fin.get('deadlines') and str(fin.get('deadlines')).strip().upper() != "N/A":
                    lines.append(f"   • {i18n.t('report_deadlines')}: {fin.get('deadlines', i18n.t('not_available_short'))}")
                hidden_costs = fin.get("hidden_costs", [])
                if hidden_costs and any(hc.strip().upper() != "N/A" for hc in hidden_costs):
                    lines.append(f"   ⚠️ {i18n.t('report_hidden_costs')}:")
                    for cost in hidden_costs:
                        lines.append(f"     - {cost}")
                penalties = fin.get("penalties", [])
                if penalties and any(p_item.strip().upper() != "N/A" for p_item in penalties):
                    lines.append(f"   🚫 {i18n.t('report_penalties')}:")
                    for penalty in penalties:
                        lines.append(f"     - {penalty}")
            lines.append("")

        legal_analysis = analysis.get("legal_analysis", {})
        if legal_analysis and any(v for v in legal_analysis.values() if v and (isinstance(v, dict) and any(sub_v for sub_v in v.values() if sub_v and str(sub_v).strip().upper() != "N/A") or isinstance(v, list) and v and any(item.strip().upper() != "N/A" for item in v) or (isinstance(v, str) and v.strip().upper() != "N/A"))):
            # Update: Replaced with direct localized string
            lines.append("JURIDISKĀ RISKU ANALĪZE")
            risk_level = legal_analysis.get("risk_level", "N/A").upper()
            if risk_level == i18n.t("risk_red").upper():
                lines.append(f"   🔴 {i18n.t('report_risk_level')}: {risk_level} - {i18n.t('risk_red_warning')}!")
            elif risk_level == i18n.t("risk_yellow").upper():
                lines.append(f"   🟡 {i18n.t('report_risk_level')}: {risk_level} - {i18n.t('risk_yellow_warning')}.")
            elif risk_level == i18n.t("risk_green").upper():
                lines.append(f"   🟢 {i18n.t('report_risk_level')}: {risk_level} - {i18n.t('risk_green_info')}.")
            else:
                lines.append(f"   {i18n.t('report_risk_level')}: {risk_level}")

            critical_points = legal_analysis.get("critical_points", [])
            if critical_points and any(cp.strip().upper() != "N/A" for cp in critical_points):
                lines.append(f"   🚨 {i18n.t('report_critical_points')}:")
                for cp in critical_points:
                    lines.append(f"     - {cp}")

            automatic_extensions = legal_analysis.get("automatic_extensions", [])
            if automatic_extensions and any(ae.strip().upper() != "N/A" for ae in automatic_extensions):
                lines.append(f"   🔄 {i18n.t('report_automatic_extensions')}:")
                for ae in automatic_extensions:
                    lines.append(f"     - {ae}")

            termination_clauses = legal_analysis.get("termination_clauses", [])
            if termination_clauses and any(tc.strip().upper() != "N/A" for tc in termination_clauses):
                lines.append(f"   ✂️ {i18n.t('report_termination_clauses')}:")
                for tc in termination_clauses:
                    lines.append(f"     - {tc}")

            italian_specs = legal_analysis.get("italian_lease_specifics", {})
            if italian_specs and any(v for v in italian_specs.values() if v and str(v).strip().upper() != "N/A"):
                lines.append(f"   🇮🇹 {i18n.t('report_italian_lease_specifics')}:")
                for k, v in italian_specs.items():
                    if v and str(v).strip().upper() != "N/A": lines.append(f"     • {k.replace('_', ' ').capitalize()}: {v}")

            date_audit = legal_analysis.get("date_audit", {})
            if date_audit and any(v for v in date_audit.values() if v and (isinstance(v, list) and v and any(item.strip().upper() != "N/A" for item in v) or (isinstance(v, str) and v.strip().upper() != "N/A"))):
                lines.append(f"   🗓️ {i18n.t('report_date_audit')}:")
                all_dates = date_audit.get("all_dates_found", [])
                if all_dates and any(d.strip().upper() != "N/A" for d in all_dates):
                    lines.append(f"     • {i18n.t('report_found_dates')}: {', '.join(all_dates)}")
                past_warnings = date_audit.get("past_dates_warnings", [])
                if past_warnings and any(w.strip().upper() != "N/A" for w in past_warnings):
                    lines.append(f"     ❗ {i18n.t('report_past_date_warnings')}:")
                    for w in past_warnings:
                        lines.append(f"       - {w}")
                illogical_ranges = date_audit.get("illogical_date_ranges", [])
                if illogical_ranges and any(ir.strip().upper() != "N/A" for ir in illogical_ranges):
                    lines.append(f"     ❌ {i18n.t('report_illogical_date_ranges')}:")
                    for ir in illogical_ranges:
                        lines.append(f"       - {ir}")

            gram_corrections = legal_analysis.get("grammatical_corrections", [])
            if gram_corrections and any(gc.strip().upper() != "N/A" for gc in gram_corrections):
                lines.append(f"   📝 {i18n.t('report_grammatical_corrections')}:")
                for gc in gram_corrections:
                    lines.append(f"     - {gc}")

            party_inconsistencies = legal_analysis.get("party_name_inconsistencies", [])
            if party_inconsistencies and any(pi.strip().upper() != "N/A" for pi in party_inconsistencies):
                lines.append(f"   👤 {i18n.t('report_party_name_inconsistencies')}:")
                for pi in party_inconsistencies:
                    lines.append(f"     - {pi}")
            lines.append("")

        suggested_corrections = analysis.get("suggested_corrections", [])
        if suggested_corrections and any(sc.strip().upper() != "N/A" for sc in suggested_corrections):
            # Update: Replaced with direct localized string
            lines.append("IETEIKTIE LABOJUMI UN KOREKCIJAS:")
            for sc in suggested_corrections:
                lines.append(f"   • {sc}")
            lines.append("")

        summary = analysis.get("summary", i18n.t("not_available_short"))
        if summary.strip().upper() != "N/A" and summary.strip() != "":
            # Update: Replaced with direct localized string
            lines.append("DOKUMENTA KOPSAVILKUMS SAPROTAMĀ VALODĀ:")
            lines.append(f"   {summary}\n")

        risks = analysis.get("risks_warnings", [])
        if risks and any(r_item for r_item in risks if r_item and r_item.strip().upper() != "N/A" and r_item.strip() != ""):
            # Update: Replaced with direct localized string
            lines.append("GALVENIE RISKI UN BRĪDINĀJUMI:")
            if isinstance(risks, list):
                for r_item in risks:
                    if r_item and r_item.strip().upper() != "N/A" and r_item.strip() != "":
                        lines.append(f"   🚩 {r_item}")
            else:
                if risks and risks.strip().upper() != "N/A" and risks.strip() != "":
                    lines.append(f"   🚩 {risks}")
            lines.append("")


        lines.append("-" * 70)
        lines.append(f"💾 {i18n.t('report_word_summary_saved')}:\n   {docx_path}\n\n")
        return "\n".join(lines)

    def select_and_process(self):
        path = filedialog.askopenfilename(
            filetypes=[(i18n.t("file_type_documents"), "*.pdf *.docx *.pptx *.txt"), (i18n.t("file_type_all_files"), "*.*")]
        )
        if not path:
            return

        p = Path(path)
        self.btn_select.configure(state="disabled")
        self.btn_open_summary.configure(state="disabled")
        
        self._update_gui_progress(0.0, i18n.t("status_starting_processing", filename=p.name))
        self.txt_output.insert("end", f"⚡ {i18n.t('status_starting_processing', filename=p.name)}\n")
        self.txt_output.see("end")

        def _worker():
            analysis = None
            res_msg = ""
            docx_path_result = None
            analysis_successful = False

            try:
                self.after(0, lambda: self._update_gui_progress(0.1, i18n.t("status_checking_file", filename=p.name)))
                if not p.exists():
                    raise FileNotFoundError(f"{i18n.t('error_file_not_found', filename=p.name)}.")

                ext = p.suffix.lower()
                payload = None
                
                self.after(0, lambda: self._update_gui_progress(0.2, i18n.t("status_reading_content", filename=p.name)))
                if ext == ".docx":
                    payload = self.service._read_docx(p)
                elif ext == ".txt":
                    payload = self.service._read_txt(p)
                elif ext == ".pptx":
                    payload = self.service._read_pptx(p)
                elif ext == ".pdf":
                    payload = p 
                else:
                    raise ValueError(i18n.t("error_unsupported_format"))

                if not payload:
                    raise ValueError(i18n.t("error_failed_to_read_content"))

                self.after(0, lambda: self._update_gui_progress(0.4, i18n.t("status_running_ai_analysis")))
                analysis = self.service._analyze_with_gemini(payload, file_type=ext.replace(".", ""))
                if "error" in analysis:
                    raise RuntimeError(analysis["error"])
                
                self.after(0, lambda: self._update_gui_progress(0.8, i18n.t("status_generating_word_and_registering")))
                out_name = f"DIGEST_{p.stem}.docx"
                out_path_base = OUTBOX_DIR / out_name
                
                docx_path_result = self.service._create_digest_docx(out_path_base, analysis, p.name)
                
                append_to_excel_registry(p.name, analysis)

                analysis_successful = True
                res_msg = str(docx_path_result)

            except Exception as e:
                analysis_successful = False
                res_msg = f"{i18n.t('error_system_error')}: {e}"
                self.after(0, lambda: self._update_gui_progress(0.0, f"{i18n.t('error_generic')}: {e}"))

            def _done():
                self.btn_select.configure(state="normal")
                
                if analysis_successful:
                    self.last_summary_path = Path(docx_path_result)
                    self.btn_open_summary.configure(state="normal")
                    
                    self._update_gui_progress(1.0, i18n.t("status_processing_complete"))
                    clean_report = self._format_clean_report(p.name, analysis, docx_path_result)
                    self.txt_output.insert("end", clean_report)
                    self.txt_output.see("end")
                else:
                    self.txt_output.insert("end", f"\n❌ {i18n.t('error_for_file', filename=p.name)}: {res_msg}\n\n")
                    self.btn_open_summary.configure(state="disabled")
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()