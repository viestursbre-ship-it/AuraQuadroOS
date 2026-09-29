import os
import time
import json
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import google.generativeai as genai

# --- Konfigurācija un pamata iestatījumi ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

WATCH_DIRECTORY = "storage/watched_files"
OUTPUT_DIRECTORY = "storage/outbox"
DIGEST_PREFIX = "DIGEST_"
WORD_LOCK_PREFIX = "~$" # Faili, kas sākas ar ~$, ir pagaidu Word bloķēšanas faili

# Nodrošina, ka nepieciešamās mapes eksistē
os.makedirs(WATCH_DIRECTORY, exist_ok=True)
os.makedirs(OUTPUT_DIRECTORY, exist_ok=True)

# --- AQ-OS galvenās komponentes (modulārais monolīts) ---
class ServiceRegistry:
    """
    Vienkāršs reģistrs iekšējiem servisiem monolītā.
    Ļauj moduļiem "reģistrēties" un atklāt citus moduļus.
    """
    def __init__(self):
        self._services = {}

    def register_service(self, name, service_instance):
        logging.info(f"ServiceRegistry: Reģistrē servisu '{name}'")
        self._services[name] = service_instance

    def get_service(self, name):
        if name not in self._services:
            logging.warning(f"ServiceRegistry: Serviss '{name}' nav atrasts.")
            return None
        return self._services[name]

class EventEngine:
    """
    Iekšējais notikumu apstrādes un izplatīšanas mehānisms.
    Nodrošina asinhronu komunikāciju starp monolīta moduļiem.
    """
    def __init__(self):
        self._listeners = {}

    def publish(self, event_type, data):
        logging.info(f"EventEngine: Publicē notikumu '{event_type}' par failu: {data.get('filepath', 'N/A')}")
        if event_type in self._listeners:
            for listener in self._listeners[event_type]:
                listener(data)

    def subscribe(self, event_type, listener_callback):
        logging.info(f"EventEngine: Abonē klausītāju notikumam '{event_type}'")
        if event_type not in self._listeners:
            self._listeners[event_type] = []
        self._listeners[event_type].append(listener_callback)

class GatewayRouter:
    """
    Iekšējais vārtejas maršrutētājs modulārā monolītā.
    Šobrīd tas ir konceptuāls, bet nākotnē apstrādās iekšējos pieprasījumus.
    """
    def __init__(self, service_registry_instance, event_engine_instance):
        self.service_registry = service_registry_instance
        self.event_engine = event_engine_instance
        logging.info("GatewayRouter inicializēts.")

    def handle_internal_request(self, path, payload):
        logging.info(f"GatewayRouter: Apstrādā iekšēju pieprasījumu ceļam '{path}'")
        # Piemērs: maršrutē uz specifisku servisu, pamatojoties uz ceļu
        if path == "/digest_document_event":
            doc_digest_service = self.service_registry.get_service("DocDigestService")
            if doc_digest_service:
                doc_digest_service.digest_document(payload["filepath"])
            else:
                logging.error("GatewayRouter: DocDigestService nav reģistrēts.")
        else:
            logging.warning(f"GatewayRouter: Nezināms ceļš '{path}'")


# --- DocDigestService: Dokumentu sagremošanas loģika ar AI un DOCX ģenerēšanu ---
class DocDigestService:
    def __init__(self, output_dir, ai_api_key):
        self.output_dir = output_dir
        self.ai_api_key = ai_api_key
        logging.info(f"DocDigestService inicializēts. Izvades mape: {self.output_dir}")

        if not self.ai_api_key:
            logging.error("AQ_AI_API_KEY vides mainīgais nav iestatīts. AI analīze nedarbosies!")
        else:
            genai.configure(api_key=self.ai_api_key)
            logging.info("Google Gemini API konfigurēts.")

    def _read_docx(self, filepath):
        """Nolasa tekstu no .docx faila, ieskaitot rindkopas un tabulas."""
        doc = Document(filepath)
        full_text = []

        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)

        for table in doc.tables:
            for row in table.rows:
                row_texts = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_texts.append(cell_text)
                if row_texts:
                    full_text.append(" | ".join(row_texts)) # Apvieno šūnu tekstus kontekstam

        return "\n".join(full_text)

    def _read_txt(self, filepath):
        """Nolasa tekstu no .txt faila."""
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()

    def _analyze_with_ai(self, text):
        """Sūta dokumenta tekstu Gemini semantiskai analīzei."""
        if not self.ai_api_key:
            logging.error("AI API atslēga trūkst. Nevar veikt AI analīzi.")
            return {"error": "AI API atslēga trūkst"}
        
        try:
            # Izmantojam gemini-1.5-flash ar JSON izvadi
            model = genai.GenerativeModel('gemini-1.5-flash',
                                          generation_config={"response_mime_type": "application/json"})
            
            prompt = f"""Tu esi eksperts dokumentu analītiķis. Analizē šo dokumenta tekstu un iegūsti galveno informāciju strukturētā JSON formātā. Esi precīzs un aptveri visus pieprasītos laukus.
            
            Dokumenta Teksts:
            {text}
            
            Iegūsti šādu informāciju:
            1.  **DocumentType**: piem., "Līgums", "Rēķins", "Sanāksmes protokols", "Vēstule", "Nezināms".
            2.  **DocumentNumber**: Oficiālais dokumenta identifikators, ja ir.
            3.  **DocumentDate**: Dokumenta galvenais datums (YYYY-MM-DD formātā).
            4.  **Parties**: Masīvs ar objektiem, katrs ar "Name" un "Role" (piem., "Maksātājs", "Pakalpojuma sniedzējs", "Klients", "Saņēmējs", "Izdevējs").
            5.  **SubjectMatter**: Kodolīgs darījuma vai satura kopsavilkums.
            6.  **FinancialTerms**: Objekts ar "TotalAmount" (skaitlis, ja atrasts), "Currency" (teksts, piem., "EUR", "USD"), "PaymentSchedule" (maksājumu datumu/posmu kopsavilkums), "Deadlines" (galvenie finanšu termiņi), "LatePaymentPenalties" (sankciju noteikumu kopsavilkums).
            7.  **OwnershipAndTermination**: Īpašuma tiesību nodošanas un līguma izbeigšanas nosacījumu kopsavilkums.
            8.  **KeyRisksOrWarnings**: Jebkādi kritiski riski, atrunas vai svarīgi brīdinājumi.
            
            Nodrošini, ka izvade ir stingri JSON. Ja lauks nav atrasts, izmanto null vai tukšu masīvu/virkni, kā atbilstoši.
            """
            
            logging.info("Sūta tekstu Gemini analīzei...")
            response = model.generate_content(prompt)
            
            # Pārbaudām, vai atbilde satur tekstu, pirms mēģinām to parsēt kā JSON
            if response.text:
                json_string = response.text
                logging.info(f"AI atbilde saņemta: {json_string[:200]}...") # Log pirmo 200 simbolus
                return json.loads(json_string)
            else:
                logging.error(f"Gemini atbilde nesatur tekstu. Iespējama kļūda API izsaukumā.")
                return {"error": "Gemini atbilde bija tukša."}
            
        except Exception as e:
            logging.error(f"Kļūda AI analīzes laikā: {e}")
            return {"error": str(e)}

    def _create_docx_report(self, original_filename, analysis_data, output_filepath):
        """Izveido formatētu .docx atskaiti no AI analīzes datiem."""
        doc = Document()

        # Pievieno virsrakstu
        title_paragraph = doc.add_paragraph()
        title_run = title_paragraph.add_run(f"AI Analīzes Kopsavilkums: {original_filename}")
        title_run.font.size = Pt(20)
        title_run.bold = True
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph() # Atstarpes rinda

        # Pievieno stilus virsrakstiem un aizzīmju sarakstiem
        styles = doc.styles
        if 'CustomHeading1' not in styles:
            styles.add_style('CustomHeading1', WD_STYLE_TYPE.PARAGRAPH)
            styles['CustomHeading1'].font.size = Pt(16)
            styles['CustomHeading1'].font.bold = True
        if 'CustomHeading2' not in styles:
            styles.add_style('CustomHeading2', WD_STYLE_TYPE.PARAGRAPH)
            styles['CustomHeading2'].font.size = Pt(14)
            styles['CustomHeading2'].font.bold = True
        
        def add_heading(text, level=1):
            style = 'CustomHeading1' if level == 1 else 'CustomHeading2'
            p = doc.add_paragraph(text, style=style)
            p.paragraph_format.space_after = Pt(10)

        def add_bullet(text):
            p = doc.add_paragraph(text, style='List Bullet') # Izmanto iebūvēto 'List Bullet' stilu
            p.paragraph_format.left_indent = Inches(0.5) # Iedaļa aizzīmēm

        # Vispārīgā Informācija
        add_heading("Vispārīgā Informācija", level=1)
        add_bullet(f"Dokumenta Tips: {analysis_data.get('DocumentType', 'Nav norādīts')}")
        add_bullet(f"Dokumenta Numurs: {analysis_data.get('DocumentNumber', 'Nav norādīts')}")
        add_bullet(f"Dokumenta Datums: {analysis_data.get('DocumentDate', 'Nav norādīts')}")

        # Puses un Lomas
        if analysis_data.get('Parties'):
            add_heading("Puses un Lomas", level=1)
            for party in analysis_data['Parties']:
                add_bullet(f"Vārds/Nosaukums: {party.get('Name', 'Nav norādīts')}, Loma: {party.get('Role', 'Nav norādīta')}")

        # Darījuma Priekšmets
        if analysis_data.get('SubjectMatter'):
            add_heading("Darījuma Priekšmets", level=1)
            doc.add_paragraph(analysis_data['SubjectMatter'])

        # Finanšu Noteikumi
        financial_terms = analysis_data.get('FinancialTerms', {})
        if financial_terms:
            add_heading("Finanšu Noteikumi", level=1)
            total_amount = financial_terms.get('TotalAmount')
            currency = financial_terms.get('Currency')
            if total_amount is not None and currency:
                add_bullet(f"Kopējā Summa: {total_amount} {currency}")
            elif total_amount is not None:
                add_bullet(f"Kopējā Summa: {total_amount}")
            add_bullet(f"Maksājumu Grafiks: {financial_terms.get('PaymentSchedule', 'Nav norādīts')}")
            add_bullet(f"Termiņi: {financial_terms.get('Deadlines', 'Nav norādīts')}")
            add_bullet(f"Kavējuma Sankcijas: {financial_terms.get('LatePaymentPenalties', 'Nav norādīts')}")

        # Īpašuma Tiesības un Līguma Izbeigšana
        if analysis_data.get('OwnershipAndTermination'):
            add_heading("Īpašuma Tiesības un Līguma Izbeigšana", level=1)
            doc.add_paragraph(analysis_data['OwnershipAndTermination'])

        # Būtiskākie Riski vai Brīdinājumi
        if analysis_data.get('KeyRisksOrWarnings'):
            add_heading("Būtiskākie Riski vai Brīdinājumi", level=1)
            doc.add_paragraph(analysis_data['KeyRisksOrWarnings'])

        doc.save(output_filepath)
        logging.info(f"Izveidota DOCX atskaite: {output_filepath}")

    def digest_document(self, filepath):
        """Apstrādā vienu dokumentu."""
        logging.info(f"DocDigestService: Apstrādā dokumentu {filepath}")
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)

        # Filtrē Word bloķēšanas failus un mūsu pašu DIGEST_ failus
        if filename.startswith(WORD_LOCK_PREFIX) or filename.startswith(DIGEST_PREFIX):
            logging.info(f"DocDigestService: Ignorē iekšēju/bloķēšanas failu: {filename}")
            return
        if not (ext.lower() == ".docx" or ext.lower() == ".txt"):
            logging.warning(f"DocDigestService: Neatbalstīts faila tips apstrādei: {filename}")
            return

        text_content = ""
        if ext.lower() == ".docx":
            text_content = self._read_docx(filepath)
        elif ext.lower() == ".txt":
            text_content = self._read_txt(filepath)
        
        if not text_content.strip():
            logging.warning(f"Dokuments {filename} nesatur nolasāmu tekstu. AI analīze tiek izlaista.")
            return

        analysis_data = self._analyze_with_ai(text_content)

        if "error" in analysis_data:
            logging.error(f"Neizdevās analizēt {filename}: {analysis_data['error']}")
            error_report_filepath = os.path.join(self.output_dir, f"{DIGEST_PREFIX}{name}_ERROR.txt")
            with open(error_report_filepath, "w", encoding="utf-8") as f:
                f.write(f"Kļūda apstrādājot {filename}:\n{analysis_data['error']}\n\nOriģinālais Teksts:\n{text_content}")
            return

        output_filename = f"{DIGEST_PREFIX}{name}.docx"
        output_filepath = os.path.join(self.output_dir, output_filename)
        self._create_docx_report(filename, analysis_data, output_filepath)
        logging.info(f"Veiksmīgi apstrādāts un izveidota atskaite par {filename} mapē {output_filepath}")


# --- FolderWatcherService: Failu izmaiņu uzraudzība ---
class FolderWatcherService(FileSystemEventHandler):
    def __init__(self, watched_dir, digest_prefix, word_lock_prefix, event_engine_instance):
        self.watched_dir = watched_dir
        self.digest_prefix = digest_prefix
        self.word_lock_prefix = word_lock_prefix
        self.event_engine = event_engine_instance
        self.processed_files = set() # Lai novērstu dubultu apstrādi

    def _is_relevant_file(self, filepath):
        filename = os.path.basename(filepath)
        # Ignorē failus, ko radījis mūsu apstrādes process, vai Word bloķēšanas failus
        if filename.startswith(self.digest_prefix) or filename.startswith(self.word_lock_prefix):
            logging.debug(f"FolderWatcher: Ignorē iekšēju/bloķēšanas failu: {filename}")
            return False
        # Ignorē mapes
        if os.path.isdir(filepath):
            logging.debug(f"FolderWatcher: Ignorē mapi: {filename}")
            return False
        # Apstrādā tikai .docx un .txt failus
        if not (filename.lower().endswith(".docx") or filename.lower().endswith(".txt")):
            logging.debug(f"FolderWatcher: Ignorē ne-DOCX/TXT failu: {filename}")
            return False
        return True

    def on_created(self, event):
        if not event.is_directory and self._is_relevant_file(event.src_path):
            logging.info(f"FolderWatcher: Detektēts jauns fails: {event.src_path}")
            if event.src_path not in self.processed_files:
                self.processed_files.add(event.src_path)
                self.event_engine.publish("file_created", {"filepath": event.src_path})
            else:
                logging.debug(f"FolderWatcher: Fails {event.src_path} jau ir apstrādes rindā.")

    def on_moved(self, event):
        # Apstrādā pārvietotus failus kā jaunradītus, ja tie nonāk uzraudzītajā mapē
        if not event.is_directory and self._is_relevant_file(event.dest_path):
            logging.info(f"FolderWatcher: Detektēts pārvietots fails: {event.dest_path}")
            if event.dest_path not in self.processed_files:
                self.processed_files.add(event.dest_path)
                self.event_engine.publish("file_created", {"filepath": event.dest_path})
            else:
                logging.debug(f"FolderWatcher: Fails {event.dest_path} jau ir apstrādes rindā.")
    
    # on_modified un on_deleted netiek izmantoti šajā versijā, lai izvairītos no dubultas apstrādes.
    # Fokuss ir uz failu izveidi/pārvietošanu.


# --- Galvenā lietojumprogrammas izpildes loģika ---
if __name__ == "__main__":
    logging.info("AQ-OS Modulārais Monolīts startē...")

    # Inicializē galvenās AQ-OS komponentes
    service_registry = ServiceRegistry()
    event_engine = EventEngine()
    gateway_router = GatewayRouter(service_registry, event_engine) # Pat ja neizmantojam tieši, tas ir arhitektūras pīlārs

    # Iegūst AI API atslēgu no vides mainīgā
    aq_ai_api_key = os.getenv("AQ_AI_API_KEY")

    # Inicializē DocDigestService un reģistrē to
    doc_digest_service = DocDigestService(OUTPUT_DIRECTORY, aq_ai_api_key)
    service_registry.register_service("DocDigestService", doc_digest_service)

    # Abonē DocDigestService notikumiem par failu izveidi
    event_engine.subscribe("file_created", lambda data: doc_digest_service.digest_document(data["filepath"]))

    # Inicializē un startē FolderWatcherService
    event_handler = FolderWatcherService(
        WATCH_DIRECTORY,
        DIGEST_PREFIX,
        WORD_LOCK_PREFIX,
        event_engine
    )
    observer = Observer()
    # Uzrauga tikai augstākā līmeņa mapi, nevis rekursīvi apakšmapes.
    # Tas ir svarīgi, lai neuzraudzītu "outbox" mapi, ja tā nejauši atrastos "watched_files" iekšienē.
    observer.schedule(event_handler, WATCH_DIRECTORY, recursive=False) 
    logging.info(f"FolderWatcherService startēts, uzrauga: {WATCH_DIRECTORY}")
    logging.info(f"Apstrādātās atskaites tiks saglabātas: {OUTPUT_DIRECTORY}")

    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logging.info("AQ-OS Modulārais Monolīts tiek apturēts.")
    observer.join()