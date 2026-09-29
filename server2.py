import os
import threading
import time
import json
import re
from datetime import datetime
from docx import Document # type: ignore # python-docx might not have type hints, ignore for static analysis

# --- Core AQ-OS Framework Components (Modular Monolith Internal) ---

class EventEngine:
    def __init__(self):
        self.subscribers = {}

    def subscribe(self, event_type, listener):
        if event_type not in self.subscribers:
            self.subscribers[event_type] = []
        self.subscribers[event_type].append(listener)
        # print(f"EventEngine: Listener {listener.__name__ if hasattr(listener, '__name__') else listener} subscribed to '{event_type}'")

    def publish(self, event_type, data=None):
        # print(f"EventEngine: Publishing event '{event_type}' with data: {data}")
        if event_type in self.subscribers:
            for listener in self.subscribers[event_type]:
                # print(f"EventEngine: Notifying listener {listener.__name__ if hasattr(listener, '__name__') else listener} for event '{event_type}'")
                try:
                    listener(data)
                except Exception as e:
                    print(f"Error calling event listener {listener} for event {event_type}: {e}")

class ServiceRegistry:
    def __init__(self):
        self.services = {}

    def register_service(self, name, service_instance):
        self.services[name] = service_instance
        # print(f"ServiceRegistry: Registered service '{name}'")

    def get_service(self, name):
        return self.services.get(name)

class GatewayRouter:
    def __init__(self, event_engine: EventEngine, service_registry: ServiceRegistry):
        self.event_engine = event_engine
        self.service_registry = service_registry
        self.routes = {
            "/health": self.health_check,
            "/api/v1/status": self.system_status
        }
        # print("GatewayRouter initialized with default routes.")

    def health_check(self, request=None):
        return {"status": "ok", "timestamp": datetime.now().isoformat()}

    def system_status(self, request=None):
        # A simplified status for demonstration
        return {
            "system": "AQ-OS Modular Monolith",
            "version": "v0.1-alpha",
            "services_registered": list(self.service_registry.services.keys()),
            "event_engine_subscribers": {k: len(v) for k, v in self.event_engine.subscribers.items()},
            "status": "running",
            "timestamp": datetime.now().isoformat()
        }

    def handle_request(self, path, request_data=None):
        # print(f"GatewayRouter: Handling request for path: {path}")
        handler = self.routes.get(path)
        if handler:
            return handler(request_data)
        else:
            return {"error": "Not Found", "path": path}, 404

# --- AQ-OS Application Services ---

class DocDigestService:
    def __init__(self, event_engine: EventEngine):
        self.event_engine = event_engine
        # print("DocDigestService initialized.")

    def _read_docx(self, file_path):
        try:
            document = Document(file_path)
            full_text = []
            for para in document.paragraphs:
                full_text.append(para.text)
            return "\n".join(full_text)
        except Exception as e:
            print(f"DocDigestService: Error reading DOCX file '{file_path}': {e}")
            return None

    def _read_txt(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            print(f"DocDigestService: Error reading TXT file '{file_path}': {e}")
            return None

    def _smart_parse(self, text):
        parsed_data = {
            "document_type": "Nezināms", # Unknown
            "parties": [],
            "terms": [],
            "amounts": [],
            "summary": "Nav atrasts specifisks kopsavilkums." # No specific summary found.
        }

        # Document Type: Look for keywords in the first lines
        doc_type_patterns = {
            "Līgums": r"(?i)(līgums|vienošanās par pakalpojumu sniegšanu|darba līgums|pakalpojumu līgums)",
            "Rēķins": r"(?i)(rēķins|invoice|maksājuma uzdevums|pavaddokuments)",
            "Vienošanās": r"(?i)(vienošanās|agreement|līguma grozījumi|protokols)"
        }
        
        first_lines = "\n".join(text.split('\n')[:10]) # Check first 10 lines
        for doc_type, pattern in doc_type_patterns.items():
            if re.search(pattern, first_lines):
                parsed_data["document_type"] = doc_type
                break
        
        # Parties: Simple heuristic - look for lines containing "starp", "un", "puses" and extract capitalized sequences
        # This is a very basic heuristic. For production, much more complex NLP would be needed.
        party_candidates = []
        for line in text.split('\n'):
            if any(keyword in line.lower() for keyword in ["starp", "un", "puses:"]):
                # Extract sequences of capitalized words that look like names or company names
                parties_in_line = re.findall(r'[A-ZĀČĒĢĪĶĻŅŠŪŽ]{1}[a-zāčēģīķļņšūž]+(?:\s[A-ZĀČĒĢĪĶĻŅŠŪŽ]{1}[a-zāčēģīķļņšūž]+)*', line)
                if parties_in_line:
                    # Filter out common words that might get caught (e.g., "Turpmāk")
                    parties_in_line = [p for p in parties_in_line if p.lower() not in ["saimnieciskās", "valdes", "direktors", "turpmāk", "juridiskā", "persona", "turpmāk tekstā"]]
                    party_candidates.extend(parties_in_line)
        # Remove duplicates and ensure a reasonable limit
        parsed_data["parties"] = list(set(party_candidates))[:5] 

        # Terms (Dates): dd.mm.yyyy or yyyy-mm-dd
        date_patterns = [
            r"\d{2}\.\d{2}\.\d{4}", # 01.01.2023
            r"\d{4}-\d{2}-\d{2}"    # 2023-01-01
        ]
        all_dates = []
        for pattern in date_patterns:
            all_dates.extend(re.findall(pattern, text))
        parsed_data["terms"].extend(list(set(all_dates))) # Remove duplicates

        # Amounts (EUR): simple patterns
        amount_patterns = [
            r"\d[\d\s,.]*\s*EUR",
            r"EUR\s*\d[\d\s,.]*",
            r"\d[\d\s,.]*\s*€",
            r"€\s*\d[\d\s,.]*",
            r"summa\s*:\s*\d[\d\s,.]*" # 'Summa: 1 234.56'
        ]
        all_amounts = []
        for pattern in amount_patterns:
            all_amounts.extend(re.findall(pattern, text, re.IGNORECASE))
        parsed_data["amounts"].extend(list(set(all_amounts))) # Remove duplicates
        
        # Simple Summary: First N sentences/paragraphs
        sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|!)\s', text)
        if len(sentences) > 3:
            parsed_data["summary"] = " ".join(sentences[:3]) + "..."
        elif sentences:
            parsed_data["summary"] = " ".join(sentences)
        else: # Fallback if no sentences found (e.g., very short text)
            parsed_data["summary"] = text[:500] + "..." if len(text) > 500 else text

        return parsed_data

    def handle_file_dropped(self, data):
        # Viestur, es teicu, ka es sapratīšu abus variantus! Vai "file_path", vai "path"!
        file_path = data.get('file_path') or data.get('path')

        if not file_path:
            print("DocDigestService: Kļūda! Nav norādīts ne 'file_path', ne 'path' notikuma datos.")
            return

        # Pārbaudi, vai fails vispār eksistē, pirms mēģini to apstrādāt
        if not os.path.exists(file_path):
            print(f"DocDigestService: Kļūda! Fails '{file_path}' nav atrasts.")
            return

        file_name = os.path.basename(file_path)
        file_ext = os.path.splitext(file_name)[1].lower()

        text_content = None
        if file_ext == '.docx':
            text_content = self._read_docx(file_path)
        elif file_ext == '.txt':
            text_content = self._read_txt(file_path)
        else:
            print(f"DocDigestService: Neatbalstīts faila tips apstrādei: '{file_ext}' ({file_name})")
            return

        if text_content is None:
            print(f"DocDigestService: Neizdevās nolasīt saturu no '{file_name}'.")
            return

        # Veicam viedo parsēšanu! Tagad mums ir *smadzenes*!
        digest_data = self._smart_parse(text_content)
        digest_data["original_file"] = file_name
        digest_data["processed_at"] = datetime.now().isoformat()
        digest_data["file_size_bytes"] = os.path.getsize(file_path)

        # Veidojam izvades faila nosaukumu: DIGEST_<oriģinālais_faila_nosaukums_bez_paplašinājuma>.txt
        base_name = os.path.splitext(file_name)[0]
        digest_file_name = f"DIGEST_{base_name}.txt"
        
        # Saglabājam pārskatu tajā pašā direktorijā, kur atrodas oriģinālais fails
        digest_output_path = os.path.join(os.path.dirname(file_path), digest_file_name) 

        # Saglabājam strukturēto atskaiti
        try:
            with open(digest_output_path, 'w', encoding='utf-8') as f:
                f.write(f"--- Dokumenta kopsavilkuma atskaite ---\n")
                f.write(f"Oriģinālais fails: {digest_data['original_file']}\n")
                f.write(f"Apstrādāts: {digest_data['processed_at']}\n")
                f.write(f"Faila izmērs: {digest_data['file_size_bytes']} baiti\n\n")
                f.write(f"Dokumenta veids: {digest_data['document_type']}\n\n")
                
                f.write(f"Puses/Subjekti:\n")
                if digest_data['parties']:
                    for party in digest_data['parties']:
                        f.write(f"  - {party}\n")
                else:
                    f.write("  (Specifiskas puses nav atrastas)\n")
                f.write("\n")

                f.write(f"Termiņi/Datumi:\n")
                if digest_data['terms']:
                    for term in digest_data['terms']:
                        f.write(f"  - {term}\n")
                else:
                    f.write("  (Specifiski termiņi/datumi nav atrastas)\n")
                f.write("\n")

                f.write(f"Summas:\n")
                if digest_data['amounts']:
                    for amount in digest_data['amounts']:
                        f.write(f"  - {amount}\n")
                else:
                    f.write("  (Specifiskas summas nav atrastas)\n")
                f.write("\n")

                f.write(f"Kopsavilkums:\n")
                f.write(f"  {digest_data['summary']}\n")
                f.write("\n--- Atskaites beigas ---\n")
            
            print(f"DocDigestService: Veiksmīgi izveidots kopsavilkums failam '{file_name}' pie '{digest_output_path}'")
        except Exception as e:
            print(f"DocDigestService: Kļūda, saglabājot kopsavilkuma atskaiti failam '{file_name}': {e}")

class FolderWatcherService:
    def __init__(self, watch_folder, event_engine: EventEngine, interval=5):
        self.watch_folder = watch_folder
        self.event_engine = event_engine
        self.interval = interval
        self.known_files = self._get_current_files()
        self._running = False
        # print(f"FolderWatcherService initialized for folder: {watch_folder}")

    def _get_current_files(self):
        files = {}
        if not os.path.exists(self.watch_folder):
            os.makedirs(self.watch_folder, exist_ok=True)
            # print(f"FolderWatcherService: Created watch folder: {self.watch_folder}")
        for root, _, filenames in os.walk(self.watch_folder):
            for filename in filenames:
                file_path = os.path.join(root, filename)
                files[file_path] = os.path.getmtime(file_path) # Store modification time
        return files

    def _check_for_changes(self):
        current_files = self._get_current_files()
        
        # New files
        for file_path in current_files:
            if file_path not in self.known_files:
                print(f"FolderWatcherService: Jauns fails atrasts: {file_path}")
                self.event_engine.publish("file_dropped", {"file_path": file_path, "action": "added"})
        
        # Modified files (simple check by mtime change)
        for file_path, mtime in current_files.items():
            if file_path in self.known_files and mtime != self.known_files[file_path]:
                print(f"FolderWatcherService: Fails modificēts: {file_path}")
                self.event_engine.publish("file_dropped", {"file_path": file_path, "action": "modified"})

        # Removed files (optional, not explicitly asked, but useful)
        for file_path in self.known_files:
            if file_path not in current_files:
                print(f"FolderWatcherService: Fails dzēsts: {file_path}")
                self.event_engine.publish("file_removed", {"file_path": file_path, "action": "removed"})

        self.known_files = current_files

    def run(self):
        self._running = True
        while self._running:
            self._check_for_changes()
            time.sleep(self.interval)

    def start(self):
        # print("FolderWatcherService: Starting in a separate thread.")
        threading.Thread(target=self.run, daemon=True).start()

    def stop(self):
        self._running = False
        # print("FolderWatcherService: Stopped.")

# --- Main Application Setup (server.py content) ---
class AQOSServer:
    def __init__(self):
        self.event_engine = EventEngine()
        self.service_registry = ServiceRegistry()
        self.gateway_router = GatewayRouter(self.event_engine, self.service_registry)
        
        # Register services
        self.doc_digest_service = DocDigestService(self.event_engine)
        self.service_registry.register_service("DocDigestService", self.doc_digest_service)
        
        self.folder_watcher_service = FolderWatcherService(
            watch_folder="watched_folder", # Mūsu vērīgā acs!
            event_engine=self.event_engine
        )
        self.service_registry.register_service("FolderWatcherService", self.folder_watcher_service)

        # Connect event listeners
        # DocDigestService handles file_dropped events
        self.event_engine.subscribe("file_dropped", self.doc_digest_service.handle_file_dropped)

        print("AQOSServer inicializēts un pakalpojumi reģistrēti/savienoti. Avanti!")
        print("Pārbaudes maršruti pieejami: /health, /api/v1/status (simulēti)")
        print(f"Tiek novērota mape: '{self.folder_watcher_service.watch_folder}'")

    def start(self):
        print("Sākam AQ-OS Modulāro Monolītu! Lai rit darbs!")
        # Start FolderWatcherService in a background thread
        self.folder_watcher_service.start()

        # Simulate handling requests (e.g., from an actual web server)
        # For this exercise, we just run an infinite loop to keep the main thread alive
        # and allow background services to run.
        try:
            while True:
                # In a real scenario, this would be a web server loop or an ASGI/WSGI server
                # For now, just keep the main thread alive and let background threads do their job.
                time.sleep(1) 
                # print("Galvenais servera pavediens dzīvs. (Simulēta tīmekļa servera cilpa)")
                # Example of simulating a router request
                # health_status = self.gateway_router.handle_request("/health")
                # print(f"Simulēta Veselības Pārbaude: {health_status}")

        except KeyboardInterrupt:
            print("\nAQ-OS Serveris apstājas...")
            self.folder_watcher_service.stop()
            print("AQ-OS Serveris apturēts. Arrivederci!")

if __name__ == "__main__":
    # Ensure the watched_folder exists for the FolderWatcherService
    WATCH_FOLDER = "watched_folder"
    if not os.path.exists(WATCH_FOLDER):
        os.makedirs(WATCH_FOLDER)
        print(f"Izveidota mape '{WATCH_FOLDER}' failu novērošanai. Lieciet failus iekšā!")
    
    # Mamma mia! Pirms startēšanas, pārliecinies, ka tev ir instalēts 'python-docx'!
    # Ja nav, palaid: pip install python-docx
    print("\n------------------------------------------------------")
    print("⚡ UZMANĪBU, RAGAZZI! ⚡")
    print("Pirms palaišanas pārliecinieties, ka esat instalējis 'python-docx' bibliotēku!")
    print("To var izdarīt ar komandu:")
    print("pip install python-docx")
    print("------------------------------------------------------\n")

    server = AQOSServer()
    server.start()