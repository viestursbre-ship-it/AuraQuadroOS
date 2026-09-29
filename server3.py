import os
import time
import threading
from datetime import datetime
import docx # Added for docx processing
import re # Added for regex parsing

# --- Core Infrastructure Services ---

class EventEngine:
    def __init__(self):
        self.listeners = {}

    def subscribe(self, event_type, listener):
        if event_type not in self.listeners:
            self.listeners[event_type] = []
        self.listeners[event_type].append(listener)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] EventEngine: Listener '{listener.__name__}' subscribed to '{event_type}'")

    def emit(self, event_type, *args, **kwargs):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] EventEngine: Emitting event '{event_type}' with args: {args}, kwargs: {kwargs}")
        if event_type in self.listeners:
            for listener in self.listeners[event_type]:
                try:
                    listener(*args, **kwargs)
                except Exception as e:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] EventEngine: Error in listener '{listener.__name__}' for event '{event_type}': {e}")

class ServiceRegistry:
    def __init__(self):
        self.services = {}

    def register(self, service_name, service_instance):
        self.services[service_name] = service_instance
        print(f"[{datetime.now().strftime('%H:%M:%S')}] ServiceRegistry: Service '{service_name}' registered.")

    def get_service(self, service_name):
        return self.services.get(service_name)

class GatewayRouter:
    def __init__(self, service_registry, event_engine):
        self.service_registry = service_registry
        self.event_engine = event_engine
        self.routes = {
            '/health': self._health_check,
            '/api/v1/status': self._status_check
        }
        print(f"[{datetime.now().strftime('%H:%M:%S')}] GatewayRouter: Initialized with routes: {list(self.routes.keys())}")

    def _health_check(self):
        return {"status": "ok", "timestamp": datetime.now().isoformat()}

    def _status_check(self):
        # In a real scenario, this would gather status from registered services
        return {"system_status": "operational", "services_active": len(self.service_registry.services)}

    def handle_request(self, path):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] GatewayRouter: Handling request for path: {path}")
        handler = self.routes.get(path)
        if handler:
            return handler()
        else:
            return {"error": "Not Found", "path": path}, 404

# --- Application Services ---

class FolderWatcherService:
    def __init__(self, watched_folder, event_engine):
        self.watched_folder = watched_folder
        self.event_engine = event_engine
        self.known_files = set()
        self.running = False
        print(f"[{datetime.now().strftime('%H:%M:%S')}] FolderWatcherService: Watching folder '{self.watched_folder}'")

    def _scan_folder(self):
        if not os.path.exists(self.watched_folder):
            print(f"[{datetime.now().strftime('%H:%M:%S')}] FolderWatcherService: Watched folder '{self.watched_folder}' does not exist. Creating it.")
            os.makedirs(self.watched_folder, exist_ok=True)

        current_files = set()
        for root, _, files in os.walk(self.watched_folder):
            for file in files:
                file_path = os.path.join(root, file)
                current_files.add(file_path)

        new_files = current_files - self.known_files
        for file_path in new_files:
            # LEO: Viesturs bija *pilnīgi* *perfetto*! Nekādu astu graizīšanas!
            if os.path.basename(file_path).startswith('DIGEST_'):
                print(f"[{datetime.now().strftime('%H:%M:%S')}] FolderWatcherService: Ignoring generated digest file: {file_path}")
                continue # Ignore generated digest files, preventing infinite loop

            print(f"[{datetime.now().strftime('%H:%M:%S')}] FolderWatcherService: Detected new file: {file_path}")
            self.event_engine.emit('file_dropped', file_path)
        
        self.known_files = current_files

    def start(self):
        self.running = True
        thread = threading.Thread(target=self._run)
        thread.daemon = True # Allow main program to exit even if thread is running
        thread.start()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] FolderWatcherService: Started background watcher.")

    def _run(self):
        while self.running:
            self._scan_folder()
            time.sleep(5) # Scan every 5 seconds

    def stop(self):
        self.running = False
        print(f"[{datetime.now().strftime('%H:%M:%S')}] FolderWatcherService: Stopped.")

class DocDigestService:
    def __init__(self, event_engine, outbox_dir):
        self.event_engine = event_engine
        self.outbox_dir = outbox_dir
        if not os.path.exists(self.outbox_dir):
            os.makedirs(self.outbox_dir, exist_ok=True)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] DocDigestService: Created outbox directory: '{self.outbox_dir}'")
        print(f"[{datetime.now().strftime('%H:%M:%S')}] DocDigestService: Initialized. Digests will be saved to '{self.outbox_dir}'")

    def handle_file_dropped(self, file_path):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] DocDigestService: Received file for digestion: {file_path}")
        
        try:
            document_text = self._read_document_text(file_path)
            digest = self._parse_document_for_digest(document_text)
            self._save_digest(file_path, digest)
        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] DocDigestService: Error processing file '{file_path}': {e}")
            # Optionally emit an error event
            self.event_engine.emit('digest_error', file_path, str(e))

    def _read_document_text(self, file_path):
        _, file_extension = os.path.splitext(file_path)
        file_extension = file_extension.lower()

        if file_extension == '.txt':
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        elif file_extension == '.docx':
            doc = docx.Document(file_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            return '\n'.join(full_text)
        else:
            raise ValueError(f"Unsupported file type for digestion: {file_extension}")

    def _parse_document_for_digest(self, text):
        digest_data = {
            "document_type": "Unknown",
            "parties": [],
            "terms": [],
            "amounts": [],
            "summary": "",
            "processed_at": datetime.now().isoformat()
        }

        # Document Type (Līgums, Rēķins, Vienošanās)
        if re.search(r'\bLīgums\b', text, re.IGNORECASE):
            digest_data["document_type"] = "Līgums"
        elif re.search(r'\bRēķins\b', text, re.IGNORECASE):
            digest_data["document_type"] = "Rēķins"
        elif re.search(r'\bVienošanās\b', text, re.IGNORECASE):
            digest_data["document_type"] = "Vienošanās"
        
        # Parties (simple example, could be more complex with company names, addresses, etc.)
        party_patterns = [
            r'(?:starp|puses?)\s*([\w\s.,&-]+)\s*(?:un|&|,)\s*([\w\s.,&-]+)', # "starp X un Y", captures X, Y
            r'(?:[Pp]uses?:?\s*)([\w\s.,&-]+(?:(?:\s*(?:un|,|&)\s*)[\w\s.,&-]+)*)', # "Puses: X, Y un Z", captures "X, Y un Z"
        ]
        unique_parties = set()
        for pattern in party_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                if isinstance(match, tuple):
                    for item in match:
                        parts = [p.strip() for p in re.split(r'\s*(?:un|,|&)\s*', item) if p.strip()]
                        unique_parties.update(parts)
                else: # single string match
                    parts = [p.strip() for p in re.split(r'\s*(?:un|,|&)\s*', match) if p.strip()]
                    unique_parties.update(parts)
        
        # Filter out very short strings or common conjunctions that might sneak in
        digest_data["parties"] = [p for p in list(unique_parties) if len(p) > 2 and p.lower() not in ['un', '&', ',', 'puses', 'starp']]
        
        # Terms (e.g., dates)
        date_patterns = [
            r'\b(\d{2}\.\d{2}\.\d{4})\b', # DD.MM.YYYY
            r'\b(\d{4}-\d{2}-\d{2})\b', # YYYY-MM-DD
            r'\bTermiņš līdz (\d{2}\.\d{2}\.\d{4})\b', # "Termiņš līdz DD.MM.YYYY"
        ]
        for pattern in date_patterns:
            terms = re.findall(pattern, text)
            for term in terms:
                if term not in digest_data["terms"]:
                    digest_data["terms"].append(term)
        
        # Amounts (e.g., EUR, €, summas)
        amount_patterns = [
            r'(?:EUR|€)\s*(\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})?)', # EUR 1 234.56 or €1.234,56 (need to handle both decimal separators)
            r'(\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})?)\s*(?:EUR|€)', # 1 234.56 EUR
            r'\b(?:Summa|Cena|Maksa):\s*(\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})?)', # Summa: 100.00
            r'\b(\d{1,}(?:[.,]\d{2})?)\s*(?:eiro|EUR|€)\b', # 100 eiro, 100 EUR
        ]
        for pattern in amount_patterns:
            amounts = re.findall(pattern, text, re.IGNORECASE)
            for amount in amounts:
                # Normalize amount string (remove spaces, replace comma with dot for float conversion)
                if isinstance(amount, tuple): # If pattern has groups, take the relevant one
                    amount_str = amount[0]
                else:
                    amount_str = amount
                
                amount_str = amount_str.replace(' ', '').replace(',', '.')
                try:
                    float_amount = float(amount_str)
                    if float_amount not in digest_data["amounts"]: # Avoid duplicates
                        digest_data["amounts"].append(float_amount)
                except ValueError:
                    pass # Ignore if it's not a valid number
        
        # Simple Summary (first few lines or a paragraph)
        lines = text.split('\n')
        non_empty_lines = [line.strip() for line in lines if line.strip()]
        digest_data["summary"] = ' '.join(non_empty_lines[:3]) + ('...' if len(non_empty_lines) > 3 else '')
        if len(digest_data["summary"]) > 200: # Truncate if too long
            digest_data["summary"] = digest_data["summary"][:200] + "..."

        return digest_data

    def _save_digest(self, original_file_path, digest_data):
        original_filename = os.path.basename(original_file_path)
        filename_without_ext, _ = os.path.splitext(original_filename)
        out_filename = f"DIGEST_{filename_without_ext}.txt"
        out_path = os.path.join(self.outbox_dir, out_filename)

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(f"--- Document Digest for '{original_filename}' ---\n")
            f.write(f"Processed At: {digest_data['processed_at']}\n\n")
            f.write(f"Document Type: {digest_data['document_type']}\n")
            f.write(f"Parties: {', '.join(digest_data['parties']) if digest_data['parties'] else 'N/A'}\n")
            f.write(f"Terms: {', '.join(digest_data['terms']) if digest_data['terms'] else 'N/A'}\n")
            f.write(f"Amounts: {', '.join(map(str, digest_data['amounts'])) if digest_data['amounts'] else 'N/A'}\n")
            f.write(f"\n--- Summary ---\n")
            f.write(digest_data['summary'] + "\n")
            f.write(f"\n--- Full Digest Data (Raw) ---\n")
            for key, value in digest_data.items():
                f.write(f"{key}: {value}\n")
        
        print(f"[{datetime.now().strftime('%H:%M:%S')}] DocDigestService: Digest saved to '{out_path}'")
        self.event_engine.emit('file_digested', out_path, original_file_path, digest_data)


# --- Main Application Setup ---
if __name__ == "__main__":
    # Define directories
    WATCHED_FOLDER = 'watched_files'
    OUTBOX_FOLDER = 'digested_output' # This must be different from WATCHED_FOLDER

    # Create directories if they don't exist
    os.makedirs(WATCHED_FOLDER, exist_ok=True)
    os.makedirs(OUTBOX_FOLDER, exist_ok=True)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: Initialized '{WATCHED_FOLDER}' and '{OUTBOX_FOLDER}' directories.")

    # Initialize Core Services
    event_engine = EventEngine()
    service_registry = ServiceRegistry()
    gateway_router = GatewayRouter(service_registry, event_engine)

    # Initialize Application Services
    folder_watcher = FolderWatcherService(WATCHED_FOLDER, event_engine)
    doc_digest_service = DocDigestService(event_engine, OUTBOX_FOLDER)

    # Register Services
    service_registry.register("EventEngine", event_engine)
    service_registry.register("ServiceRegistry", service_registry)
    service_registry.register("GatewayRouter", gateway_router)
    service_registry.register("FolderWatcherService", folder_watcher)
    service_registry.register("DocDigestService", doc_digest_service)

    # Subscribe DocDigestService to 'file_dropped' events
    event_engine.subscribe('file_dropped', doc_digest_service.handle_file_dropped)

    # Start FolderWatcherService in a background thread
    folder_watcher.start()

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: System started. Monitoring '{WATCHED_FOLDER}'. Digests to '{OUTBOX_FOLDER}'.")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: Try accessing /health or /api/v1/status via GatewayRouter.handle_request(path).")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: Or drop .txt or .docx files into '{WATCHED_FOLDER}'.")
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: Press Ctrl+C to exit.")

    try:
        # Keep the main thread alive for demonstration purposes
        # In a real web server, this would be a Flask/FastAPI/Django run loop
        while True:
            # Simulate external requests, for now just idle
            # print(gateway_router.handle_request('/health'))
            time.sleep(10) # Main thread just sleeps, background watcher does its job
    except KeyboardInterrupt:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: KeyboardInterrupt detected. Shutting down...")
        folder_watcher.stop()
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Main: Goodbye!")