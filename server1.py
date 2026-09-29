import os
import time
import threading
import datetime
import uuid
import json
from collections import defaultdict
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- 1. ServiceRegistry ---
class ServiceRegistry:
    """
    A simple registry for services within the modular monolith.
    Allows services to register themselves and discover others.
    """
    def __init__(self):
        self._services = {}

    def register(self, name, service_instance):
        if name in self._services:
            print(f"Warning: Service '{name}' already registered. Overwriting.")
        self._services[name] = service_instance
        print(f"Service '{name}' registered.")

    def get(self, name):
        service = self._services.get(name)
        if not service:
            print(f"Error: Service '{name}' not found in registry.")
        return service

# --- 2. EventEngine ---
class EventEngine:
    """
    The internal event processing and distribution mechanism.
    Supports asynchronous communication between modules.
    """
    def __init__(self):
        self._subscribers = defaultdict(list)
        self._event_queue = [] # For simple async, could be thread-safe queue for real-world
        self._running = False
        self._worker_thread = None
        self._lock = threading.Lock() # For thread-safe queue access

    def subscribe(self, event_type, handler):
        self._subscribers[event_type].append(handler)
        print(f"Handler {handler.__name__} subscribed to {event_type} events.")

    def publish(self, event_type, data=None):
        event_id = str(uuid.uuid4())
        timestamp = datetime.datetime.now().isoformat()
        event = {"event_id": event_id, "timestamp": timestamp, "type": event_type, "data": data}
        with self._lock:
            self._event_queue.append(event)
        print(f"Event '{event_type}' published with data: {data}")

    def _process_events(self):
        while self._running:
            event = None
            with self._lock:
                if self._event_queue:
                    event = self._event_queue.pop(0) # FIFO
            
            if event:
                event_type = event['type']
                data = event['data']
                print(f"Processing event '{event_type}'...")
                for handler in self._subscribers[event_type]:
                    try:
                        handler(data) # Pass only data to the handler
                        print(f"  Handler {handler.__name__} processed {event_type} event.")
                    except Exception as e:
                        print(f"  Error processing event '{event_type}' with handler {handler.__name__}: {e}")
            else:
                time.sleep(0.1) # Don't busy-wait

    def start(self):
        if not self._running:
            self._running = True
            self._worker_thread = threading.Thread(target=self._process_events, daemon=True)
            self._worker_thread.start()
            print("EventEngine started.")

    def stop(self):
        self._running = False
        if self._worker_thread:
            self._worker_thread.join(timeout=1) # Give it a moment to finish
        print("EventEngine stopped.")


# --- 3. GatewayRouter ---
class GatewayRouter:
    """
    A simple router for handling internal "API" requests and dispatching them.
    In this modular monolith, it acts as a basic API gateway.
    """
    def __init__(self, service_registry, event_engine):
        self.service_registry = service_registry
        self.event_engine = event_engine
        self._routes = {}
        self.register_route("/health", self._handle_health)
        self.register_route("/api/v1/status", self._handle_status)
        print("GatewayRouter initialized with basic routes.")

    def register_route(self, path, handler):
        self._routes[path] = handler
        print(f"Route '{path}' registered.")

    def handle_request(self, path, method="GET", body=None):
        print(f"GatewayRouter received request: {method} {path}")
        handler = self._routes.get(path)
        if handler:
            return handler(method, body)
        else:
            return {"status": "error", "message": f"Route '{path}' not found"}, 404

    def _handle_health(self, method, body):
        return {"status": "ok", "message": "System is healthy"}, 200

    def _handle_status(self, method, body):
        # A placeholder for more detailed status later
        return {"status": "ok", "system_status": "operational", "version": "v0.1-alpha"}, 200

    def start(self):
        print("GatewayRouter ready to handle requests.")
        # In a real app, this would be a web server loop (e.g., Flask/FastAPI)
        # For now, it's just a conceptual "start"
        print("To test, manually call `router.handle_request()`.")


# --- 4. FolderWatcherService content ---
class FolderWatcherService(FileSystemEventHandler):
    """
    Monitors a specified folder for new files and publishes 'file_dropped' events.
    Utilizes the watchdog library for efficient file system event monitoring.
    """
    def __init__(self, path_to_watch, event_engine):
        super().__init__()
        self.path_to_watch = path_to_watch
        self.event_engine = event_engine
        self.observer = Observer()
        print(f"FolderWatcherService initialized for path: {path_to_watch}")

    def on_created(self, event):
        if not event.is_directory:
            print(f"File created: {event.src_path}")
            self.event_engine.publish("file_dropped", {"file_path": event.src_path})

    def start(self):
        if not os.path.exists(self.path_to_watch):
            os.makedirs(self.path_to_watch)
            print(f"Created missing watch folder: {self.path_to_watch}")

        self.observer.schedule(self, self.path_to_watch, recursive=False)
        self.observer.start()
        print(f"FolderWatcherService started monitoring '{self.path_to_watch}' in a background thread.")

    def stop(self):
        self.observer.stop()
        self.observer.join()
        print("FolderWatcherService stopped.")

# --- 5. DocDigestService content (updated) ---
class DocDigestService:
    """
    Processes 'file_dropped' events, digests file info,
    creates a summary, and moves the original file.
    """
    def __init__(self, storage_root):
        self.storage_root = storage_root
        self.inbox_path = os.path.join(storage_root, "inbox")
        self.processed_path = os.path.join(storage_root, "processed")
        self.outbox_path = os.path.join(storage_root, "outbox")
        self._ensure_storage_dirs()
        print(f"DocDigestService initialized with storage root: {storage_root}")

    def _ensure_storage_dirs(self):
        os.makedirs(self.inbox_path, exist_ok=True)
        os.makedirs(self.processed_path, exist_ok=True)
        os.makedirs(self.outbox_path, exist_ok=True)
        print(f"Ensured storage directories exist: {self.inbox_path}, {self.processed_path}, {self.outbox_path}")

    def handle_file_dropped(self, data):
        file_path = data.get('file_path') or data.get('path') # Robustness for both keys
        if not file_path:
            print("Error: No 'file_path' or 'path' provided in file_dropped event data.")
            return

        # Simple check to avoid processing watchdog internal events or duplicate events
        # This can happen if a file is moved or copied with specific OS behaviors
        if not os.path.exists(file_path):
            print(f"Warning: File '{file_path}' not found, possibly already moved or deleted before processing.")
            return
        
        # Avoid processing files that are still being written to (simple heuristic)
        # In a real system, you might need a more robust file locking or completion detection.
        try:
            with open(file_path, 'rb') as f:
                pass # Try to open and immediately close to check for access
        except IOError:
            print(f"Warning: File '{file_path}' is currently inaccessible, skipping for now. Will retry if event re-occurs.")
            return


        try:
            # 1. Read basic file data
            file_name = os.path.basename(file_path)
            file_size_bytes = os.path.getsize(file_path)
            file_size_kb = round(file_size_bytes / 1024, 2)
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            print(f"Processing file: {file_name} (Size: {file_size_kb} KB)")

            # 2. Create digest file in storage/outbox
            digest_file_name = f"DIGEST_{file_name}.txt"
            digest_file_path = os.path.join(self.outbox_path, digest_file_name)

            digest_content = (
                f"--- Document Digest ---\n"
                f"Original Filename: {file_name}\n"
                f"File Size: {file_size_kb} KB\n"
                f"Processing Timestamp: {timestamp}\n"
                f"--- End Digest ---\n"
            )

            with open(digest_file_path, "w", encoding="utf-8") as f:
                f.write(digest_content)
            print(f"Digest created: {digest_file_path}")

            # 3. Move original file to storage/processed
            destination_path = os.path.join(self.processed_path, file_name)
            os.rename(file_path, destination_path) # os.rename is atomic on POSIX systems
            print(f"Original file moved to: {destination_path}")

        except Exception as e:
            print(f"Error processing file '{file_path}': {e}")


# --- Main Application Setup ---
def main():
    print("AQ-OS Modular Monolith Starting...")

    # 0. Setup base directories
    STORAGE_ROOT = "storage"
    INBOX_PATH = os.path.join(STORAGE_ROOT, "inbox")
    os.makedirs(INBOX_PATH, exist_ok=True)
    print(f"Ensured storage/inbox exists at: {INBOX_PATH}")

    # 1. Initialize core components
    event_engine = EventEngine()
    service_registry = ServiceRegistry()
    gateway_router = GatewayRouter(service_registry, event_engine)

    # Register core components (optional, but good practice for service discovery)
    service_registry.register("EventEngine", event_engine)
    service_registry.register("ServiceRegistry", service_registry)
    service_registry.register("GatewayRouter", gateway_router)


    # 2. Initialize and register application services
    folder_watcher = FolderWatcherService(INBOX_PATH, event_engine)
    doc_digest_service = DocDigestService(STORAGE_ROOT)

    service_registry.register("FolderWatcherService", folder_watcher)
    service_registry.register("DocDigestService", doc_digest_service)

    # 3. Connect services (Event-driven architecture)
    event_engine.subscribe("file_dropped", doc_digest_service.handle_file_dropped)

    # 4. Start core components (EventEngine must start first to process events)
    event_engine.start() # Start the event processing thread
    folder_watcher.start() # Start file watching in its own thread
    gateway_router.start() # Just conceptually "starts" for now

    print("\nAQ-OS Modular Monolith is operational. Waiting for files in storage/inbox...")
    print("To test GatewayRouter routes, try calling:")
    print("  router_response, status_code = gateway_router.handle_request('/health')")
    print("  print(f'Health Check: {router_response} (Status: {status_code})')")


    try:
        # Keep the main thread alive for background services
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nShutting down AQ-OS Modular Monolith...")
    finally:
        folder_watcher.stop()
        event_engine.stop()
        print("AQ-OS Modular Monolith stopped gracefully.")

if __name__ == "__main__":
    main()