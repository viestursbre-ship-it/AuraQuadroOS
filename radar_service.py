# -*- coding: utf-8 -*-
import os
import sys
import json
import threading
from pathlib import Path
from datetime import datetime, timedelta
import customtkinter as ctk
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from cockpit import CockpitWindow

BASE_DIR = Path(__file__).resolve().parent
STATE_FILE = BASE_DIR / "radar_state.json"
IGNORE_FILE = BASE_DIR / "radar_ignored_senders.json"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def get_ignored_senders():
    if not os.path.exists(IGNORE_FILE):
        return set()
    try:
        with open(IGNORE_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    except Exception:
        return set()

def add_ignored_sender(sender_identifier):
    if not sender_identifier:
        return
    ignored = get_ignored_senders()
    ignored.add(sender_identifier.strip().lower())
    try:
        with open(IGNORE_FILE, "w", encoding="utf-8") as f:
            json.dump(list(ignored), f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def open_email_in_outlook(entry_id):
    if not HAS_WIN32COM:
        return
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        msg = outlook.GetItemFromID(entry_id)
        msg.Display()
    except Exception as e:
        print(f"⚠️ Nevarēja atvērt e-pastu: {e}")
    finally:
        pythoncom.CoUninitialize()

def clean_msg_text(msg):
    text = ""
    try:
        if hasattr(msg, "Body") and msg.Body:
            text = msg.Body
        elif hasattr(msg, "HTMLBody") and msg.HTMLBody:
            import re
            text = re.sub(r'<[^<]+?>', '', msg.HTMLBody)
    except Exception:
        text = ""
    text = text.replace(r"\d", "ī").replace(r"\t", "ī").replace(r"\i", "ī").replace(r"\j", "ā")
    return text.strip()

def collect_all_folders(root_folder):
    folders = [root_folder]
    try:
        for sub in root_folder.Folders:
            folders.extend(collect_all_folders(sub))
    except Exception:
        pass
    return folders

def scan_outlook_mailbox():
    if not HAS_WIN32COM:
        return 0

    pythoncom.CoInitialize()
    new_count = 0
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        
        current_user_name = ""
        current_user_email = ""
        try:
            current_user_name = outlook.CurrentUser.Name.lower()
        except Exception:
            pass

        try:
            current_user_email = outlook.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
        except Exception:
            try:
                current_user_email = outlook.CurrentUser.Address.lower()
            except Exception:
                pass

        default_keywords = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin"]
        keywords = default_keywords
        days_back = 30
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as cf:
                    cfg = json.load(cf)
                    keywords = cfg.get("keywords", default_keywords)
                    days_back = cfg.get("days_back", 30)
            except Exception:
                pass

        all_target_folders = collect_all_folders(inbox)
        start_date = (datetime.now() - timedelta(days=days_back)).strftime("%d/%m/%Y")
        filter_query = f"[ReceivedTime] >= '{start_date}'"

        items_to_save = {}
        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, "r", encoding="utf-8") as f:
                    items_to_save = json.load(f)
            except Exception:
                pass

        ignored_list = get_ignored_senders()

        for folder in all_target_folders:
            folder_name = getattr(folder, "Name", "Nezināma")
            try:
                messages = folder.Items
                filtered_messages = messages.Restrict(filter_query)
                filtered_messages.Sort("[ReceivedTime]", True)
                total = filtered_messages.Count
            except Exception:
                continue

            for i in range(1, total + 1):
                try:
                    msg = filtered_messages.Item(i)
                    entry_id = msg.EntryID
                    if entry_id in items_to_save:
                        continue

                    subj = msg.Subject or "(Bez temata)"
                    sender = msg.SenderName or "Nezināms"
                    body = clean_msg_text(msg)

                    if any(ign in sender.lower() for ign in ignored_list):
                        continue

                    subj_lower = subj.lower()
                    body_lower = body.lower()
                    sender_lower = sender.lower()

                    is_self_task = False
                    if current_user_name and current_user_name in sender_lower:
                        is_self_task = True
                    elif current_user_email and current_user_email in sender_lower:
                        is_self_task = True
                    elif "outlook for android" in body_lower or "outlook for ios" in body_lower:
                        if "līdz rītdienai" in body_lower or "radar" in subj_lower or "rīt" in body_lower:
                            is_self_task = True
                    elif "radar" in subj_lower:
                        is_self_task = True

                    matches_keywords = any(k in subj_lower for k in keywords)

                    if matches_keywords or is_self_task:
                        import re
                        date_match = re.search(r'\b(\d{1,2}[\./]\d{1,2}(?:[\./]\d{2,4})?)\b', subj + " " + body[:300])
                        
                        if date_match:
                            found_deadline = f"Līdz {date_match.group(1)}"
                        elif "rīt" in body_lower or "rītdien" in body_lower:
                            found_deadline = "Līdz rītdienai"
                        else:
                            found_deadline = "Steidzams"

                        cat = "Zibens uzdevums" if is_self_task else ("Cenu pieprasījums" if "cenu" in subj_lower else "Kritisks")

                        items_to_save[entry_id] = {
                            "subject": subj,
                            "sender": f"{sender} ({folder_name})",
                            "core_request": body[:120].strip() + "..." if body else subj,
                            "deadline": found_deadline,
                            "category": cat,
                            "status": "ACTIVE",
                            "received": msg.ReceivedTime.strftime("%d.%m %H:%M") if hasattr(msg, "ReceivedTime") else ""
                        }
                        new_count += 1
                except Exception:
                    continue

        if new_count > 0:
            with open(STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(items_to_save, f, ensure_ascii=False, indent=2)

    except Exception as e:
        print(f"⚠️ [RADARS KĻŪDA]: {e}")
    finally:
        pythoncom.CoUninitialize()
    return new_count

# ==========================================
# GALVENĀ RADARA LIETOTNE
# ==========================================
class RadarMainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("⚡ AQ-OS Radar")
        self.geometry("480x680")
        self.resizable(False, False)

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"480x680+{sw - 500}+{sh - 760}")

        self.protocol("WM_DELETE_WINDOW", self.hide_window)

        # Cockpit logs sākotnēji nav atvērts
        self.cockpit_win = None

        # Augšējā josla
        header = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=0)
        header.pack(fill="x", padx=0, pady=0)

        ctk.CTkLabel(
            header, 
            text="📡 AQ RADARS", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            text_color="#38bdf8"
        ).pack(side="left", padx=12, pady=10)

        btn_box = ctk.CTkFrame(header, fg_color="transparent")
        btn_box.pack(side="right", padx=10)

        ctk.CTkButton(
            btn_box, 
            text="🎛️ Vadības Centrs", 
            width=110, 
            height=26, 
            fg_color="#0284c7", 
            hover_color="#0369a1",
            font=ctk.CTkFont(size=11), 
            command=self.open_cockpit
        ).pack(side="left", padx=3)

        self.btn_scan = ctk.CTkButton(
            btn_box, 
            text="📥 Skenēt", 
            width=70, 
            height=26, 
            fg_color="#2563eb", 
            hover_color="#1d4ed8", 
            font=ctk.CTkFont(size=11), 
            command=self.trigger_scan
        )
        self.btn_scan.pack(side="left", padx=3)

        ctk.CTkButton(
            btn_box, 
            text="🔄", 
            width=30, 
            height=26, 
            fg_color="#334155", 
            command=self.load_cards
        ).pack(side="left", padx=2)

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=450, height=580, fg_color="#0b1329")
        self.scroll_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.load_cards()
        self.start_background_watcher()
        self.start_tray_icon()

    def hide_window(self):
        self.withdraw()

    def show_window(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self.load_cards()

    def open_cockpit(self):
        if self.cockpit_win is None or not self.cockpit_win.winfo_exists():
            self.cockpit_win = CockpitWindow(self)
        else:
            self.cockpit_win.deiconify()
            self.cockpit_win.lift()
            self.cockpit_win.focus_force()

    def start_background_watcher(self, interval_sec=300):
        def _loop():
            while True:
                threading.Event().wait(interval_sec)
                try:
                    added = scan_outlook_mailbox()
                    if added > 0:
                        self.after(0, self.load_cards)
                except Exception:
                    pass
        threading.Thread(target=_loop, daemon=True).start()

    def trigger_scan(self):
        self.btn_scan.configure(state="disabled")
        def _task():
            scan_outlook_mailbox()
            self.after(0, lambda: self.btn_scan.configure(state="normal"))
            self.after(0, self.load_cards)
        threading.Thread(target=_task, daemon=True).start()

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

    def ignore_sender_and_remove(self, sender, entry_id, card_widget):
        add_ignored_sender(sender)
        self.mark_as_done(entry_id, card_widget)

    def load_cards(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()

        if not os.path.exists(STATE_FILE):
            ctk.CTkLabel(self.scroll_frame, text="Nav aktīvu datu.", text_color="#94a3b8").pack(pady=40)
            return

        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}

        ignored_list = get_ignored_senders()
        active_items = [
            (k, v) for k, v in data.items() 
            if v.get("status") == "ACTIVE" and not any(ign in v.get("sender", "").lower() for ign in ignored_list)
        ]

        def parse_date(item_tuple):
            val = item_tuple[1].get("received", "")
            try:
                return datetime.strptime(f"2026.{val}", "%Y.%d.%m %H:%M")
            except Exception:
                return datetime.min

        active_items.sort(key=parse_date, reverse=True)

        if not active_items:
            ctk.CTkLabel(self.scroll_frame, text="✨ Visi uzdevumi nokārtoti!", font=ctk.CTkFont(size=14), text_color="#10b981").pack(pady=40)
            return

        for entry_id, item_data in active_items:
            card = ctk.CTkFrame(self.scroll_frame, corner_radius=8, fg_color="#131b2e")
            card.pack(pady=6, padx=4, fill="x")

            cat = item_data.get("category", "Cits")
            badge_bg = "#ef4444" if cat in ["Līgums", "Kritisks"] else "#f59e0b" if "pieprasījums" in cat.lower() else "#3b82f6"

            top_bar = ctk.CTkFrame(card, fg_color="transparent")
            top_bar.pack(fill="x", padx=10, pady=(8, 2))
            ctk.CTkLabel(top_bar, text=f" {cat} ", fg_color=badge_bg, corner_radius=4, text_color="white", font=ctk.CTkFont(size=10, weight="bold")).pack(side="left")

            dl = item_data.get("deadline")
            if dl:
                ctk.CTkLabel(top_bar, text=f" ⏳ {dl} ", fg_color="#dc2626", corner_radius=4, text_color="white", font=ctk.CTkFont(size=10, weight="bold")).pack(side="left", padx=(6, 0))

            if item_data.get("received"):
                ctk.CTkLabel(top_bar, text=item_data.get("received"), text_color="#64748b", font=ctk.CTkFont(size=10)).pack(side="right")

            ctk.CTkLabel(card, text=item_data.get("subject", ""), font=ctk.CTkFont(size=12, weight="bold"), wraplength=410, justify="left").pack(anchor="w", padx=10, pady=(2, 0))
            ctk.CTkLabel(card, text=f"No: {item_data.get('sender', '')}", font=ctk.CTkFont(size=11), text_color="#94a3b8", wraplength=410, justify="left").pack(anchor="w", padx=10)

            req = item_data.get("core_request")
            if req:
                ctk.CTkLabel(card, text=req, font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=410, justify="left").pack(anchor="w", padx=10, pady=4)

            btn_row = ctk.CTkFrame(card, fg_color="transparent")
            btn_row.pack(fill="x", padx=10, pady=(4, 8))
            
            ctk.CTkButton(btn_row, text="↗ Atvērt", width=75, height=24, fg_color="#2563eb", hover_color="#1d4ed8", font=ctk.CTkFont(size=11),
                          command=lambda eid=entry_id: open_email_in_outlook(eid)).pack(side="left", padx=(0, 6))
            
            ctk.CTkButton(btn_row, text="✓ Nokārtots", width=85, height=24, fg_color="#059669", hover_color="#047857", font=ctk.CTkFont(size=11),
                          command=lambda eid=entry_id, c=card: self.mark_as_done(eid, c)).pack(side="left", padx=(0, 6))

            ctk.CTkButton(btn_row, text="🚫 Ignorēt", width=75, height=24, fg_color="#334155", hover_color="#dc2626", font=ctk.CTkFont(size=11),
                          command=lambda s=item_data.get("sender", ""), eid=entry_id, c=card: self.ignore_sender_and_remove(s, eid, c)).pack(side="left")

    def quit_completely(self):
        if hasattr(self, "tray_icon") and self.tray_icon:
            self.tray_icon.stop()
        self.destroy()
        sys.exit(0)

    def start_tray_icon(self):
        img = Image.new("RGBA", (64, 64), color=(0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.ellipse((6, 6, 58, 58), fill="#f59e0b")

        def on_open_radar(icon, itm):
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass

        def on_open_cockpit(icon, itm):
            try:
                self.open_cockpit()
            except Exception:
                pass

        def on_quit(icon, itm):
            try:
                icon.stop()
                self.quit()
            except Exception:
                pass

        menu = pystray.Menu(
            item("📡 Atvērt Radaru", on_open_radar, default=True),
            item("🎛️ Vadības Centrs", on_open_cockpit),
            item("Iziet", on_quit)
        )
        self.tray_icon = pystray.Icon("AQ_Cockpit", img, "AQ-OS Radar", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

if __name__ == "__main__":
    app = RadarMainApp()
    app.mainloop()
