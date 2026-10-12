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
from daily_briefing import generate_daily_report

class CTkToolTip:
    def __init__(self, widget, text, delay_ms=400):
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self.tip_window = None
        self.schedule_id = None
        self.widget.bind("<Enter>", self._on_enter, add="+")
        self.widget.bind("<Leave>", self._on_leave, add="+")
        self.widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, event=None):
        self._cancel_schedule()
        self.schedule_id = self.widget.after(self.delay_ms, self._show_tip)

    def _on_leave(self, event=None):
        self._cancel_schedule()
        self._hide_tip()

    def _cancel_schedule(self):
        if self.schedule_id:
            self.widget.after_cancel(self.schedule_id)
            self.schedule_id = None

    def _show_tip(self):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip_window = tw = ctk.CTkToplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)
        ctk.CTkLabel(
            tw, text=self.text, fg_color="#334155", text_color="#f8fafc",
            corner_radius=6, padx=8, pady=4, font=("Segoe UI", 11, "bold")
        ).pack()

    def _hide_tip(self):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

BASE_DIR = Path(__file__).resolve().parent
STATE_FILE = BASE_DIR / "radar_state.json"
IGNORE_FILE = BASE_DIR / "radar_ignored_senders.json"
VOICE_DIR = BASE_DIR / "storage" / "voice"
INVOICE_DIR = BASE_DIR / "storage" / "invoices"
WARRANTY_DIR = BASE_DIR / "storage" / "warranty"

VOICE_DIR.mkdir(parents=True, exist_ok=True)
INVOICE_DIR.mkdir(parents=True, exist_ok=True)
WARRANTY_DIR.mkdir(parents=True, exist_ok=True)

AUDIO_EXTENSIONS = {".ogg", ".opus", ".mp3", ".m4a", ".wav", ".aac"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}

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
    new_count = 0

    if not HAS_WIN32COM:
        try:
            import tkinter.messagebox as mb
            mb.showerror("AQ Radar Kļūda", "Outlook MAPI modulis (win32com) nav ielādēts!")
        except Exception:
            pass
        return 0

    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
        except Exception as mapi_err:
            try:
                import tkinter.messagebox as mb
                mb.showerror("Outlook Savienojums", f"Nevar pieslēgties Outlook:\n{mapi_err}")
            except Exception:
                pass
            return 0

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

        default_keywords = ["cenu piepras", "iepirkums", "termin", "vid", "pasutijum", "līgums", "rekins", "rēķin", "leasing", "līzing", "pied"]
        keywords = default_keywords
        days_back = 30

        # Nosakām patieso bāzes mapi
        if getattr(sys, 'frozen', False):
            BASE_DIR = Path(sys.executable).parent
        else:
            BASE_DIR = Path(__file__).resolve().parent

        STORAGE_DIR = BASE_DIR / "storage"
        INVOICE_DIR = STORAGE_DIR / "invoices"
        VOICE_DIR = STORAGE_DIR / "voice"
        WARRANTY_DIR = STORAGE_DIR / "warranty"

        INVOICE_DIR.mkdir(parents=True, exist_ok=True)
        VOICE_DIR.mkdir(parents=True, exist_ok=True)
        WARRANTY_DIR.mkdir(parents=True, exist_ok=True)

        cfg_path = BASE_DIR / "config.json"

        if cfg_path.exists():
            try:
                with open(cfg_path, "r", encoding="utf-8") as cf:
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
                        if "līdz rītdienai" in body_lower or "radar" in subj_lower or "rīt" in body_lower or "garantij" in subj_lower or "warranty" in subj_lower:
                            is_self_task = True
                    elif "radar" in subj_lower:
                        is_self_task = True

                    saved_audio_path = None
                    saved_image_path = None
                    saved_pdf_path = None

                    # Pielikumu nolasīšana
                    if hasattr(msg, "Attachments") and msg.Attachments.Count > 0:
                        is_strict_warranty_mail = is_self_task and any(w in subj_lower for w in ["garant", "war"])
                        is_invoice_mail = any(w in subj_lower for w in ["rēķin", "rekin", "invoice", "leasing", "līzing"]) or any(k in subj_lower for k in keywords)

                        for att_idx in range(1, msg.Attachments.Count + 1):
                            try:
                                att = msg.Attachments.Item(att_idx)
                                att_name = att.FileName
                                ext = Path(att_name).suffix.lower()

                                if ext in AUDIO_EXTENSIONS and not saved_audio_path:
                                    clean_name = f"voice_{att_name}"
                                    target_path = VOICE_DIR / clean_name
                                    if not target_path.exists():
                                        att.SaveAsFile(str(target_path))
                                    saved_audio_path = str(target_path)

                                elif ext in IMAGE_EXTENSIONS and not saved_image_path and is_strict_warranty_mail:
                                    if hasattr(att, "Size") and att.Size < 20000:
                                        continue
                                    clean_name = f"warranty_{att_name}"
                                    target_path = WARRANTY_DIR / clean_name
                                    if not target_path.exists():
                                        att.SaveAsFile(str(target_path))
                                    saved_image_path = str(target_path)

                                # C) PDF rēķini
                                elif ext == ".pdf" and not saved_pdf_path and (is_invoice_mail or "rekins" in att_name.lower() or "invoice" in att_name.lower()):
                                    try:
                                        INVOICE_DIR.mkdir(parents=True, exist_ok=True)
                                        clean_name = f"inv_{att_name}"
                                        target_path = (INVOICE_DIR / clean_name).resolve()
                                        att.SaveAsFile(str(target_path))
                                        saved_pdf_path = str(target_path)
                                        print(f"✅ [PDF SAGLABĀTS]: {saved_pdf_path}")
                                    except Exception as save_err:
                                        print(f"⚠️ [KĻŪDA SAGLABĀJOT PDF]: {save_err}")

                            except Exception:
                                pass

                    matches_keywords = any(k in subj_lower for k in keywords)

                    if matches_keywords or is_self_task or saved_audio_path or saved_image_path or saved_pdf_path:
                        import re
                        date_match = re.search(r'\b(\d{1,2}[\./]\d{1,2}(?:[\./]\d{2,4})?)\b', subj + " " + body[:300])

                        if saved_audio_path:
                            found_deadline = "Balss apstrāde"
                            cat = "🎙️ Balss ziņa"
                        elif "garantij" in subj_lower or "warranty" in subj_lower:
                            found_deadline = "Steidzams"
                            cat = "🛡️ Garantija"
                        elif any(w in subj_lower for w in ["rēķin", "rekin", "invoice", "leasing", "līzing"]) or saved_pdf_path:
                            found_deadline = f"Līdz {date_match.group(1)}" if date_match else "Apmaksai"
                            cat = "Rēķins"
                        elif date_match:
                            found_deadline = f"Līdz {date_match.group(1)}"
                            cat = "Zibens uzdevums" if is_self_task else ("Cenu pieprasījums" if "cenu" in subj_lower else "Kritisks")
                        elif "rīt" in body_lower or "rītdien" in body_lower:
                            found_deadline = "Līdz rītdienai"
                            cat = "Zibens uzdevums" if is_self_task else ("Cenu pieprasījums" if "cenu" in subj_lower else "Kritisks")
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
                            "audio_path": saved_audio_path,
                            "image_path": saved_image_path,
                            "pdf_path": saved_pdf_path,
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

        self.current_lang = "LV"
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    self.current_lang = json.load(f).get("ui_language", "LV")
            except Exception:
                pass
        
        self.cockpit_win = None

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

        btn_cp_txt = "🎛️ Control Center" if self.current_lang == "EN" else "🎛️ Vadības Centrs"
        btn_sc_txt = "🔄 Scan Now" if self.current_lang == "EN" else "📥 Skenēt"

        self.btn_cockpit = ctk.CTkButton(
            btn_box,
            text=btn_cp_txt,
            width=110,
            height=26,
            fg_color="#0284c7",
            hover_color="#0369a1",
            font=ctk.CTkFont(size=11),
            command=self.open_cockpit
        )
        self.btn_cockpit.pack(side="left", padx=3)
        CTkToolTip(self.btn_cockpit, "Atvērt AQ-OS Vadības Centru (Cockpit)")

        self.btn_scan = ctk.CTkButton(
            btn_box, 
            text=btn_sc_txt, 
            width=80, 
            height=26, 
            fg_color="#2563eb", 
            hover_color="#1d4ed8", 
            font=ctk.CTkFont(size=11), 
            command=self.trigger_scan
        )
        self.btn_scan.pack(side="left", padx=3)
        CTkToolTip(self.btn_scan, "Skenēt Outlook iesūtni tūlīt")

        # --- ŠĪ IR TAVA POGA: '🔄' PĀRTOPOŠA PAR '☕' KVARKU ---
        self.btn_briefing = ctk.CTkButton(
            btn_box, 
            text="☕", 
            width=30, 
            height=26, 
            fg_color="#334155", 
            hover_color="#d97706",
            command=self.open_kvark_briefing
        )
        self.btn_briefing.pack(side="left", padx=2)
        CTkToolTip(self.btn_briefing, "Komandanta Kvarka Rīta Brīfings")

        # Žurnāla/pārskata poga
        self.btn_log = ctk.CTkButton(
            btn_box, 
            text="📋", 
            width=30, 
            height=26, 
            fg_color="#475569", 
            hover_color="#334155", 
            command=self.trigger_briefing
        )
        self.btn_log.pack(side="left", padx=2)
        CTkToolTip(self.btn_log, "Vakara atskaite / arhīvs")

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=450, height=580, fg_color="#0b1329")
        self.scroll_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.load_cards()
        self.start_background_watcher()
        self.start_tray_icon()

    def set_ui_language(self, lang: str):
        self.current_lang = lang
        if lang == "EN":
            if hasattr(self, "btn_cockpit"):
                self.btn_cockpit.configure(text="🎛️ Control Center")
            if hasattr(self, "btn_scan"):
                self.btn_scan.configure(text="🔄 Scan Now")
        else:
            if hasattr(self, "btn_cockpit"):
                self.btn_cockpit.configure(text="🎛️ Vadības Centrs")
            if hasattr(self, "btn_scan"):
                self.btn_scan.configure(text="📥 Skenēt")
        self.load_cards()

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

    def open_kvark_briefing(self):
        try:
            # Ja nosauci failu par morning.py, tad: from morning import KvarkBriefingDialog
            from morning_briefing import KvarkBriefingDialog
            KvarkBriefingDialog(self)
        except Exception as e:
            import tkinter.messagebox as mb
            mb.showerror("Kļūda", f"Nevar atvērt Kvarka brīfingu: {e}")

    def trigger_offer_flow(self, item_data):
        self.open_cockpit()
        if self.cockpit_win:
            self.cockpit_win.open_for_offer(
                subject=item_data.get("subject", ""),
                sender=item_data.get("sender", ""),
                body=item_data.get("core_request", "")
            )

    def trigger_bank_flow(self, pdf_path):
        """Atver Cockpit un pārslēdz uz Bank Studio cilni ar PDF failu."""
        self.open_cockpit()
        if self.cockpit_win:
            try:
                tab_name = "🏦 Bank Studio"
                if hasattr(self.cockpit_win, "UI_TEXTS") and hasattr(self.cockpit_win, "current_lang"):
                    t = self.cockpit_win.UI_TEXTS.get(self.cockpit_win.current_lang, {})
                    tab_name = t.get("tab_bank", "🏦 Bank Studio")
                self.cockpit_win.tabview.set(tab_name)
            except Exception:
                try:
                    self.cockpit_win.tabview.set("🏦 Bank Studio")
                except Exception:
                    pass

            if hasattr(self.cockpit_win, "bank_module") and pdf_path:
                self.cockpit_win.bank_module.process_pdf_direct(pdf_path)

    def trigger_voice_flow(self, audio_path):
        self.open_cockpit()
        if self.cockpit_win:
            self.cockpit_win.open_for_voice(audio_path)

    def trigger_warranty_flow(self, image_path):
        self.open_cockpit()
        if self.cockpit_win:
            self.cockpit_win.open_for_warranty(image_path)

    def show_briefing_popup(self, report_text):
        win = ctk.CTkToplevel(self)
        win.title("📋 Vakara Komandanta Atskaite" if self.current_lang == "LV" else "📋 Evening Briefing")
        win.geometry("520x460")
        win.attributes("-topmost", True)
        win.configure(fg_color="#0b1329")

        header = ctk.CTkFrame(win, fg_color="#0f172a", height=45)
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="🐾 KVARKA KOMANDANTA ATSKAITE" if self.current_lang == "LV" else "🐾 QUARK COMMANDANT BRIEFING",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#f59e0b"
        ).pack(side="left", padx=15, pady=10)

        txt = ctk.CTkTextbox(win, fg_color="#131b2e", font=ctk.CTkFont(size=12), wrap="word")
        txt.pack(fill="both", expand=True, padx=15, pady=15)
        txt.insert("end", report_text)
        txt.configure(state="disabled")

        btn_ack = "Sapratu! (Pieņemt zināšanai)" if self.current_lang == "LV" else "Acknowledged!"
        ctk.CTkButton(
            win,
            text=btn_ack,
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=win.destroy
        ).pack(pady=(0, 15))

    def trigger_briefing(self):
        report = generate_daily_report()
        if report:
            self.show_briefing_popup(report)

    def start_background_watcher(self):
        self.last_briefing_date = None
        self.scan_tick = 0

        def _loop():
            while True:
                threading.Event().wait(5)
                try:
                    self.scan_tick += 5
                    if self.scan_tick >= 300:
                        self.scan_tick = 0
                        added = scan_outlook_mailbox()
                        if added > 0:
                            self.after(0, self.load_cards)

                    cfg_time = "18:00"
                    if os.path.exists("config.json"):
                        try:
                            with open("config.json", "r", encoding="utf-8") as cf:
                                cfg_time = json.load(cf).get("briefing_time", "18:00")
                        except Exception:
                            pass

                    target_hour, target_min = map(int, cfg_time.split(":"))
                    now = datetime.now()
                    today = now.date()

                    if now.hour == target_hour and now.minute == target_min and self.last_briefing_date != today:
                        self.last_briefing_date = today
                        report = generate_daily_report()
                        if report:
                            self.after(0, lambda r=report: self.show_briefing_popup(r))
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
        # 1. Notīrām iepriekšējos logrīkus un atbrīvojam grafiskos resursus
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.update_idletasks()

        is_en = getattr(self, "current_lang", "LV") == "EN"

        if not os.path.exists(STATE_FILE):
            msg = "No active tasks." if is_en else "Nav aktīvu datu."
            ctk.CTkLabel(self.scroll_frame, text=msg, text_color="#94a3b8").pack(pady=40)
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

        # Ierobežojam rādāmo skaitu līdz 20, lai Windows nepārsniegtu User/GDI rokturu limitu
        active_items = active_items[:20]

        if not active_items:
            empty_msg = "✨ All tasks completed!" if is_en else "✨ Visi uzdevumi nokārtoti!"
            ctk.CTkLabel(self.scroll_frame, text=empty_msg, font=ctk.CTkFont(size=14), text_color="#10b981").pack(pady=40)
            return

        for entry_id, item_data in active_items:
            card = ctk.CTkFrame(self.scroll_frame, corner_radius=8, fg_color="#131b2e")
            card.pack(pady=6, padx=4, fill="x")

            cat = item_data.get("category", "Cits")
            cat_display = cat
            cat_lower = cat.lower()
            if "balss" in cat_lower or "voice" in cat_lower:
                cat_display = "Voice Memo" if is_en else "Balss ziņa"
                badge_bg = "#8b5cf6"
            elif "garant" in cat_lower or "war" in cat_lower:
                cat_display = "Warranty" if is_en else "Garantija"
                badge_bg = "#059669"
            elif "rēķin" in cat_lower or "rekin" in cat_lower or "leas" in cat_lower or "līzing" in cat_lower:
                cat_display = "Invoice" if is_en else "Rēķins"
                badge_bg = "#0284c7"
            elif "kritisk" in cat_lower or "crit" in cat_lower:
                cat_display = "Critical" if is_en else "Kritisks"
                badge_bg = "#ef4444"
            elif "līgum" in cat_lower or "contract" in cat_lower:
                cat_display = "Contract" if is_en else "Līgums"
                badge_bg = "#ef4444"
            elif "piepras" in cat_lower or "inquiry" in cat_lower:
                cat_display = "Inquiry" if is_en else "Pieprasījums"
                badge_bg = "#f59e0b"
            else:
                cat_display = "Info" if is_en else cat
                badge_bg = "#3b82f6"

            # Augšējā josla ar statusa birkām
            top_bar = ctk.CTkFrame(card, fg_color="transparent")
            top_bar.pack(fill="x", padx=10, pady=(8, 2))
            ctk.CTkLabel(top_bar, text=f" {cat_display} ", fg_color=badge_bg, corner_radius=4, text_color="white", font=ctk.CTkFont(size=10, weight="bold")).pack(side="left")

            dl = item_data.get("deadline")
            if dl:
                dl_prefix = "⏳ Due: " if is_en else "⏳ Līdz: "
                ctk.CTkLabel(top_bar, text=f" {dl_prefix}{dl} ", fg_color="#dc2626", corner_radius=4, text_color="white", font=ctk.CTkFont(size=10, weight="bold")).pack(side="left", padx=(6, 0))

            if item_data.get("received"):
                ctk.CTkLabel(top_bar, text=item_data.get("received"), text_color="#64748b", font=ctk.CTkFont(size=10)).pack(side="right")

            # Galvenā informācija
            ctk.CTkLabel(card, text=item_data.get("subject", ""), font=ctk.CTkFont(size=12, weight="bold"), wraplength=410, justify="left").pack(anchor="w", padx=10, pady=(2, 0))
            
            sender_prefix = "From: " if is_en else "No: "
            ctk.CTkLabel(card, text=f"{sender_prefix}{item_data.get('sender', '')}", font=ctk.CTkFont(size=11), text_color="#94a3b8", wraplength=410, justify="left").pack(anchor="w", padx=10)

            req = item_data.get("core_request")
            if req:
                ctk.CTkLabel(card, text=req, font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=410, justify="left").pack(anchor="w", padx=10, pady=4)

            # Pogu rinda
            btn_row = ctk.CTkFrame(card, fg_color="transparent")
            btn_row.pack(fill="x", padx=10, pady=(4, 8))
            
            btn_open_txt = "↗ Open" if is_en else "↗ Atvērt"
            btn_voice_txt = "🎙️ Voice" if is_en else "🎙️ Balss"
            btn_warranty_txt = "🛡️ Warranty" if is_en else "🛡️ Garantija"
            btn_bank_txt = "🏦 Bank" if is_en else "🏦 Rēķins"
            btn_offer_txt = "💼 Quote" if is_en else "💼 Piedāvājums"
            btn_done_txt = "✓ Done" if is_en else "✓ Nokārtots"
            btn_dismiss_txt = "🚫 Dismiss" if is_en else "🚫 Ignorēt"

            # 1. Poga "Atvērt e-pastu Outlook"
            ctk.CTkButton(btn_row, text=btn_open_txt, width=54, height=24, fg_color="#2563eb", hover_color="#1d4ed8", font=ctk.CTkFont(size=11),
                          command=lambda eid=entry_id: open_email_in_outlook(eid)).pack(side="left", padx=(0, 2))

            # Saturs un nosacījumi
            audio_p = item_data.get("audio_path")
            img_p = item_data.get("image_path", "")
            pdf_p = item_data.get("pdf_path") or item_data.get("attachment_path", "")
            subj_l = str(item_data.get("subject", "")).lower()

            is_strict_warranty = "warranty" in subj_l or "garantij" in subj_l
            is_inv = bool(pdf_p) or any(w in subj_l for w in ["rēķin", "rekin", "invoice", "leasing", "līzing"])

            # 2. Viedā moduļa poga
            if audio_p:
                ctk.CTkButton(btn_row, text=btn_voice_txt, width=64, height=24, fg_color="#8b5cf6", hover_color="#7c3aed", font=ctk.CTkFont(size=11, weight="bold"),
                              command=lambda path=audio_p: self.trigger_voice_flow(path)).pack(side="left", padx=(0, 2))
            elif is_strict_warranty:
                target_p = img_p or pdf_p or ""
                ctk.CTkButton(btn_row, text=btn_warranty_txt, width=74, height=24, fg_color="#059669", hover_color="#047857", font=ctk.CTkFont(size=11, weight="bold"),
                              command=lambda path=target_p: self.trigger_warranty_flow(path)).pack(side="left", padx=(0, 2))
            elif is_inv:
                ctk.CTkButton(btn_row, text=btn_bank_txt, width=64, height=24, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=11, weight="bold"),
                              command=lambda path=pdf_p: self.trigger_bank_flow(path)).pack(side="left", padx=(0, 2))
            else:
                ctk.CTkButton(btn_row, text=btn_offer_txt, width=70, height=24, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=11),
                              command=lambda item=item_data: self.trigger_offer_flow(item)).pack(side="left", padx=(0, 2))

            # 3. Labās puses pārvaldības pogas
            ctk.CTkButton(btn_row, text=btn_dismiss_txt, width=58, height=24, fg_color="#475569", hover_color="#334155", font=ctk.CTkFont(size=11),
                          command=lambda s=item_data.get("sender", ""), eid=entry_id, c=card: self.ignore_sender_and_remove(s, eid, c)).pack(side="right", padx=(2, 0))

            ctk.CTkButton(btn_row, text=btn_done_txt, width=58, height=24, fg_color="#10b981", hover_color="#059669", font=ctk.CTkFont(size=11),
                          command=lambda eid=entry_id, c=card: self.mark_as_done(eid, c)).pack(side="right", padx=(2, 0))

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

        is_en = getattr(self, "current_lang", "LV") == "EN"
        m_radar = "📡 Open Radar" if is_en else "📡 Atvērt Radaru"
        m_cockpit = "🎛️ Control Center" if is_en else "🎛️ Vadības Centrs"
        m_exit = "Exit" if is_en else "Iziet"

        menu = pystray.Menu(
            item(m_radar, on_open_radar, default=True),
            item(m_cockpit, on_open_cockpit),
            item(m_exit, on_quit)
        )
        self.tray_icon = pystray.Icon("AQ_Cockpit", img, "AQ-OS Radar", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

if __name__ == "__main__":
    app = RadarMainApp()
    app.mainloop()
