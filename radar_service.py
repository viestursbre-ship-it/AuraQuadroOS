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

# IMAP specific imports
import imaplib
import email
import re
from email.header import decode_header
from email.utils import parsedate_to_datetime

try:
    from alert_service import notify_urgent_request
except ImportError:
    notify_urgent_request = None

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

# Corrected BASE_DIR handling to be global and consistent
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
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

def collect_all_folders(root_folder):
    """Rekursīvi savāc visas apakšmapes no Inbox."""
    folders = [root_folder]
    try:
        for sub in root_folder.Folders:
            folders.extend(collect_all_folders(sub))
    except Exception:
        pass
    return folders

def open_email_in_outlook(entry_id):
    if str(entry_id).startswith("imap_"):
        try:
            import tkinter.messagebox as mb
            mb.showinfo("E-pasts (IMAP)", "Šis e-pasts tika nolasīts caur IMAP serveri.\nAtveriet to savā e-pasta programmā.")
        except Exception:
            pass
        return

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
    """Clean and extract text from Outlook MAPI message body."""
    text = ""
    try:
        if hasattr(msg, "Body") and msg.Body:
            text = msg.Body
        elif hasattr(msg, "HTMLBody") and msg.HTMLBody:
            text = re.sub(r'<[^<]+?>', '', msg.HTMLBody)
    except Exception:
        text = ""
    # These replacements seem specific to the original context, keeping them.
    text = text.replace(r"\d", "ī").replace(r"\t", "ī").replace(r"\i", "ī").replace(r"\j", "ā")
    return text.strip()


def _parse_imap_message_content(msg, email_id_raw, keywords, is_warranty_topic, is_self_task):
    """Helper function to extract body and attachments from an IMAP email message."""
    body_text = ""
    saved_audio_path = None
    saved_image_path = None
    saved_pdf_path = None
    
    msg_subj_lower = ""
    if msg['Subject']:
        try:
            subj_decoded = ""
            for segment, charset in decode_header(msg['Subject']):
                if isinstance(segment, bytes):
                    subj_decoded += segment.decode(charset or 'utf-8', errors='ignore')
                else:
                    subj_decoded += segment
            msg_subj_lower = subj_decoded.lower()
        except Exception:
            msg_subj_lower = str(msg['Subject']).lower()

    is_invoice_mail = any(w in msg_subj_lower for w in ["rēķin", "rekin", "invoice", "leasing", "līzing"]) or any(k in msg_subj_lower for k in keywords)

    for part in msg.walk():
        if part.is_multipart():
            continue

        content_type = part.get_content_type()
        charset = part.get_content_charset() or 'utf-8'

        if part.get_filename(): # It's an attachment
            filename = part.get_filename()
            try:
                # Decode filename if it's encoded
                decoded_filename = ""
                for segment, cset in decode_header(filename):
                    if isinstance(segment, bytes):
                        decoded_filename += segment.decode(cset or 'utf-8', errors='ignore')
                    else:
                        decoded_filename += segment
                filename = decoded_filename.strip()
            except Exception:
                pass # Use original filename if decoding fails

            ext = Path(filename).suffix.lower()
            payload = part.get_payload(decode=True)
            if not payload:
                continue

            # A) Audio
            if ext in AUDIO_EXTENSIONS and not saved_audio_path:
                clean_name = f"imap_voice_{email_id_raw.decode()}_{filename}"
                target_path = VOICE_DIR / clean_name
                if not target_path.exists():
                    try:
                        with open(target_path, "wb") as f:
                            f.write(payload)
                        saved_audio_path = str(target_path)
                    except Exception as e:
                        print(f"⚠️ Nevarēja saglabāt balss pielikumu '{filename}': {e}")

            # B) Garantijas attēli
            elif ext in IMAGE_EXTENSIONS and not saved_image_path and (is_warranty_topic or is_self_task):
                # Simple check for size, avoid tiny images
                if len(payload) < 20000:
                    continue
                clean_name = f"imap_warranty_{email_id_raw.decode()}_{filename}"
                target_path = WARRANTY_DIR / clean_name
                if not target_path.exists():
                    try:
                        with open(target_path, "wb") as f:
                            f.write(payload)
                        saved_image_path = str(target_path)
                        print(f"🛡️ [IMAP GARANTIJA SAGLABĀTA]: {saved_image_path}")
                    except Exception as e:
                        print(f"⚠️ Nevarēja saglabāt garantijas attēlu '{filename}': {e}")

            # C) PDF rēķini
            elif ext == ".pdf" and not saved_pdf_path and (is_invoice_mail or "rekins" in filename.lower() or "invoice" in filename.lower()):
                clean_name = f"imap_inv_{email_id_raw.decode()}_{filename}"
                target_path = (INVOICE_DIR / clean_name).resolve()
                if not target_path.exists():
                    try:
                        with open(target_path, "wb") as f:
                            f.write(payload)
                        saved_pdf_path = str(target_path)
                        print(f"✅ [IMAP PDF SAGLABĀTS]: {saved_pdf_path}")
                    except Exception as e:
                        print(f"⚠️ Nevarēja saglabāt PDF pielikumu '{filename}': {e}")
        else: # It's a body part
            if content_type == 'text/plain':
                try:
                    body_text += part.get_payload(decode=True).decode(charset, errors='ignore') + "\n"
                except Exception:
                    pass
            elif content_type == 'text/html':
                try:
                    html_content = part.get_payload(decode=True).decode(charset, errors='ignore')
                    body_text += re.sub(r'<[^<]+?>', '', html_content) + "\n"
                except Exception:
                    pass
    
    # Apply special replacements from original clean_msg_text for body_text
    body_text = body_text.replace(r"\d", "ī").replace(r"\t", "ī").replace(r"\i", "ī").replace(r"\j", "ā")
    return body_text.strip(), saved_audio_path, saved_image_path, saved_pdf_path


def scan_via_imap(items_to_save, keywords, default_keywords, imap_cfg):
    """Skenē e-pastus, izmantojot IMAP protokolu."""
    new_count = 0

    imap_server = imap_cfg.get("server")
    imap_user = imap_cfg.get("user")
    imap_password = imap_cfg.get("password")
    imap_port = imap_cfg.get("port", 993)
    imap_days_back = imap_cfg.get("days_back", 30)
    
    # Use existing fingerprints to avoid duplicates
    existing_fingerprints = set()
    for itm in items_to_save.values():
        s_clean = str(itm.get("subject", "")).strip().lower()
        r_clean = str(itm.get("received", "")).strip()
        if s_clean:
            existing_fingerprints.add((s_clean, r_clean))

    if not (imap_server and imap_user and imap_password):
        print("⚠️ IMAP nav konfigurēts config.json. Pārslēdzos uz manuālo režīmu.")
        return 0

    mail = None
    try:
        mail = imaplib.IMAP4_SSL(imap_server, port=imap_port)
        mail.login(imap_user, imap_password)
        mail.select("INBOX")

        date_since = (datetime.now() - timedelta(days=imap_days_back)).strftime("%d-%b-%Y")
        
        # Search for all emails since date_since. Could add UNSEEN if desired.
        status, email_ids = mail.search(None, 'SINCE', date_since)
        email_id_list = email_ids[0].split()
        
        # Process messages from newest to oldest
        for email_id in reversed(email_id_list):
            try:
                imap_entry_id = f"imap_{email_id.decode()}"
                if imap_entry_id in items_to_save:
                    continue

                status, msg_data = mail.fetch(email_id, '(RFC822)')
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # Decode subject
                subj_decoded = ""
                if msg['Subject']:
                    for segment, charset in decode_header(msg['Subject']):
                        if isinstance(segment, bytes):
                            subj_decoded += segment.decode(charset or 'utf-8', errors='ignore')
                        else:
                            subj_decoded += segment

                subj = subj_decoded.strip() or "(Bez temata)"
                sender_raw = msg['From'] if msg['From'] else "Nezināms"
                sender = sender_raw # Will try to extract more specific sender later if needed

                # Try to extract sender name or email for ignored list matching
                sender_identifier = ""
                try:
                    sender_email_match = re.search(r'<([^>]+)>', sender_raw)
                    if sender_email_match:
                        sender_identifier = sender_email_match.group(1)
                    else: # Fallback to sender_raw if no email in angle brackets
                        sender_identifier = sender_raw
                except Exception:
                    sender_identifier = sender_raw
                
                ignored_list = get_ignored_senders()
                if any(ign in sender_identifier.lower() for ign in ignored_list) or \
                   any(ign in sender.lower() for ign in ignored_list):
                    continue

                rec_time = parsedate_to_datetime(msg['Date'])
                rec_time_str = rec_time.strftime("%d.%m %H:%M") if rec_time else ""

                current_fp = (subj.strip().lower(), rec_time_str)
                if current_fp in existing_fingerprints:
                    continue

                subj_lower = subj.lower()
                sender_lower = sender.lower()

                # No direct equivalent of CurrentUser.Name/Email for IMAP without extra config, keep it simple for now or assume a default
                is_self_task = False 
                if "radar" in subj_lower: # Basic self-task check
                    is_self_task = True

                is_warranty_topic = any(w in subj_lower for w in ["garant", "war", "warranty"])

                body, saved_audio_path, saved_image_path, saved_pdf_path = _parse_imap_message_content(
                    msg, email_id, keywords, is_warranty_topic, is_self_task
                )
                body_lower = body.lower()

                matches_keywords = any(k in subj_lower for k in keywords) or any(k in body_lower[:500] for k in keywords)

                # If matches keywords, is self-task, or found any relevant files
                if matches_keywords or is_self_task or is_warranty_topic or saved_audio_path or saved_image_path or saved_pdf_path:
                    date_match = re.search(r'\b(\d{1,2}[\./]\d{1,2}(?:[\./]\d{2,4})?)\b', subj + " " + body[:300])

                    if saved_audio_path:
                        found_deadline = "Balss apstrāde"
                        cat = "🎙️ Balss ziņa"
                    elif is_warranty_topic or saved_image_path:
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

                    items_to_save[imap_entry_id] = {
                        "subject": subj,
                        "sender": f"{sender} (IMAP)", # Indicate IMAP origin
                        "core_request": body[:120].strip() + "..." if body else subj,
                        "deadline": found_deadline,
                        "category": cat,
                        "status": "ACTIVE",
                        "audio_path": saved_audio_path,
                        "image_path": saved_image_path,
                        "pdf_path": saved_pdf_path,
                        "received": rec_time_str
                    }
                    new_count += 1
                    existing_fingerprints.add(current_fp)
                    print(f"📥 [IMAP JAUNS IERAKSTS]: {cat} | {subj} | No: {sender} (IMAP)")

                    try:
                        if notify_urgent_request and cat in ["Cenu pieprasījums", "🛡️ Garantija", "Kritisks"]:
                            notify_urgent_request(category=cat, subject=subj, sender=sender, deadline=found_deadline)
                    except Exception as alert_e:
                        print(f"⚠️ Kļūda sūtot brīdinājumu uz Telegram: {alert_e}")

            except Exception as e_msg:
                print(f"⚠️ Kļūda apstrādājot IMAP ziņu {email_id.decode()}: {e_msg}")
                continue

    except imaplib.IMAP4.error as imap_err:
        print(f"⚠️ IMAP savienojuma/autorizācijas kļūda: {imap_err}. Pārslēdzos uz manuālo režīmu.")
        return 0
    except Exception as e:
        print(f"⚠️ [IMAP RADARS KĻŪDA]: {e}")
        return 0
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass

    return new_count


def scan_outlook_mailbox():
    new_count = 0

    default_keywords = [
        "cenu piepras", "iepirkums", "termin", "vid", "pasutijum", 
        "līgums", "rekins", "rēķin", "leasing", "līzing", "pied", 
        "garantij", "warranty", "radar"
    ]
    keywords = default_keywords
    days_back = 30

    cfg = {} # Initialize cfg dictionary
    cfg_path = BASE_DIR / "config.json"
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as cf:
                cfg = json.load(cf)
                keywords = cfg.get("keywords", default_keywords)
                days_back = cfg.get("days_back", 30)
        except Exception as e:
            print(f"⚠️ Kļūda nolasot config.json: {e}")
            pass

    items_to_save = {}
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                items_to_save = json.load(f)
        except Exception:
            pass

    # Izveidojam esošo ierakstu pirkstu nospiedumus, lai novērstu dublikātus
    existing_fingerprints = set()
    for itm in items_to_save.values():
        s_clean = str(itm.get("subject", "")).strip().lower()
        r_clean = str(itm.get("received", "")).strip()
        if s_clean:
            existing_fingerprints.add((s_clean, r_clean))
    
    # 1. Režīma nolasīšana no config.json
    radar_mode = cfg.get("radar_mode", "manual").lower()
    if cfg.get("force_imap", False):
        radar_mode = "imap"

    # 2. Režīmu loģika
    if radar_mode == "manual" or radar_mode == "offline":
        print("ℹ️ Radars darbojas manuālajā (offline) režīmā. E-pasts netiek skenēts.")
        return 0

    elif radar_mode == "imap" or radar_mode == "gmail":
        print(f"📡 Radars darbojas IMAP/Gmail režīmā.")
        new_count = scan_via_imap(items_to_save, keywords, default_keywords, cfg.get("imap", {}))
        if new_count > 0:
            try:
                with open(STATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(items_to_save, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"⚠️ Kļūda saglabājot stāvokļa failu pēc IMAP skenēšanas: {e}")
        return new_count

    elif radar_mode == "outlook":
        print(f"📡 Radars darbojas Outlook (MAPI) režīmā.")
        outlook = None
        if not HAS_WIN32COM:
            print("⚠️ Outlook MAPI modulis (win32com) nav pieejams. Pārslēdzos uz manuālo režīmu.")
            return 0
        
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        except Exception as mapi_err:
            print(f"⚠️ Classic Outlook MAPI nav pieejams ({mapi_err}). Pārslēdzos uz manuālo režīmu.")
            return 0 # Fail gracefully and switch to manual
    else:
        print(f"⚠️ Neatpazīts radara režīms '{radar_mode}'. Pārslēdzos uz manuālo režīmu.")
        return 0


    # If we reached here, it means radar_mode is "outlook" and MAPI is available.
    try:
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

        # Savācam visas mapes no visiem kontiem
        all_target_folders = []
        try:
            for store in outlook.Folders:
                all_target_folders.extend(collect_all_folders(store))
        except Exception:
            try:
                inbox = outlook.GetDefaultFolder(6)
                all_target_folders = collect_all_folders(inbox)
            except Exception:
                all_target_folders = []

        # Datuma filtrs
        start_date = (datetime.now() - timedelta(days=days_back)).strftime("%d/%m/%Y")
        filter_query = f"[ReceivedTime] >= '{start_date}'"

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
                    sender = getattr(msg, "SenderName", getattr(msg, "SenderEmailAddress", "Sistēmas Ziņa")) or "Nezināms"
                    body = clean_msg_text(msg)

                    if any(ign in sender.lower() for ign in ignored_list):
                        continue

                    subj_lower = subj.lower()
                    body_lower = body.lower()
                    sender_lower = sender.lower()

                    # Saņemšanas laiks un dublikātu pārbaude
                    rec_time_str = msg.ReceivedTime.strftime("%d.%m %H:%M") if hasattr(msg, "ReceivedTime") and msg.ReceivedTime else ""
                    current_fp = (subj.strip().lower(), rec_time_str)

                    if current_fp in existing_fingerprints:
                        continue  # Izlaižam dublikātu, ja tas jau reģistrēts no citas mapes

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

                    is_warranty_topic = any(w in subj_lower for w in ["garant", "war", "warranty"])

                    saved_audio_path = None
                    saved_image_path = None
                    saved_pdf_path = None

                    # Pielikumu nolasīšana
                    if hasattr(msg, "Attachments") and msg.Attachments.Count > 0:
                        is_invoice_mail = any(w in subj_lower for w in ["rēķin", "rekin", "invoice", "leasing", "līzing"]) or any(k in subj_lower for k in keywords)

                        for att_idx in range(1, msg.Attachments.Count + 1):
                            try:
                                att = msg.Attachments.Item(att_idx)
                                att_name = att.FileName
                                ext = Path(att_name).suffix.lower()

                                # A) Audio
                                if ext in AUDIO_EXTENSIONS and not saved_audio_path:
                                    clean_name = f"voice_{att_name}"
                                    target_path = VOICE_DIR / clean_name
                                    if not target_path.exists():
                                        att.SaveAsFile(str(target_path))
                                    saved_audio_path = str(target_path)

                                # B) Garantijas attēli
                                elif ext in IMAGE_EXTENSIONS and not saved_image_path and (is_warranty_topic or is_self_task):
                                    if hasattr(att, "Size") and att.Size < 20000:
                                        continue
                                    clean_name = f"warranty_{att_name}"
                                    target_path = WARRANTY_DIR / clean_name
                                    if not target_path.exists():
                                        att.SaveAsFile(str(target_path))
                                    saved_image_path = str(target_path)
                                    print(f"🛡️ [GARANTIJA SAGLABĀTA]: {saved_image_path}")

                                # C) PDF rēķini
                                elif ext == ".pdf" and not saved_pdf_path and (is_invoice_mail or "rekins" in att_name.lower() or "invoice" in att_name.lower()):
                                    clean_name = f"inv_{att_name}"
                                    target_path = (INVOICE_DIR / clean_name).resolve()
                                    if not target_path.exists():
                                        att.SaveAsFile(str(target_path))
                                    saved_pdf_path = str(target_path)
                                    print(f"✅ [PDF SAGLABĀTS]: {saved_pdf_path}")

                            except Exception:
                                pass

                    matches_keywords = any(k in subj_lower for k in keywords)

                    # Ja atbilst atslēgvārdiem, ir paša uzdevums, vai atrasts kāds no failiem
                    if matches_keywords or is_self_task or is_warranty_topic or saved_audio_path or saved_image_path or saved_pdf_path:
                        date_match = re.search(r'\b(\d{1,2}[\./]\d{1,2}(?:[\./]\d{2,4})?)\b', subj + " " + body[:300])

                        if saved_audio_path:
                            found_deadline = "Balss apstrāde"
                            cat = "🎙️ Balss ziņa"
                        elif is_warranty_topic or saved_image_path:
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
                            "received": rec_time_str
                        }
                        new_count += 1
                        existing_fingerprints.add(current_fp)
                        print(f"📥 [JAUNS IERAKSTS]: {cat} | {subj} | No: {sender} ({folder_name})")

                        # Trauksmes zvans uz Telegram
                        try:
                            if notify_urgent_request and cat in ["Cenu pieprasījums", "🛡️ Garantija", "Kritisks"]:
                                notify_urgent_request(category=cat, subject=subj, sender=sender, deadline=found_deadline)
                        except Exception as alert_e:
                            print(f"⚠️ Kļūda sūtot brīdinājumu uz Telegram: {alert_e}")

                except Exception as e_msg:
                    print(f"⚠️ Kļūda apstrādājot Outlook ziņu '{subj}' no {sender}: {e_msg}")
                    continue

        if new_count > 0:
            with open(STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(items_to_save, f, ensure_ascii=False, indent=2)

    except Exception as e:
        print(f"⚠️ [RADARS KĻŪDA]: {e}")
        return 0 # Ensure 0 is returned on error
    finally:
        if HAS_WIN32COM and 'outlook' in locals() and outlook is not None: # Only call if win32com was actually initialized
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass # Already uninitialized or error

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

        # Poga: Reakcijas audits (Follow-Up Hub)
        self.btn_stale = ctk.CTkButton(
            btn_box, 
            text="⏳", 
            width=30, 
            height=26, 
            fg_color="#334155", 
            hover_color="#d97706", 
            command=self.open_followup_hub
        )
        self.btn_stale.pack(side="left", padx=2)
        CTkToolTip(self.btn_stale, "Neaktīvo klientu un uzdevumu reakcijas audits (> 3 d.)")

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

    def draft_smart_reply(self, entry_id, card_data, btn_widget):
        """Asinhroni ar Gemini uzģenerē atbildes melnrakstu un atver to Outlook."""
        if str(entry_id).startswith("imap_"):
            import tkinter.messagebox as messagebox
            messagebox.showinfo("Smart Reply (IMAP)", "Viedā atbilde nav pieejama IMAP e-pastiem, jo nav tiešas integrācijas ar jūsu e-pasta klientu. Jums būs jākopē atbildes melnraksts manuāli.")
            return

        if not HAS_WIN32COM:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Kļūda", "Outlook modulis (win32com) nav pieejams!")
            return

        # Vizuālais indikators uz pogas
        orig_text = btn_widget.cget("text")
        btn_widget.configure(text="⏳ Raksta...", state="disabled")

        def _worker():
            try:
                # 1. Sagatavojam kontekstu Gemini
                subj = card_data.get("subject", "")
                sender = card_data.get("sender", "")
                body_snippet = card_data.get("core_request", "")

                api_key = os.environ.get("GEMINI_API_KEY", "") # Fallback to env variable
                if not api_key:
                    # Mēģinām paņemt no config.json
                    cfg = {}
                    cfg_path = BASE_DIR / "config.json"
                    if cfg_path.exists():
                        try:
                            with open(cfg_path, "r", encoding="utf-8") as cf:
                                cfg = json.load(cf)
                        except Exception:
                            pass
                    api_key = cfg.get("gemini_api_key", "")

                reply_text = ""
                if api_key:
                    import google.generativeai as genai
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel("gemini-2.5-flash")
                    prompt = f"""Tu esi profesionāls, atsaucīgs un lietišķs partneris uzņēmuma komunikācijā.
                    Sagatavo kodolīgu, pieklājīgu un profesionālu e-pasta atbildes melnrakstu latviešu valodā uz šādu vēstuli.
                    Sūtītājs: {sender}
                    Temats: {subj}
                    Saturs / konteksts:
                    {body_snippet}

                    Norādījumi:
                    - Esi pieklājīgs, lietišķs un pozitīvs.
                    - Apstiprini saņemšanu un norādi, ka jautājums tiek/tiks operatīvi izskatīts.
                    - Neizmanto liekus ievadvārdus vai paskaidrojumus — sniedz TIKAI gatavu e-pasta tekstu (ar uzrunu un nobeiguma sveicienu)."""
                    response = model.generate_content(prompt)
                    reply_text = response.text.strip()
                else:
                    reply_text = "Labdien!\n\nPaldies par informāciju un vēstuli. Esam to saņēmuši un tuvākajā laikā sazināsimies ar Jums.\n\nAr cieņu,\nKomanda"

                # 2. Pieslēdzamies Outlook MAPI caur COM šajā pavedienā
                pythoncom.CoInitialize()
                try:
                    outlook_app = win32com.client.Dispatch("Outlook.Application")
                    namespace = outlook_app.GetNamespace("MAPI")
                    mail_item = namespace.GetItemFromID(entry_id)

                    # Izveidojam atbildi
                    reply_item = mail_item.Reply()
                    # Ievietojam uzģenerēto tekstu pirms oriģinālās vēstules
                    reply_item.Body = reply_text + "\n\n" + reply_item.Body
                    
                    # Saglabājam pie melnrakstiem un parādām lietotājam logu
                    reply_item.Save()
                    reply_item.Display()
                finally:
                    pythoncom.CoUninitialize()

            except Exception as e:
                print(f"⚠️ [Smart Reply Kļūda]: {e}")
                self.after(0, lambda: messagebox.showerror("Smart Reply Kļūda", f"Neizdevās izveidot melnrakstu:\n{e}"))
            finally:
                self.after(0, lambda: btn_widget.configure(text=orig_text, state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    def open_kvark_briefing(self):
        try:
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
                    if self.scan_tick >= 300: # Scan every 5 minutes (300 seconds)
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
                # The format in items_to_save is "%d.%m %H:%M". We need a year to parse it for sorting.
                # Assuming current year if no year is present.
                year = datetime.now().year
                return datetime.strptime(f"{val}.{year}", "%d.%m %H:%M.%Y")
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
            btn_reply_txt = "✍️ Reply" if is_en else "✍️ Atbilde"
            btn_voice_txt = "🎙️ Voice" if is_en else "🎙️ Balss"
            btn_warranty_txt = "🛡️ Warranty" if is_en else "🛡️ Garantija"
            btn_bank_txt = "🏦 Bank" if is_en else "🏦 Rēķins"
            btn_offer_txt = "💼 Quote" if is_en else "💼 Piedāvājums"
            btn_done_txt = "✓ Done" if is_en else "✓ Nokārtots"
            btn_dismiss_txt = "🚫 Dismiss" if is_en else "🚫 Ignorēt"

            # 1. Poga "Atvērt e-pastu Outlook"
            btn_open_mail = ctk.CTkButton(
                btn_row, text=btn_open_txt, width=50, height=24, fg_color="#2563eb", hover_color="#1d4ed8", font=ctk.CTkFont(size=11),
                command=lambda eid=entry_id: open_email_in_outlook(eid)
            )
            btn_open_mail.pack(side="left", padx=(0, 2))
            CTkToolTip(btn_open_mail, "Atvērt oriģinālo e-pastu Outlook")

            # 1.1 Viedā atbildes poga (Smart Reply)
            btn_smart_reply = ctk.CTkButton(
                btn_row, text=btn_reply_txt, width=65, height=24, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=11, weight="bold")
            )
            btn_smart_reply.configure(
                command=lambda eid=entry_id, cdata=item_data, b=btn_smart_reply: self.draft_smart_reply(eid, cdata, b)
            )
            btn_smart_reply.pack(side="left", padx=(0, 2))
            CTkToolTip(btn_smart_reply, "Ar Gemini sagatavot pieklājīgu atbildes melnrakstu Outlook")

            # Saturs un nosacījumi
            audio_p = item_data.get("audio_path")
            img_p = item_data.get("image_path", "")
            pdf_p = item_data.get("pdf_path") or item_data.get("attachment_path", "")
            subj_l = str(item_data.get("subject", "")).lower()

            is_strict_warranty = "warranty" in subj_l or "garantij" in subj_l
            is_inv = bool(pdf_p) or any(w in subj_l for w in ["rēķin", "rekin", "invoice", "leasing", "līzing"])

            # 2. Viedā moduļa poga
            if audio_p:
                btn_mod = ctk.CTkButton(btn_row, text=btn_voice_txt, width=60, height=24, fg_color="#8b5cf6", hover_color="#7c3aed", font=ctk.CTkFont(size=11, weight="bold"),
                                        command=lambda path=audio_p: self.trigger_voice_flow(path))
                btn_mod.pack(side="left", padx=(0, 2))
                CTkToolTip(btn_mod, "Pāriet uz balss ieraksta atšifrēšanu")
            elif is_strict_warranty:
                target_p = img_p or pdf_p or ""
                btn_mod = ctk.CTkButton(btn_row, text=btn_warranty_txt, width=70, height=24, fg_color="#059669", hover_color="#047857", font=ctk.CTkFont(size=11, weight="bold"),
                                        command=lambda path=target_p: self.trigger_warranty_flow(path))
                btn_mod.pack(side="left", padx=(0, 2))
                CTkToolTip(btn_mod, "Reģistrēt garantijas dokumentu")
            elif is_inv:
                btn_mod = ctk.CTkButton(btn_row, text=btn_bank_txt, width=60, height=24, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=11, weight="bold"),
                                        command=lambda path=pdf_p: self.trigger_bank_flow(path))
                btn_mod.pack(side="left", padx=(0, 2))
                CTkToolTip(btn_mod, "Atvērt Bank Studio un sagatavot maksājumu")
            else:
                btn_mod = ctk.CTkButton(btn_row, text=btn_offer_txt, width=65, height=24, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=11),
                                        command=lambda item=item_data: self.trigger_offer_flow(item))
                btn_mod.pack(side="left", padx=(0, 2))
                CTkToolTip(btn_mod, "Atvērt Quote Studio cenu piedāvājuma izveidei")

            # 3. Labās puses pārvaldības pogas
            btn_dismiss = ctk.CTkButton(btn_row, text=btn_dismiss_txt, width=54, height=24, fg_color="#475569", hover_color="#334155", font=ctk.CTkFont(size=11),
                                        command=lambda s=item_data.get("sender", ""), eid=entry_id, c=card: self.ignore_sender_and_remove(s, eid, c))
            btn_dismiss.pack(side="right", padx=(2, 0))
            CTkToolTip(btn_dismiss, "Paslēpt uzdevumu un ignorēt šo sūtītāju nākotnē")

            btn_done = ctk.CTkButton(btn_row, text=btn_done_txt, width=54, height=24, fg_color="#10b981", hover_color="#059669", font=ctk.CTkFont(size=11),
                                     command=lambda eid=entry_id, c=card: self.mark_as_done(eid, c))
            btn_done.pack(side="right", padx=(2, 0))
            CTkToolTip(btn_done, "Atzīmēt uzdevumu kā izpildītu un pārvietot uz arhīvu")

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

    def open_followup_hub(self):
        try:
            from followup_hub import FollowUpDialog
            FollowUpDialog(self)
        except Exception as e:
            import tkinter.messagebox as mb
            mb.showerror("Kļūda", f"Nevar atvērt audita moduli: {e}")

if __name__ == "__main__":
    app = RadarMainApp()
    app.mainloop()