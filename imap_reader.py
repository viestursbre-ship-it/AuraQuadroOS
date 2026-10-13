# -*- coding: utf-8 -*-
import imaplib
import email
from email.header import decode_header
import os
import sys
import json
from pathlib import Path

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

def load_config():
    cfg_path = get_base_dir() / "config.json"
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def decode_mime_words(raw_header):
    """Atkodē MIME hederus (piem. UTF-8 latviešu burtus)."""
    if not raw_header:
        return ""
    decoded_fragments = decode_header(raw_header)
    result = []
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            if encoding:
                try:
                    result.append(fragment.decode(encoding, errors="replace"))
                except LookupError:
                    result.append(fragment.decode("utf-8", errors="replace"))
            else:
                result.append(fragment.decode("utf-8", errors="replace"))
        else:
            result.append(str(fragment))
    return "".join(result)

def fetch_emails_via_imap(limit=15):
    """
    Nolasa pēdējos e-pastus no IMAP servera un saglabā pielikumus (PDF) storage/invoices/
    Atgriež sarakstu ar vārdnīcām, kas saskan ar radara datu struktūru.
    """
    cfg = load_config()
    imap_cfg = cfg.get("imap", {})

    server = imap_cfg.get("server", "outlook.office365.com")
    port = imap_cfg.get("port", 993)
    user = imap_cfg.get("user", "")
    password = imap_cfg.get("password", "")

    if not user or not password:
        print("⚠️ [IMAP]: Nav norādīts lietotājs vai parole config.json failā.")
        return []

    base_dir = get_base_dir()
    inv_dir = base_dir / "storage" / "invoices"
    inv_dir.mkdir(parents=True, exist_ok=True)

    email_items = []

    try:
        mail = imaplib.IMAP4_SSL(server, port)
        mail.login(user, password)
        mail.select("INBOX")

        # Meklējam visas vai jaunākās vēstules
        status, messages = mail.search(None, "ALL")
        if status != "OK":
            return []

        msg_ids = messages[0].split()
        if not msg_ids:
            return []

        # Ņemam pēdējās vēstules
        latest_ids = msg_ids[-limit:]
        latest_ids.reverse()

        for mid in latest_ids:
            res, data = mail.fetch(mid, "(RFC822)")
            if res != "OK":
                continue

            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_mime_words(msg.get("Subject", "(Bez temata)"))
            sender = decode_mime_words(msg.get("From", "Nezināms"))
            date_str = msg.get("Date", "")

            body = ""
            pdf_path = ""

            # Apstrādājam saturu un pielikumus
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))

                    # Pārbaudām pielikumus (PDF)
                    filename = part.get_filename()
                    if filename:
                        filename = decode_mime_words(filename)
                        if filename.lower().endswith(".pdf"):
                            safe_name = f"inv_{mid.decode('ascii')}_{filename}"
                            target_file = inv_dir / safe_name
                            try:
                                with open(target_file, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                pdf_path = str(target_file)
                            except Exception as e:
                                print(f"⚠️ Kļūda saglabājot IMAP PDF: {e}")

                    # Pārbaudām tekstu
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        charset = part.get_content_charset() or "utf-8"
                        try:
                            body = part.get_payload(decode=True).decode(charset, errors="replace")
                        except Exception:
                            pass
            else:
                charset = msg.get_content_charset() or "utf-8"
                try:
                    body = msg.get_payload(decode=True).decode(charset, errors="replace")
                except Exception:
                    body = str(msg.get_payload())

            # Formatējam vienotajā Radar struktūrā
            email_items.append({
                "entry_id": f"imap_{mid.decode('ascii')}",
                "subject": subject,
                "sender": sender,
                "received": date_str[:16] if date_str else "",
                "core_request": body[:500] if body else subject,
                "pdf_path": pdf_path,
                "source": "IMAP"
            })

        mail.close()
        mail.logout()

    except Exception as e:
        print(f"⚠️ [IMAP Kļūda]: {e}")

    return email_items
