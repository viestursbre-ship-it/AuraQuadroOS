# -*- coding: utf-8 -*-
import os
import sys
import json
import threading
from datetime import datetime, timedelta
from pathlib import Path
import customtkinter as ctk

try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False


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


def check_stale_records(days_limit=3):
    """Pārbauda radar_state.json un atrod neaktīvos ierakstus virs sliekšņa."""
    base_dir = get_base_dir()
    state_file = base_dir / "radar_state.json"
    stale_list = []

    if not state_file.exists():
        return stale_list

    try:
        with open(state_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        now = datetime.now()
        threshold = now - timedelta(days=days_limit)

        for entry_id, item in data.items():
            if item.get("status") != "ACTIVE":
                continue

            # --- FILTRS: Izlaižam rēķinus, vebinārus un automatizētos pasta paziņojumus ---
            subj_lower = str(item.get("subject", "")).lower()
            cat_lower = str(item.get("category", "")).lower()
            sender_lower = str(item.get("sender", "")).lower()

            # Nevajag sūtīt follow-up uz rēķiniem, vebināru ielūgumiem vai robotu e-pastiem
            ignore_keywords = [
                "vebinār", "webinar", "rēķin", "invoice", "līzing", "leasing", 
                "pasta burgers", "jauns likums", "no-reply", "noreply", "info nodokļu"
            ]
            if any(kw in subj_lower or kw in cat_lower or kw in sender_lower for kw in ignore_keywords):
                continue
            # --------------------------------------------------------------------------

            rec_str = item.get("received", "").strip()
            if not rec_str:
                continue

            try:
                # Parsējam DD.MM HH:MM
                item_dt = datetime.strptime(f"{rec_str} {now.year}", "%d.%m %H:%M %Y")
                if item_dt > now:
                    item_dt = datetime.strptime(f"{rec_str} {now.year - 1}", "%d.%m %H:%M %Y")

                if item_dt < threshold:
                    days_waiting = (now - item_dt).days
                    stale_list.append({
                        "entry_id": entry_id,
                        "subject": item.get("subject", "(Bez temata)"),
                        "sender": item.get("sender", "Nezināms"),
                        "category": item.get("category", "Uzdevums"),
                        "core_request": item.get("core_request", ""),
                        "days": days_waiting,
                        "received": rec_str
                    })
            except Exception:
                pass

    except Exception as e:
        print(f"⚠️ [FollowUp Hub Kļūda]: {e}")

    # Sakārtojam: visilgāk gaidošie augšā
    stale_list.sort(key=lambda x: x["days"], reverse=True)
    return stale_list


class FollowUpDialog(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("⏳ Klientu Reakcijas Audits (Follow-Up Hub)")
        self.geometry("720x520")
        self.minsize(580, 400)
        self.configure(fg_color="#0f172a")

        header = ctk.CTkFrame(self, fg_color="#1e293b", corner_radius=0, height=50)
        header.pack(fill="x")

        ctk.CTkLabel(
            header,
            text="⏳ Neaktīvie e-pasti un klienti (gaida atbildi > 3 dienām)",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#f59e0b"
        ).pack(side="left", padx=20, pady=12)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=20, pady=15)

        self.render_items()

    def render_items(self):
        for w in self.scroll.winfo_children():
            w.destroy()

        items = check_stale_records(days_limit=3)

        if not items:
            lbl = ctk.CTkLabel(
                self.scroll,
                text="✨ Visi procesi rit nevainojami!\nNav neviena e-pasta vai klienta, kas gaidītu ilgāk par 3 dienām.",
                font=ctk.CTkFont(size=13),
                text_color="#10b981"
            )
            lbl.pack(pady=50)
            return

        for it in items:
            card = ctk.CTkFrame(self.scroll, fg_color="#1e293b", corner_radius=8)
            card.pack(fill="x", pady=5, padx=2)

            info_box = ctk.CTkFrame(card, fg_color="transparent")
            info_box.pack(side="left", fill="both", expand=True, padx=12, pady=10)

            title_txt = f"[{it['category']}] {it['subject']}"
            if len(title_txt) > 65:
                title_txt = title_txt[:62] + "..."

            ctk.CTkLabel(
                info_box,
                text=title_txt,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#f8fafc",
                anchor="w"
            ).pack(fill="x")

            ctk.CTkLabel(
                info_box,
                text=f"👤 {it['sender']} • 🕒 Saņemts: {it['received']} • ⚠️ Gaida: {it['days']} d.",
                font=ctk.CTkFont(size=11),
                text_color="#94a3b8",
                anchor="w"
            ).pack(fill="x")

            # Poga: Follow-Up melnraksts
            btn_act = ctk.CTkButton(
                card,
                text="✉️ Follow-Up",
                width=85,
                height=26,
                fg_color="#0284c7",
                hover_color="#0369a1",
                font=ctk.CTkFont(size=11, weight="bold")
            )
            btn_act.configure(command=lambda itm=it, b=btn_act: self.send_followup(itm, b))
            btn_act.pack(side="right", padx=12)

    def send_followup(self, item, btn_widget):
        if not HAS_WIN32COM:
            return

        btn_widget.configure(text="⏳...", state="disabled")

        def _worker():
            try:
                cfg = load_config()
                api_key = cfg.get("gemini_api_key") or os.environ.get("GEMINI_API_KEY", "")
                draft_body = ""

                if api_key:
                    import google.generativeai as genai
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel("gemini-2.5-flash")
                    prompt = f"""Tu esi profesionāls partneris.
Sagatavo ļoti pieklājīgu, maigu un lietišķu 'follow-up' atgādinājuma e-pastu latviešu valodā.
Vēstules temats: {item['subject']}
Sūtītājs: {item['sender']}
Sākotnējais pieprasījums: {item['core_request']}

Norāde: Pavaicā, vai ir izdevies iepazīties ar informāciju un vai varam sniegt kādas papildu atbildes. Tikai tīrs vēstules teksts."""
                    res = model.generate_content(prompt)
                    draft_body = res.text.strip()
                else:
                    draft_body = f"Labdien!\n\nAtgādinām par mūsu saziņu saistībā ar '{item['subject']}'. Vai esat paspējuši izskatīt informāciju?\n\nAr cieņu,\nKomanda"

                pythoncom.CoInitialize()
                try:
                    outlook_app = win32com.client.Dispatch("Outlook.Application")
                    ns = outlook_app.GetNamespace("MAPI")
                    mail = ns.GetItemFromID(item["entry_id"])
                    reply = mail.Reply()
                    reply.Body = draft_body + "\n\n" + reply.Body
                    reply.Save()
                    reply.Display()
                finally:
                    pythoncom.CoUninitialize()

            except Exception as e:
                print(f"⚠️ Kļūda ģenerējot follow-up: {e}")
            finally:
                self.after(0, lambda: btn_widget.configure(text="✉️ Follow-Up", state="normal"))

        threading.Thread(target=_worker, daemon=True).start()
