# -*- coding: utf-8 -*-
import os
import re
import json
from pathlib import Path
import sys
import threading
from datetime import datetime
import xml.etree.ElementTree as ET
from xml.dom import minidom
import customtkinter as ctk
from tkinter import filedialog, messagebox
import google.generativeai as genai

def get_app_config():
    if getattr(sys, 'frozen', False):
        app_dir = Path(sys.executable).parent
    else:
        app_dir = Path(__file__).resolve().parent
    cfg_path = app_dir / "config.json"
    if cfg_path.exists():
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

VENDORS_DB_FILE = "known_vendors.json"

# ==========================================
# KV нейро/DATU BĀZES UN DROŠĪBAS LOĢIKA
# ==========================================
def load_vendors_db():
    if not os.path.exists(VENDORS_DB_FILE):
        # Sākotnējā bāze ar tipiskiem partneriem
        initial_db = {
            "40003000000": {"name": "ELKO Latvija AS", "ibans": ["LV88PARX0011223344556"]},
            "40003010000": {"name": "ALSO Latvia SIA", "ibans": ["LV55HABA0001112223334"]},
            "40003032949": {"name": "Latvenergo AS", "ibans": ["LV11UNLA0000001111111"]}
        }
        with open(VENDORS_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(initial_db, f, indent=4, ensure_ascii=False)
        return initial_db
    try:
        with open(VENDORS_DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_vendors_db(db):
    try:
        with open(VENDORS_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Kļūda saglabājot partneru bāzi: {e}")

def verify_vendor_security(reg_nr, iban, vendor_name):
    """
    Pārbauda piegādātāju pret zināmo sarakstu:
    🟢 trusted: Reģ. nr. un IBAN sakrīt
    🟡 new: Jauns partneris (jāapstiprina)
    🔴 mismatch: Reģ. nr. zināms, bet IBAN ir mainījies (iespējama krāpšana / Kvarka diversija!)
    """
    db = load_vendors_db()
    clean_reg = re.sub(r'[^A-Za-z0-9]', '', str(reg_nr or "")).upper()
    clean_iban = re.sub(r'\s+', '', str(iban or "")).upper()

    if clean_reg in db:
        known = db[clean_reg]
        known_ibans = [re.sub(r'\s+', '', x).upper() for x in known.get("ibans", [])]
        if clean_iban in known_ibans:
            return "trusted", f"🟢 Pārbaudīts partneris: {known.get('name', vendor_name)}"
        else:
            return "mismatch", f"🔴 ALARM! Partnerim {known.get('name')} mainīts IBAN!\nZināmais: {', '.join(known_ibans)}\nRēķinā: {clean_iban}"
    return "new", f"🟡 Jauns piegādātājs: {vendor_name} (Reģ: {clean_reg})"

# ==========================================
# CITADELE ISO 20022 (pain.001.001.03) ĢENERATORS
# ==========================================
def generate_citadele_xml(invoice_data, my_iban="LV00PARX0000000000000", my_name="Mans Uzņēmums SIA"):
    """Uzģenerē Citadeles bankai derīgu ISO 20022 maksājumu uzdevumu XML."""
    now_str = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    msg_id = f"AQOS-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    root = ET.Element("Document", {
        "xmlns": "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
    })

    cstmr_pmt = ET.SubElement(root, "CstmrCdtTrfInitn")

    # Group Header
    grp_hdr = ET.SubElement(cstmr_pmt, "GrpHdr")
    ET.SubElement(grp_hdr, "MsgId").text = msg_id
    ET.SubElement(grp_hdr, "CreDtTm").text = now_str
    ET.SubElement(grp_hdr, "NbOfTxs").text = "1"
    initg_pty = ET.SubElement(grp_hdr, "InitgPty")
    ET.SubElement(initg_pty, "Nm").text = my_name

    # Payment Information
    pmt_inf = ET.SubElement(cstmr_pmt, "PmtInf")
    ET.SubElement(pmt_inf, "PmtInfId").text = f"PMT-{msg_id}"
    ET.SubElement(pmt_inf, "PmtMtd").text = "TRF"
    ET.SubElement(pmt_inf, "NbOfTxs").text = "1"

    pmt_tp_inf = ET.SubElement(pmt_inf, "PmtTpInf")
    svc_lvl = ET.SubElement(pmt_tp_inf, "SvcLvl")
    ET.SubElement(svc_lvl, "Cd").text = "SEPA"

    reqd_exctn_dt = ET.SubElement(pmt_inf, "ReqdExctnDt")
    reqd_exctn_dt.text = invoice_data.get("due_date") or datetime.now().strftime("%Y-%m-%d")

    # Debtor (Maksātājs)
    dbtr = ET.SubElement(pmt_inf, "Dbtr")
    ET.SubElement(dbtr, "Nm").text = my_name
    dbtr_acct = ET.SubElement(pmt_inf, "DbtrAcct")
    dbtr_id = ET.SubElement(dbtr_acct, "Id")
    ET.SubElement(dbtr_id, "IBAN").text = re.sub(r'\s+', '', my_iban)

    dbtr_agt = ET.SubElement(pmt_inf, "DbtrAgt")
    fin_inst = ET.SubElement(dbtr_agt, "FinInstnId")
    ET.SubElement(fin_inst, "BIC").text = "PARXLV22"  # Citadele BIC

    # Credit Transfer Transaction (Saņēmējs un Maksājums)
    tx_inf = ET.SubElement(pmt_inf, "CdtTrfTxInf")
    pmt_id = ET.SubElement(tx_inf, "PmtId")
    ET.SubElement(pmt_id, "EndToEndId").text = f"INV-{invoice_data.get('invoice_no', 'NO')}"

    amt = ET.SubElement(tx_inf, "Amt")
    instd_amt = ET.SubElement(amt, "InstdAmt", {"Ccy": "EUR"})
    instd_amt.text = f"{float(invoice_data.get('amount', 0.0)):.2f}"

    # Creditor (Piegādātājs)
    cdtr = ET.SubElement(tx_inf, "Cdtr")
    ET.SubElement(cdtr, "Nm").text = invoice_data.get("vendor_name", "Nezināms")

    cdtr_acct = ET.SubElement(tx_inf, "CdtrAcct")
    cdtr_id = ET.SubElement(cdtr_acct, "Id")
    ET.SubElement(cdtr_id, "IBAN").text = re.sub(r'\s+', '', invoice_data.get("iban", ""))

    # RmtInf (Maksājuma mērķis)
    rmt_inf = ET.SubElement(tx_inf, "RmtInf")
    ET.SubElement(rmt_inf, "Ustrd").text = f"Apmaksa par rēķinu {invoice_data.get('invoice_no', '')}"

    xml_str = ET.tostring(root, encoding="utf-8")
    parsed_xml = minidom.parseString(xml_str)
    return parsed_xml.toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")

# ==========================================
# BANK STUDIO UI LOGS
# ==========================================
class BankStudioFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None):
        super().__init__(master, fg_color="transparent")
        self.get_gemini_key_fn = get_gemini_key_fn
        self.current_invoice = None

        # Augšējā josla ar darbībām
        top_bar = ctk.CTkFrame(self, fg_color="#1e293b", corner_radius=8)
        top_bar.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(
            top_bar, 
            text="🏦 Citadele Bank Studio (ISO 20022)", 
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#f8fafc"
        ).pack(side="left", padx=15, pady=10)

        self.btn_select_pdf = ctk.CTkButton(
            top_bar, 
            text="📄 Ielādēt PDF Rēķinu", 
            command=self.load_pdf_invoice,
            fg_color="#2563eb", 
            hover_color="#1d4ed8"
        )
        self.btn_select_pdf.pack(side="right", padx=10, pady=10)

        # Galvenais rāmis ar laukiem
        self.form_frame = ctk.CTkFrame(self, fg_color="#0f172a", corner_radius=8)
        self.form_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Statusa / Kvarka Drošinātāja baneris
        self.lbl_status = ctk.CTkLabel(
            self.form_frame, 
            text="Gatavs darbam. Ielādējiet PDF rēķinu no e-pasta vai datora.", 
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#94a3b8",
            wraplength=600,
            justify="left"
        )
        self.lbl_status.pack(fill="x", padx=20, pady=15)

        # Rēķina detaļu kartīte
        self.details_box = ctk.CTkTextbox(
            self.form_frame, 
            font=ctk.CTkFont(size=13), 
            text_color="#e2e8f0", 
            fg_color="#1e293b",
            height=200
        )
        self.details_box.pack(fill="both", expand=True, padx=20, pady=5)
        self.details_box.configure(state="disabled")

        # Apakšējā pogu josla
        bottom_bar = ctk.CTkFrame(self.form_frame, fg_color="transparent")
        bottom_bar.pack(fill="x", padx=20, pady=15)

        self.btn_whitelist = ctk.CTkButton(
            bottom_bar, 
            text="⭐ Pievienot partneri bāzei", 
            command=self.whitelist_current_vendor,
            state="disabled",
            fg_color="#475569"
        )
        self.btn_whitelist.pack(side="left")

        self.btn_export_xml = ctk.CTkButton(
            bottom_bar, 
            text="🚀 Ģenerēt Citadele XML", 
            command=self.export_xml,
            state="disabled",
            fg_color="#16a34a", 
            hover_color="#15803d",
            font=ctk.CTkFont(weight="bold")
        )
        self.btn_export_xml.pack(side="right")

    def _get_api_key(self):
        if self.get_gemini_key_fn:
            key = self.get_gemini_key_fn()
            if key:
                return key
        return os.environ.get("GEMINI_API_KEY", "")

    def load_pdf_invoice(self):
        file_path = filedialog.askopenfilename(
            title="Izvēlieties rēķina PDF",
            filetypes=[("PDF faili", "*.pdf")]
        )
        if not file_path:
            return

        self.lbl_status.configure(text="⏳ Gemini analizē rēķinu un rekvizītus...", text_color="#38bdf8")
        self.btn_export_xml.configure(state="disabled")
        self.btn_whitelist.configure(state="disabled")

        threading.Thread(target=self._parse_invoice_thread, args=(file_path,), daemon=True).start()

    def _parse_invoice_thread(self, file_path):
        api_key = self._get_api_key()
        if not api_key:
            self._update_status_ui("⚠️ Gemini API atslēga nav atrasta!", "#ef4444")
            return

        try:
            genai.configure(api_key=api_key)

            # 1. Vispirms mēģinām iegūt tekstu no PDF lokāli
            extracted_text = ""
            try:
                import pypdf
                reader = pypdf.PdfReader(file_path)
                for page in reader.pages:
                    extracted_text += (page.extract_text() or "") + "\n"
            except Exception:
                extracted_text = ""

            model = genai.GenerativeModel("gemini-2.5-flash")

            prompt = """Tu esi finanšu un rēķinu grāmatvedības eksperts.
Izanalizē pievienoto rēķinu un atgriez TIKAI derīgu JSON objektu bez jebkādiem ievada vai nobeiguma tekstiem.
Struktūra:
{
    "vendor_name": "Piegādātāja nosaukums",
    "reg_nr": "Reģistrācijas numurs",
    "iban": "IBAN bankas konts",
    "invoice_no": "Rēķina numurs",
    "amount": 0.00,
    "due_date": "YYYY-MM-DD"
}"""

            # Ja teksts ir atrasts, nododam to tieši; ja nē (skenēts PDF), nododam baitu struktūru
            if len(extracted_text.strip()) > 30:
                full_prompt = f"{prompt}\n\nRĒĶINA SATURS:\n{extracted_text}"
                response = model.generate_content(full_prompt)
            else:
                with open(file_path, "rb") as f:
                    pdf_bytes = f.read()
                pdf_blob = {"mime_type": "application/pdf", "data": pdf_bytes}
                response = model.generate_content([prompt, pdf_blob])

            raw_text = response.text.strip() if response and response.text else ""

            # Droša JSON izgūšana no atbildes (atrod pirmo { un pēdējo })
            match = re.search(r'\{.*\}', raw_text, re.DOTALL)
            if not match:
                raise ValueError(f"Modelis neatgrieza derīgu JSON: {raw_text[:120]}")

            clean_json = match.group(0)
            data = json.loads(clean_json)

            if isinstance(data, list) and len(data) > 0:
                data = data[0]

            self.current_invoice = data

            # Drošības pārbaude ("Kvarka Drošinātājs")
            status_code, msg = verify_vendor_security(
                data.get("reg_nr"),
                data.get("iban"),
                data.get("vendor_name")
            )

            color_map = {
                "trusted": "#22c55e",
                "new": "#eab308",
                "mismatch": "#ef4444"
            }

            self.after(0, lambda: self._render_invoice_data(
                data, msg, color_map.get(status_code, "#94a3b8"), status_code
            ))

        except Exception as e:
            self._update_status_ui(f"⚠️ Kļūda nolasot rēķinu: {e}", "#ef4444")

    def _render_invoice_data(self, data, status_msg, color, status_code):
        self.lbl_status.configure(text=status_msg, text_color=color)

        # 1. Droši aprēķinām summu pirms teksta veidošanas
        raw_amt = data.get("amount")
        try:
            amt_val = float(raw_amt) if raw_amt is not None else 0.0
        except (ValueError, TypeError):
            amt_val = 0.0

        # 2. Sagatavojam kopsavilkuma tekstu
        summary_text = (
            f"Piegādātājs:       {data.get('vendor_name')}\n"
            f"Reģistrācijas Nr: {data.get('reg_nr')}\n"
            f"Konta IBAN:       {data.get('iban')}\n"
            f"Rēķina Numurs:    {data.get('invoice_no')}\n"
            f"Summa apmaksai:   {amt_val:.2f} EUR\n"
            f"Apmaksas termiņš: {data.get('due_date')}\n"
        )

        self.details_box.configure(state="normal")
        self.details_box.delete("1.0", "end")
        self.details_box.insert("end", summary_text)
        self.details_box.configure(state="disabled")

        self.btn_export_xml.configure(state="normal")
        if status_code in ["new", "mismatch"]:
            self.btn_whitelist.configure(state="normal", fg_color="#0284c7")
        else:
            self.btn_whitelist.configure(state="disabled", fg_color="#475569")

    def _update_status_ui(self, msg, color):
        self.after(0, lambda: self.lbl_status.configure(text=msg, text_color=color))

    def whitelist_current_vendor(self):
        if not self.current_invoice:
            return
        reg = re.sub(r'[^A-Za-z0-9]', '', str(self.current_invoice.get("reg_nr", ""))).upper()
        iban = re.sub(r'\s+', '', str(self.current_invoice.get("iban", ""))).upper()
        name = self.current_invoice.get("vendor_name", "")

        if not reg or not iban:
            messagebox.showwarning("Brīdinājums", "Trūkst reģ. numura vai IBAN!")
            return

        db = load_vendors_db()
        if reg not in db:
            db[reg] = {"name": name, "ibans": [iban]}
        else:
            if iban not in db[reg]["ibans"]:
                db[reg]["ibans"].append(iban)

        save_vendors_db(db)
        messagebox.showinfo("Veiksmīgi", f"Partneris '{name}' pievienots uzticamajam sarakstam!")
        self.lbl_status.configure(text=f"🟢 Pārbaudīts partneris: {name}", text_color="#22c55e")
        self.btn_whitelist.configure(state="disabled", fg_color="#475569")

    def export_xml(self):
        if not self.current_invoice:
            return

        cfg = get_app_config()
        my_iban = cfg.get("my_iban", "LV00PARX0000000000000").replace(" ", "").strip()
        my_name = cfg.get("my_name", "Mans Uzņēmums SIA").strip()

        xml_content = generate_citadele_xml(self.current_invoice, my_iban=my_iban, my_name=my_name)

        # Attīrām rēķina numuru no slīpsvītrām un citiem Windows aizliegtajiem simboliem:
        raw_inv_no = str(self.current_invoice.get('invoice_no', 'Export'))
        clean_inv_no = re.sub(r'[\\/*?:"<>|]', '_', raw_inv_no)

        default_filename = f"Citadele_Pmt_{clean_inv_no}.xml"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xml",
            initialfile=default_filename,
            filetypes=[("XML faili", "*.xml")]
        )

        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(xml_content)
            messagebox.showinfo("Eksports pabeigts", f"Citadeles ISO 20022 XML fails veiksmīgi saglabāts:\n{save_path}")

    def process_pdf_direct(self, file_path):
        """Metode, ko izsauc Radars, kad nospiež pogu [🏦 Bank]."""
        if not file_path:
            self._update_status_ui("⚠️ Faila ceļš nav norādīts!", "#ef4444")
            return

        resolved_path = Path(file_path)
        
        # Ja tiešais ceļš neeksistē, pameklējam relatīvi pret lietotnes mapi vai storage/invoices
        if not resolved_path.exists():
            base_dir = Path(__file__).resolve().parent
            candidate1 = base_dir / file_path
            candidate2 = base_dir / "storage" / "invoices" / Path(file_path).name
            
            if candidate1.exists():
                resolved_path = candidate1
            elif candidate2.exists():
                resolved_path = candidate2
            else:
                self._update_status_ui(f"⚠️ Fails nav atrasts: {resolved_path.name}", "#ef4444")
                return

        file_path_str = str(resolved_path)
        self.lbl_status.configure(text="⏳ Gemini analizē ienākošo rēķinu no Radara...", text_color="#38bdf8")
        self.btn_export_xml.configure(state="disabled")
        self.btn_whitelist.configure(state="disabled")

        threading.Thread(target=self._parse_invoice_thread, args=(file_path_str,), daemon=True).start()
