# -*- coding: utf-8 -*-
import os
import json
import webbrowser
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image

try:
    from google import genai
    from google.genai import types
    HAS_GENAI_NEW = True
except ImportError:
    HAS_GENAI_NEW = False
    try:
        import google.generativeai as legacy_genai
        HAS_GENAI_OLD = True
    except ImportError:
        HAS_GENAI_OLD = False

# Ražotāju oficiālie garantijas pārbaudes portāli
WARRANTY_URLS = {
    "HP": "https://support.hp.com/lv-lv/check-warranty",
    "DELL": "https://www.dell.com/support/home/product-support/servicetag/{sn}/overview",
    "LENOVO": "https://pcsupport.lenovo.com/lv/en/warranty-lookup",
    "APPLE": "https://checkcoverage.apple.com/",
    "ASUS": "https://www.asus.com/support/warranty-status-inquiry/",
    "GENERIC": "https://www.google.com/search?q={vendor}+{model}+warranty+check"
}

def analyze_device_label_image(image_path: str, api_key: str) -> dict:
    """Nolasa ražotāja uzlīmes datus no foto ar Gemini Vision."""
    if not api_key:
        return {"error": "Trūkst API atslēgas"}
    
    prompt = """Izanalizē šo datora/iekārtas apakšas uzlīmes fotogrāfiju.
Atrodi un precīzi nolasi:
1. vendor: Ražotājs (piemēram, HP, Dell, Lenovo, Apple, Asus u.c.)
2. model: Modeļa nosaukums (piem., EliteBook 840 G8, ThinkPad T14)
3. sn: Sērijas numurs (Serial Number / S/N / Service Tag)
4. pn: Produkta kods (Product Number / P/N / ProdID)

Atgriez TIKAI derīgu JSON šādā formātā:
{
  "vendor": "HP",
  "model": "EliteBook 840 G8",
  "sn": "5CG1234XYZ",
  "pn": "3V2W4EA#ABB"
}
Ja kādu parametru nevari droši noteikt, atstāj "".
"""
    try:
        img = Image.open(image_path)
        raw = ""
        if HAS_GENAI_NEW:
            client = genai.Client(api_key=api_key)
            res = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[img, prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            raw = res.text.strip()
        elif HAS_GENAI_OLD:
            legacy_genai.configure(api_key=api_key)
            model = legacy_genai.GenerativeModel("gemini-2.5-flash")
            res = model.generate_content([img, prompt])
            raw = res.text.strip()
            
        return json.loads(raw)
    except Exception as e:
        return {"error": str(e), "vendor": "", "model": "", "sn": "", "pn": ""}

class WarrantyHubFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None, on_send_to_offer_fn=None):
        super().__init__(master, fg_color="transparent")
        self.get_gemini_key_fn = get_gemini_key_fn
        self.on_send_to_offer_fn = on_send_to_offer_fn
        self.current_lang = "LV"

        # Augšējā vadības josla
        top_bar = ctk.CTkFrame(self, fg_color="#0f172a", height=50)
        top_bar.pack(fill="x", padx=10, pady=(10, 5))

        self.btn_select_img = ctk.CTkButton(
            top_bar, text="📷 Ielādēt Uzlīmes Foto", width=160, height=32,
            fg_color="#2563eb", hover_color="#1d4ed8", font=ctk.CTkFont(size=12, weight="bold"),
            command=self.browse_image
        )
        self.btn_select_img.pack(side="left", padx=10, pady=8)

        self.lbl_status = ctk.CTkLabel(top_bar, text="", font=ctk.CTkFont(size=12), text_color="#38bdf8")
        self.lbl_status.pack(side="left", padx=10)

        # Galvenā satura zona: 2 paneļi
        main_box = ctk.CTkFrame(self, fg_color="transparent")
        main_box.pack(fill="both", expand=True, padx=10, pady=5)

        # Kreisā puse: Attēla priekšskatījums
        self.left_panel = ctk.CTkFrame(main_box, fg_color="#0b1329", width=360)
        self.left_panel.pack(side="left", fill="both", expand=False, padx=(0, 5))

        self.lbl_img_title = ctk.CTkLabel(self.left_panel, text="🖼️ Uzlīmes Attēls:", font=ctk.CTkFont(size=12, weight="bold"), text_color="#94a3b8")
        self.lbl_img_title.pack(anchor="w", padx=12, pady=(10, 5))

        self.lbl_preview = ctk.CTkLabel(self.left_panel, text="Nav izvēlēts attēls\n(Ievelciet vai nospiediet 'Ielādēt Foto')", text_color="#64748b")
        self.lbl_preview.pack(fill="both", expand=True, padx=10, pady=10)

        # Labā puse: Nolasītie dati un darbības
        self.right_panel = ctk.CTkFrame(main_box, fg_color="#0b1329")
        self.right_panel.pack(side="right", fill="both", expand=True, padx=(5, 0))

        self.lbl_info_title = ctk.CTkLabel(self.right_panel, text="🛡️ Iekārtas Identifikācija & Garantija:", font=ctk.CTkFont(size=13, weight="bold"), text_color="#38bdf8")
        self.lbl_info_title.pack(anchor="w", padx=15, pady=(12, 10))

        # Lauki: Ražotājs
        self.create_field(self.right_panel, "Ražotājs (Vendor):", "entry_vendor")
        self.create_field(self.right_panel, "Modelis (Model):", "entry_model")
        self.create_field(self.right_panel, "Sērijas Nr. (S/N):", "entry_sn", with_copy=True)
        self.create_field(self.right_panel, "Produkta Kods (P/N):", "entry_pn", with_copy=True)

        # Garantijas pogu rinda
        act_box = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        act_box.pack(fill="x", padx=15, pady=(20, 10))

        self.btn_open_portal = ctk.CTkButton(
            act_box, text="🌐 Atvērt Garantijas Portālu", height=36,
            fg_color="#059669", hover_color="#047857", font=ctk.CTkFont(size=12, weight="bold"),
            command=self.open_warranty_portal
        )
        self.btn_open_portal.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.btn_to_offer = ctk.CTkButton(
            act_box, text="💼 Sagatavot Piedāvājumu (Upgrade)", height=36,
            fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=12, weight="bold"),
            command=self.send_to_offer_studio
        )
        self.btn_to_offer.pack(side="right", fill="x", expand=True, padx=(6, 0))

    def create_field(self, parent, label_txt, attr_name, with_copy=False):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=15, pady=4)
        lbl = ctk.CTkLabel(row, text=label_txt, width=140, anchor="w", font=ctk.CTkFont(size=11, weight="bold"), text_color="#cbd5e1")
        lbl.pack(side="left")
        entry = ctk.CTkEntry(row, height=28, font=ctk.CTkFont(size=12))
        entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        setattr(self, attr_name, entry)
        if with_copy:
            btn = ctk.CTkButton(row, text="📋", width=32, height=28, fg_color="#334155", hover_color="#475569",
                                command=lambda e=entry: self.copy_to_clip(e.get()))
            btn.pack(side="right")

    def copy_to_clip(self, text):
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.lbl_status.configure(text=f"Nokopēts: {text}")

    def browse_image(self):
        f = filedialog.askopenfilename(filetypes=[("Attēli", "*.jpg *.jpeg *.png"), ("Visi faili", "*.*")])
        if f:
            self.load_and_analyze(f)

    def load_and_analyze(self, image_path):
        self.lbl_status.configure(text="Gemini Vision nolasa uzlīmi...")
        
        # Parāda bildes priekšskatījumu
        try:
            pil_img = Image.open(image_path)
            pil_img.thumbnail((320, 240))
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=pil_img.size)
            self.lbl_preview.configure(image=ctk_img, text="")
        except Exception:
            pass

        import threading
        def _worker():
            api_key = self.get_gemini_key_fn() if self.get_gemini_key_fn else os.environ.get("GEMINI_API_KEY", "")
            data = analyze_device_label_image(image_path, api_key)
            def _done():
                self.set_device_data(
                    data.get("vendor", ""),
                    data.get("model", ""),
                    data.get("sn", ""),
                    data.get("pn", "")
                )
                self.lbl_status.configure(text="✅ Uzlīme sekmīgi nolasīta!")
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    def set_device_data(self, vendor, model, sn, pn):
        self.entry_vendor.delete(0, "end")
        self.entry_vendor.insert(0, vendor)
        self.entry_model.delete(0, "end")
        self.entry_model.insert(0, model)
        self.entry_sn.delete(0, "end")
        self.entry_sn.insert(0, sn)
        self.entry_pn.delete(0, "end")
        self.entry_pn.insert(0, pn)

    def open_warranty_portal(self):
        vendor = self.entry_vendor.get().strip().upper()
        sn = self.entry_sn.get().strip()
        model = self.entry_model.get().strip()

        # Kopējam S/N uzreiz starpliktuvē ērtībai
        if sn:
            self.clipboard_clear()
            self.clipboard_append(sn)

        if "DELL" in vendor and sn:
            url = WARRANTY_URLS["DELL"].format(sn=sn)
        elif "HP" in vendor:
            url = WARRANTY_URLS["HP"]
        elif "LENOVO" in vendor:
            url = WARRANTY_URLS["LENOVO"]
        elif "APPLE" in vendor:
            url = WARRANTY_URLS["APPLE"]
        elif "ASUS" in vendor:
            url = WARRANTY_URLS["ASUS"]
        else:
            url = WARRANTY_URLS["GENERIC"].format(vendor=vendor, model=model)

        webbrowser.open(url)
        self.lbl_status.configure(text=f"Atvērts {vendor} portāls (S/N nokopēts starpliktuvē)!")

    def send_to_offer_studio(self):
        vendor = self.entry_vendor.get().strip()
        model = self.entry_model.get().strip()
        sn = self.entry_sn.get().strip()
        pn = self.entry_pn.get().strip()
        if self.on_send_to_offer_fn:
            req_text = f"Iekārtas apkope / paplašinājums:\nRažotājs: {vendor}\nModelis: {model}\nS/N: {sn}\nP/N: {pn}\nNepieciešama specifikācija un cenas opcijas."
            self.on_send_to_offer_fn(req_text)
