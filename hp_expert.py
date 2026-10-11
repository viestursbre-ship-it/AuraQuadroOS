# -*- coding: utf-8 -*-
import os
import threading
import customtkinter as ctk
import google.generativeai as genai

# ==========================================
# HP EKSPERTA SISTĒMAS UZVEDNE (PROMPT)
# ==========================================
HP_EXPERT_SYSTEM_PROMPT = """Tu esi sertificēts HP Inc. vecākais risinājumu arhitekts un ekosistēmas eksperts (HP Master Certified System Engineer).
Tava pamatkompetence aptver visu HP komerciālo iekārtu klāstu, sākot no G6 paaudzes līdz jaunākajiem modeļiem:

1. PĀRNĒSĀJAMIE UN GALDA DATORI:
   - EliteBook (sērijas 600, 800, 1000 no G6 līdz G10 un jaunā Next-Gen AI PC nomenklatūra)
   - ProBook (sērijas 400, 440, 450 no G6 līdz G10)
   - EliteDesk / ProDesk (Mini, SFF, Tower no G6 uz augšu)

2. STACIONĀRĀS UN MOBILĀS DARBSTACIJAS (HP Z WORKSTATIONS / WKS):
   - Mobilās darbstacijas: ZBook (Firefly, Power, Studio, Fury paaudzēs G6–G11)
   - Galda darbstacijas (WKS): Z2 (Mini, SFF, Tower), Z4, Z6, Z8 (ar Intel Xeon W un AMD Threadripper PRO)
   - WKS specifika: ECC / RDIMM atmiņa, barošanas bloki (PSU), NVIDIA RTX Ada / A-sērijas profesionālā grafika.

3. HP DISPLEJI UN MONITORI:
   - E-Series (ergonomiskie biznesa), P-Series (standarta), Z-Series (DreamColor krāsu precīzie studiju monitori).
   - Konferences un USB-C/Thunderbolt Hub monitori (E24m, Z32k, Z40c) ar USB-C Power Delivery (PD 65W–140W), RJ-45 un Daisy Chain (MST).

4. HP PRINTERI UN MULTIFUNKCIONĀLĀS IEKĀRTAS (PRINTING & MFP):
   - HP LaserJet Pro un LaserJet Enterprise (mono un Color MFP).
   - Izejmateriāli: oriģinālās kasetnes (standarta 'A' vs palielinātā 'X' resursa, piem., 59A/59X, 139A/139X), Maintenance Kits, papildu papīra kasetes.

5. AKSESUĀRI UN DOKI:
   - HP Thunderbolt Dock G2/G4, USB-C Dock G5, barošanas bloki (65W, 100W, 120W, 280W Smart AC / USB-C PD).

MEKLĒŠANAS UN IDENTIFIKĀCIJAS STRATĒĢIJA:
- Ja lietotājs ievada produkta kodu (P/N, ProdID, piem., 401N4EA vai 401N4EA#B1R):
  * Izmanto meklēšanas rīku, lai atrastu precīzo HP rūpnīcas modeli un procesoru no oficiālajiem HP datiem (QuickSpecs, support.hp.com).
  * Paskaidro koda reģionu (piem., #B1R = Baltija, #ABB = Eiropa).
  * Kad modelis un procesors atrasts (piem., 401N4EA ir HP EliteBook 845 G8 ar AMD Ryzen 5 PRO 5650U), sniedz detalizētu specifikāciju:
    1. Procesors un arhitektūra (Zen 3 / Tiger Lake / Alder Lake utt.).
    2. RAM arhitektūra (SODIMM slotu skaits, DDR4 vai DDR5, vai nav ielodēts).
    3. Glabātuve (M.2 NVMe PCIe atbalsts).
    4. Ekrāns un grafika.
    5. Ieteicamās dokstacijas un aksesuāri (piem., HP USB-C Dock G5).

ATBILDES STILS:
- Profesionāls, inženiera cienīgs tonis latviešu valodā.
- Precīzos tehniskos terminus saglabā angliski (SODIMM, RDIMM, NVMe PCIe Gen4, Thunderbolt 4, Power Delivery, QuickSpecs).
"""

# ==========================================
# HP EKSPERTA UI CILNE (CTkFrame)
# ==========================================
class HPExpertFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None):
        super().__init__(master, fg_color="transparent")
        self.get_gemini_key_fn = get_gemini_key_fn

        self.chat_box = ctk.CTkTextbox(
            self, 
            font=ctk.CTkFont(size=13), 
            text_color="#f1f5f9", 
            fg_color="#0f172a",
            wrap="word"
        )
        self.chat_box.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        self.chat_box.insert(
            "end", 
            "HP Solution Architect gatavs darbam.\n"
            "Uzdodiet jautājumu par EliteBook, ZBook, Z darbstacijām, dokiem, displejiem vai printeriem...\n\n"
        )
        self.chat_box.configure(state="disabled")

        input_frame = ctk.CTkFrame(self, fg_color="transparent")
        input_frame.pack(fill="x", padx=10, pady=10)

        self.hp_input = ctk.CTkEntry(
            input_frame, 
            placeholder_text="Piemēram: Kāda ir 401N4EA#B1R specifikācija? Kāds doks der ZBook Fury?",
            font=ctk.CTkFont(size=13),
            height=40
        )
        self.hp_input.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.hp_input.bind("<Return>", lambda e: self.ask_hp_expert())

        self.btn_ask = ctk.CTkButton(
            input_frame, 
            text="Jautāt ➔", 
            width=100, 
            height=40,
            command=self.ask_hp_expert
        )
        self.btn_ask.pack(side="right")

    def _get_api_key(self):
        if self.get_gemini_key_fn:
            key = self.get_gemini_key_fn()
            if key:
                return key
        return os.environ.get("GEMINI_API_KEY", "")

    def ask_hp_expert(self):
        query = self.hp_input.get().strip()
        if not query:
            return

        self.hp_input.delete(0, "end")
        self.chat_box.configure(state="normal")
        self.chat_box.insert("end", f"\n👤 Jūs: {query}\n")
        self.chat_box.insert("end", "⏳ HP Eksperts meklē oficiālajos avotos un analizē specifikāciju...\n")
        self.chat_box.see("end")
        self.chat_box.configure(state="disabled")

        threading.Thread(target=self._query_gemini_hp, args=(query,), daemon=True).start()

    def _query_gemini_hp(self, query):
        response_text = ""
        api_key = self._get_api_key()

        try:
            if not api_key:
                response_text = "⚠️ Gemini API atslēga nav atrasta! Pārbaudiet iestatījumus."
            else:
                genai.configure(api_key=api_key)

                # Izmantojam modeli ar iebūvēto Google Search rīku tiešsaistes meklēšanai
                model = genai.GenerativeModel(
                    model_name="gemini-2.5-flash",
                    system_instruction=HP_EXPERT_SYSTEM_PROMPT,
                    tools="google_search_retrieval"
                )

                res = model.generate_content(query)
                response_text = res.text

        except Exception as e:
            # Ja gadījumā konkrētajam modelim rīks ir ar citu nosaukumu, rezerves izsaukums:
            try:
                model_fallback = genai.GenerativeModel(
                    model_name="gemini-2.5-flash",
                    system_instruction=HP_EXPERT_SYSTEM_PROMPT
                )
                res = model_fallback.generate_content(query)
                response_text = res.text
            except Exception as e2:
                response_text = f"⚠️ Kļūda saziņā ar ekspertu: {e2}"

        def update_ui():
            self.chat_box.configure(state="normal")
            content = self.chat_box.get("1.0", "end")
            marker = "⏳ HP Eksperts meklē oficiālajos avotos un analizē specifikāciju...\n"
            idx = content.rfind(marker)
            if idx != -1:
                self.chat_box.delete(f"1.0 + {idx} chars", "end")
            self.chat_box.insert("end", f"🧠 HP Eksperts:\n{response_text.strip()}\n\n")
            self.chat_box.see("end")
            self.chat_box.configure(state="disabled")

        self.after(0, update_ui)
