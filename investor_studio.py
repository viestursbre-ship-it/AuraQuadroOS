# -*- coding: utf-8 -*-
import os
import sys
import json
import threading
from pathlib import Path
import customtkinter as ctk
import yfinance as yf
from datetime import datetime

try:
    import google.generativeai as genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
PORTFOLIO_FILE = BASE_DIR / "portfolio.json"

GROWTH_WATCHLIST = [
    {"ticker": "PLTR", "name": "Palantir (Enterprise AI / Defense)"},
    {"ticker": "ARM", "name": "ARM Holdings (AI & Mobile Chips)"},
    {"ticker": "CRWD", "name": "CrowdStrike (Next-Gen Cybersecurity)"},
    {"ticker": "SMCI", "name": "Super Micro Computer (AI Servers)"},
    {"ticker": "SNOW", "name": "Snowflake (Data Cloud / AI infra)"},
    {"ticker": "RKLB", "name": "Rocket Lab (Space Economy / Satellites)"},
    {"ticker": "IONQ", "name": "IonQ (Quantum Computing Pioneer)"}
]


class InvestorStudioFrame(ctk.CTkFrame):
    def __init__(self, master, get_gemini_key_fn=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.get_gemini_key = get_gemini_key_fn

        # Augšējā ciļņu josla moduļa iekšienē
        self.tabview = ctk.CTkTabview(self, fg_color="#0f172a", segmented_button_fg_color="#1e293b")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_rentgen = self.tabview.add("🔍 Viena Tikera Rentgens")
        self.tab_portfolio = self.tabview.add("💼 Mans eToro Portfelis & Audits")
        self.tab_growth = self.tabview.add("🚀 Kvarka Izaugsmes Mednieks")
        self.tab_dividends = self.tabview.add("💰 Dividendes & Naudas Plūsma")
        self.tab_market = self.tabview.add("🕵️ Tirgus Klimats & Insaideri")

        self._setup_rentgen_tab()
        self._setup_portfolio_tab()
        self._setup_growth_tab()
        self._setup_dividends_tab()
        self._setup_market_tab()

    # ==========================================
    # 1. CILNE: VIENA TIKERA RENTGENS
    # ==========================================
    def _setup_rentgen_tab(self):
        f = self.tab_rentgen
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        ctk.CTkLabel(left, text="📈 TIKERA ANALĪZE", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=15, pady=(15, 6))

        in_row = ctk.CTkFrame(left, fg_color="transparent")
        in_row.pack(fill="x", padx=15, pady=(0, 10))

        self.entry_ticker = ctk.CTkEntry(in_row, width=130, height=32, placeholder_text="piem., NVDA", font=ctk.CTkFont(size=13, weight="bold"))
        self.entry_ticker.pack(side="left", padx=(0, 6))
        self.entry_ticker.bind("<Return>", lambda e: self.start_single_analysis())

        self.btn_analyze = ctk.CTkButton(in_row, text="⚡ Rentgens", width=85, height=32, fg_color="#0284c7", hover_color="#0369a1", font=ctk.CTkFont(size=12, weight="bold"), command=self.start_single_analysis)
        self.btn_analyze.pack(side="left")

        q_frame = ctk.CTkFrame(left, fg_color="transparent")
        q_frame.pack(fill="x", padx=15, pady=(0, 8))
        for t in ["NVDA", "AAPL", "MSFT", "VOO", "PLTR"]:
            ctk.CTkButton(q_frame, text=t, width=48, height=22, fg_color="#1e293b", hover_color="#334155", font=ctk.CTkFont(size=10), command=lambda v=t: self._quick_set(v)).pack(side="left", padx=2)

        self.stats_box = ctk.CTkTextbox(left, fg_color="#0b1329", font=ctk.CTkFont(family="Consolas", size=11), corner_radius=8)
        self.stats_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.stats_box.insert("1.0", "Ievadiet tikeri un nospiediet 'Rentgens'...")
        self.stats_box.configure(state="disabled")

        ctk.CTkLabel(right, text="🐾 KOMANDANTA KVARKA REZOLŪCIJA", font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b").pack(anchor="w", padx=15, pady=(15, 6))
        self.analysis_box = ctk.CTkTextbox(right, fg_color="#0b1329", text_color="#f8fafc", font=ctk.CTkFont(size=12), wrap="word", corner_radius=8)
        self.analysis_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.analysis_box.insert("1.0", "Gaidu pieprasījumu...")
        self.analysis_box.configure(state="disabled")

    def _quick_set(self, t):
        self.entry_ticker.delete(0, "end")
        self.entry_ticker.insert(0, t)
        self.start_single_analysis()

    def start_single_analysis(self):
        t = self.entry_ticker.get().strip().upper()
        if not t:
            return
        self.btn_analyze.configure(state="disabled", text="⏳ Vāc...")
        self.stats_box.configure(state="normal")
        self.stats_box.delete("1.0", "end")
        self.stats_box.insert("1.0", f"Nolasa tirgus datus par {t}...\n")
        self.stats_box.configure(state="disabled")

        self.analysis_box.configure(state="normal")
        self.analysis_box.delete("1.0", "end")
        self.analysis_box.insert("1.0", "Kvarks ģenerē fundamentālo analīzi...\n")
        self.analysis_box.configure(state="disabled")

        threading.Thread(target=self._single_worker, args=(t,), daemon=True).start()

    def _single_worker(self, ticker):
        stats_text, payload = self._fetch_stock_data(ticker)
        ai_reply = self._ask_gemini_single(payload, stats_text)

        def _ui():
            self.btn_analyze.configure(state="normal", text="⚡ Rentgens")
            self.stats_box.configure(state="normal")
            self.stats_box.delete("1.0", "end")
            self.stats_box.insert("1.0", stats_text)
            self.stats_box.configure(state="disabled")

            self.analysis_box.configure(state="normal")
            self.analysis_box.delete("1.0", "end")
            self.analysis_box.insert("1.0", ai_reply)
            self.analysis_box.configure(state="disabled")

        self.after(0, _ui)

    # ==========================================
    # 2. CILNE: PORTFEĻA AUDITS & OPTIMIZĀCIJA
    # ==========================================
    def _setup_portfolio_tab(self):
        f = self.tab_portfolio
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        ctk.CTkLabel(left, text="💼 MANAS ETORO AKCIJAS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=15, pady=(15, 4))
        ctk.CTkLabel(left, text="Ievadiet savus tikerus (atdalītus ar komatu vai jaunā rindā):\nPiemērs: AAPL, NVDA, TSLA, INTC, PYPL", font=ctk.CTkFont(size=11), text_color="#94a3b8", justify="left").pack(anchor="w", padx=15, pady=(0, 6))

        self.txt_my_portfolio = ctk.CTkTextbox(left, fg_color="#0b1329", font=ctk.CTkFont(size=12), height=140)
        self.txt_my_portfolio.pack(fill="x", padx=15, pady=(0, 10))

        cash_row = ctk.CTkFrame(left, fg_color="transparent")
        cash_row.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkLabel(cash_row, text="💵 Brīvais atlikums (USD):", font=ctk.CTkFont(size=12)).pack(side="left")
        self.entry_cash = ctk.CTkEntry(cash_row, width=90, height=28, font=ctk.CTkFont(size=12, weight="bold"))
        self.entry_cash.pack(side="left", padx=10)
        self.entry_cash.insert(0, "373.00")

        # Riska profila slēdzis
        ctk.CTkLabel(
            left, 
            text="🎯 Riska Stratēģija:", 
            font=ctk.CTkFont(size=11, weight="bold"), 
            text_color="#94a3b8"
        ).pack(anchor="w", padx=15, pady=(4, 2))

        self.seg_risk = ctk.CTkSegmentedButton(
            left,
            values=["🛡️ Konservatīvs (5-10%)", "🚀 Venture & Growth"],
            height=28,
            selected_color="#0284c7",
            selected_hover_color="#0369a1"
        )
        self.seg_risk.set("🛡️ Konservatīvs (5-10%)")
        self.seg_risk.pack(fill="x", padx=15, pady=(0, 12))

        # Pogu rinda
        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(0, 10))

        self.btn_save_p = ctk.CTkButton(
            btn_row, 
            text="💾 Saglabāt", 
            width=80, 
            height=34, 
            fg_color="#334155", 
            hover_color="#475569", 
            command=self.save_portfolio
        )
        self.btn_save_p.pack(side="left", padx=(0, 8))

        self.btn_audit = ctk.CTkButton(
            btn_row, 
            text="🐾 Kvarka Audits", 
            height=34, 
            fg_color="#059669", 
            hover_color="#10b981", 
            font=ctk.CTkFont(size=12, weight="bold"), 
            command=self.start_portfolio_audit
        )
        self.btn_audit.pack(side="left", fill="x", expand=True)

        self.lbl_p_status = ctk.CTkLabel(left, text="", font=ctk.CTkFont(size=11), text_color="#10b981")
        self.lbl_p_status.pack(anchor="w", padx=15)

        header_r = ctk.CTkFrame(right, fg_color="transparent")
        header_r.pack(fill="x", padx=15, pady=(15, 6))

        ctk.CTkLabel(
            header_r, 
            text="⚖️ KVARKA AUDITS & PLĀNS", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            text_color="#f59e0b"
        ).pack(side="left")

        self.btn_copy_audit = ctk.CTkButton(
            header_r,
            text="📋 Kopēt kopsavilkumu",
            width=150,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.copy_audit_summary
        )
        self.btn_copy_audit.pack(side="right")

        self.audit_box = ctk.CTkTextbox(right, fg_color="#0b1329", text_color="#f8fafc", font=ctk.CTkFont(size=12), wrap="word", corner_radius=8)
        self.audit_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.audit_box.insert("1.0", "Ievadiet savas akcijas un nospiediet 'Kvarka Audits'...\nKvarks noteiks, kurus 'sīkuļus' likvidēt un kur gudri ieguldīt līdzekļus!")
        self.audit_box.configure(state="disabled")

        self.load_portfolio()

    def save_portfolio(self):
        raw = self.txt_my_portfolio.get("1.0", "end").strip()
        tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip()]
        cash = self.entry_cash.get().strip()
        data = {"tickers": tickers, "cash_usd": cash}
        try:
            with open(PORTFOLIO_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.lbl_p_status.configure(text="✅ Portfelis saglabāts!")
            self.after(3000, lambda: self.lbl_p_status.configure(text=""))
        except Exception as e:
            self.lbl_p_status.configure(text=f"⚠️ Kļūda: {e}")

    def load_portfolio(self):
        if PORTFOLIO_FILE.exists():
            try:
                with open(PORTFOLIO_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    tickers = data.get("tickers", [])
                    cash = data.get("cash_usd", "373.00")
                    self.txt_my_portfolio.delete("1.0", "end")
                    self.txt_my_portfolio.insert("1.0", ", ".join(tickers))
                    self.entry_cash.delete(0, "end")
                    self.entry_cash.insert(0, cash)
            except Exception:
                pass

    def _add_ticker_to_portfolio(self, ticker):
        current_tickers_raw = self.txt_my_portfolio.get("1.0", "end").strip()
        current_tickers = [x.strip().upper() for x in current_tickers_raw.replace("\n", ",").split(",") if x.strip()]
        
        if ticker.upper() not in current_tickers:
            current_tickers.append(ticker.upper())
            self.txt_my_portfolio.delete("1.0", "end")
            self.txt_my_portfolio.insert("1.0", ", ".join(current_tickers))
            self.save_portfolio() # Save and show status
            # Provide feedback in the growth tab as well
            if hasattr(self, 'growth_results_container'):
                feedback_label = ctk.CTkLabel(self.growth_results_container, text=f"✅ {ticker} pievienots portfelim!", text_color="#10b981")
                feedback_label.pack(pady=5)
                self.after(3000, feedback_label.destroy)
        else:
            if hasattr(self, 'growth_results_container'):
                feedback_label = ctk.CTkLabel(self.growth_results_container, text=f"ℹ️ {ticker} jau ir portfelī.", text_color="#f59e0b")
                feedback_label.pack(pady=5)
                self.after(3000, feedback_label.destroy)


    def copy_audit_summary(self):
        text_to_copy = self.audit_box.get("1.0", "end").strip()
        if not text_to_copy or "Ievadiet savas akcijas" in text_to_copy:
            return

        msg = (
            "🐾 *KOMANDANTA KVARKA INVESTĪCIJU AUDITS* 🐾\n"
            f"🎯 Stratēģija: {self.seg_risk.get() if hasattr(self, 'seg_risk') else 'Audits'}\n"
            "----------------------------------------\n\n"
            f"{text_to_copy}\n\n"
            "----------------------------------------\n"
            "⚡ Ģenerēts ar AQ-OS Investor Studio"
        )

        try:
            self.clipboard_clear()
            self.clipboard_append(msg)
            self.update()

            orig_text = self.btn_copy_audit.cget("text")
            orig_color = self.btn_copy_audit.cget("fg_color")
            self.btn_copy_audit.configure(text="✅ Nokopēts!", fg_color="#10b981")

            self.after(2000, lambda: self.btn_copy_audit.configure(text=orig_text, fg_color=orig_color))
        except Exception as e:
            print(f"⚠️ Kļūda kopējot: {e}")

    def start_portfolio_audit(self):
        raw = self.txt_my_portfolio.get("1.0", "end").strip()
        tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip()]
        cash = self.entry_cash.get().strip() or "373.00"

        if not tickers:
            return

        self.btn_audit.configure(state="disabled", text="⏳ Auditē...")
        self.audit_box.configure(state="normal")
        self.audit_box.delete("1.0", "end")
        self.audit_box.insert("1.0", f"Komandants Kvarks caurskata {len(tickers)} pozīcijas un finanšu rādītājus...\n")
        self.audit_box.configure(state="disabled")

        threading.Thread(target=self._audit_worker, args=(tickers, cash), daemon=True).start()

    def _audit_worker(self, tickers, cash):
        risk_mode = self.seg_risk.get() if hasattr(self, "seg_risk") else "🛡️ Konservatīvs (5-10%)"

        audit_summary = []
        for t in tickers:
            _, payload = self._fetch_stock_data(t)
            if payload:
                audit_summary.append({
                    "ticker": t,
                    "name": payload.get("name"),
                    "price": payload.get("price"),
                    "pe": payload.get("pe"),
                    "peg": payload.get("peg"),
                    "recommendation": payload.get("recommendation"),
                    "sector": payload.get("sector"),
                    "summary": payload.get("summary") # Add summary for richer context
                })

        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        reply = ""
        if api_key and HAS_GENAI and audit_summary:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")

                if "Venture" in risk_mode:
                    strategy_instruction = """
                    STRATĒĢIJA: VENTURE & GROWTH HUNTER (Agresīvs riska kapitāls).
                    - Vērtē aktīvus no nākotnes tirgus dominances, monopolstāvokļa un AI/tehnoloģiju lēciena prizmas.
                    - Augsts P/E vai īslaicīgs peļņas trūkums nav šķērslis, ja ieņēmumu lēciens un infrastruktūras nozīme ir milzīga (piemēram, ARM, SNOW, PLTR).
                    - Izbrāķē tikai stagnējošus un lēnus uzņēmumus bez tehnoloģiskas nākotnes.
                    """
                else:
                    strategy_instruction = """
                    STRATĒĢIJA: KONSERVATĪVS (Zems risks, 5-10% stabils pieaugums gadā).
                    - Vērtē aktīvus pēc reālas naudas plūsmas, veselīga P/E (< 30-35), dividenžu stabilitātes un zemiem riskiem.
                    - Ja uzņēmuma P/E ir pārmērīgs (> 80-100) vai nav peļņas un PEG > 2, norādi to kā bīstamu burbuli un iesaki samazināt pozīciju.
                    - Dod priekšroku tirgus indeksiem (VOO/SPY) un stabiliem industrijas milžiem.
                    """

                prompt = f"""Tu esi Komandants Kvarks un profesionāls finanšu auditors.
Lietotājs ir izvēlējies šādu riska režīmu:
{strategy_instruction}

Lietotāja eToro pozīcijas un brīvie līdzekļi {cash} USD:
{json.dumps(audit_summary, ensure_ascii=False, indent=2)}

Sagatavo latviešu valodā tiešu un strukturētu auditu atbilstoši izvēlētajai stratēģijai:
1. 🟢 TURĒT / PAPILDINĀT: kuri uzņēmumi atbilst šai stratēģijai.
2. ✂️ PĀRDOT / LIKVIDĒT: kuri aktīvi neatbilst šim riska profilam vai bremzē portfeli.
3. 🎯 BRĪVĀS NAUDAS ({cash} USD) SADALE: precīzs plāns, kur ieguldīt brīvos līdzekļus."""

                resp = model.generate_content(prompt)
                reply = resp.text.strip()
            except Exception as e:
                reply = f"⚠️ Audita kļūda: {e}"
        else:
            reply = "⚠️ Trūkst API atslēgas vai neizdevās nolasīt datus."

        def _ui():
            self.btn_audit.configure(state="normal", text="🐾 Kvarka Audits")
            self.audit_box.configure(state="normal")
            self.audit_box.delete("1.0", "end")
            self.audit_box.insert("1.0", reply)
            self.audit_box.configure(state="disabled")

        self.after(0, _ui)

    # ==========================================
    # 3. CILNE: KVARKA IZAUGSMES MEDNIEKS
    # ==========================================
    def _setup_growth_tab(self):
        f = self.tab_growth
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        ctk.CTkLabel(left, text="🚀 POTENCIĀLA RADARS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=15, pady=(15, 6))
        ctk.CTkLabel(left, text="Perspektīvi izaugsmes uzņēmumi (AI, čipi, aizsardzība):", font=ctk.CTkFont(size=11), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(0, 10))

        scroll = ctk.CTkScrollableFrame(left, fg_color="#0b1329", height=240)
        scroll.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        for item in GROWTH_WATCHLIST:
            card = ctk.CTkFrame(scroll, fg_color="#131b2e", corner_radius=6)
            card.pack(fill="x", pady=4, padx=2)
            ctk.CTkLabel(card, text=item["ticker"], font=ctk.CTkFont(size=12, weight="bold"), text_color="#38bdf8").pack(side="left", padx=8, pady=4)
            ctk.CTkLabel(card, text=item["name"], font=ctk.CTkFont(size=10), text_color="#cbd5e1").pack(side="left", padx=4)
            ctk.CTkButton(card, text="Pārbaudīt", width=65, height=22, fg_color="#0284c7", font=ctk.CTkFont(size=10), command=lambda t=item["ticker"]: self._inspect_growth(t)).pack(side="right", padx=6)

        # RENAMED BUTTON
        self.btn_scan_all_growth = ctk.CTkButton(left, text="🚀 Medīt Jaunos Vienradžus", height=34, fg_color="#d97706", hover_color="#b45309", font=ctk.CTkFont(size=12, weight="bold"), command=self.start_growth_hunter)
        self.btn_scan_all_growth.pack(fill="x", padx=15, pady=(0, 15))

        # Replaced Textbox with ScrollableFrame for dynamic cards
        ctk.CTkLabel(right, text="🎯 IZAUGSMES REZOLŪCIJA & RISKA NOVĒRTĒJUMS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b").pack(anchor="w", padx=15, pady=(15, 6))
        self.growth_results_container = ctk.CTkScrollableFrame(right, fg_color="#0b1329", corner_radius=8)
        self.growth_results_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
        initial_label = ctk.CTkLabel(self.growth_results_container, text="Izvēlieties uzņēmumu vai nospiediet '🚀 Medīt Jaunos Vienradžus', lai Kvarks atlasa nākamo lielo lēcienu...",
                                     font=ctk.CTkFont(size=12), text_color="#f8fafc", wraplength=400)
        initial_label.pack(fill="x", padx=10, pady=10)


    def _inspect_growth(self, ticker):
        self.tabview.set("🔍 Viena Tikera Rentgens")
        self._quick_set(ticker)

    def start_growth_hunter(self):
        self.btn_scan_all_growth.configure(state="disabled", text="⏳ Medī labākos...")
        
        # Clear previous results and show loading message
        for widget in self.growth_results_container.winfo_children():
            widget.destroy()
        
        loading_label = ctk.CTkLabel(self.growth_results_container, text="Kvarks apkopo datus par visiem izaugsmes kandidātiem un meklē vienradžus...\n",
                                     font=ctk.CTkFont(size=12), text_color="#f8fafc", wraplength=400)
        loading_label.pack(fill="x", padx=10, pady=10)

        threading.Thread(target=self._growth_worker, daemon=True).start()

    def _growth_worker(self):
        growth_data = []
        for g in GROWTH_WATCHLIST:
            # Fetch richer data for growth analysis, including summary for qualitative aspects
            _, payload = self._fetch_stock_data(g["ticker"])
            if payload:
                growth_data.append(payload)

        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        
        unicorn_candidates = []
        sector_concentration_warning = ""
        reply_fallback = "Kvarks neatrada piemērotus vienradžus vai radās kļūda."

        if api_key and HAS_GENAI and growth_data:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")
                
                # Kvarks's 4 golden criteria
                kvark_criteria_prompt = """
                Komandanta Kvarka 4 zelta kritēriji Jauno Vienradžu medībām:
                1.  **Revenue Growth > 30-50%**: Spēcīga ieņēmumu izaugsme (vismaz virs 30%, ideāli virs 50%). Novērtēt pēc "summary" vai vispārējās tirgus izpratnes.
                2.  **Milzīgs TAM (Total Addressable Market)**: Uzņēmumam jādarbojas tirgū ar gigantisku nākotnes potenciālu (piemēram, AI, kvantu skaitļošana, kosmosa ekonomika, nākamās paaudzes kiberdrošība). Novērtēt pēc "summary" un sektora.
                3.  **Moat (Konkurences grāvis)**: Spēcīgas konkurences priekšrocības (patentēta tehnoloģija, tīkla efekti, zīmols, augstas maiņas izmaksas klientiem). Novērtēt pēc "summary".
                4.  **Founder-led (Dibinātāja vadīts)**: Vēlams, ka uzņēmumu joprojām vada tā dibinātājs vai spēcīga, ilgtermiņā domājoša vadošā komanda ar skaidru vīziju. Novērtēt pēc "summary" (var būt netieši).
                """

                prompt = f"""Tu esi agresīvs izaugsmes investīciju eksperts un Komandants Kvarks.
Tavs uzdevums ir atrast "Jaunos Vienradžus" - uzņēmumus ar 10x-100x izaugsmes potenciālu.
Lūk, dati par strauji augošiem uzņēmumiem nākotnes tehnoloģijās, katram ir pievienots arī kopsavilkums (summary):
{json.dumps(growth_data, ensure_ascii=False, indent=2)}

{kvark_criteria_prompt}

Veic analīzi un izvadi rezultātu stingri JSON formātā.
JSON atbildē ir jābūt diviem laukiem:
1. `unicorn_candidates`: Saraksts (array) ar dict (objektiem). Katram objektam jāsatur šādi lauki:
    - `ticker`: Uzņēmuma tikers.
    - `name`: Uzņēmuma nosaukums.
    - `niche`: "Slepenais ierocis" (īss 1-2 teikumu nišas apraksts, kas dod konkurences priekšrocības).
    - `why_10x_100x`: "Kāpēc 10x-100x" (ass 2 teikumu pamatojums par milzīgo izaugsmes potenciālu, balstoties uz Kvarka kritērijiem).
    - `black_swan_risk`: "Melnais Gulbis" (objektīvs 1-2 teikumu riska faktors, kas var apturēt izaugsmi).
2. `sector_concentration_warning`: "Sarkano Karogu" brīdinājums par pārmērīgu sektora koncentrāciju, ja Kvarks to pamana starp atlasītajiem kandidātiem, vai tukša stringa, ja brīdinājuma nav.

Tikai un vienīgi JSON atbilde. Nekādus citus ievadvārdus vai paskaidrojumus."""
                
                resp = model.generate_content(prompt)
                json_response_str = resp.text.strip()
                
                # Attempt to parse JSON
                try:
                    parsed_response = json.loads(json_response_str)
                    unicorn_candidates = parsed_response.get("unicorn_candidates", [])
                    sector_concentration_warning = parsed_response.get("sector_concentration_warning", "")
                    reply_fallback = "" # If successful, clear fallback
                except json.JSONDecodeError:
                    reply_fallback = f"⚠️ Gemini atgrieza nestrukturētu atbildi, nevis JSON. Daļa: {json_response_str[:500]}..."
                    unicorn_candidates = []
                    sector_concentration_warning = "⚠️ AI atbilde nav parsējama. Mēģiniet vēlreiz vai pārbaudiet API."

            except Exception as e:
                reply_fallback = f"⚠️ Kļūda sazinoties ar Gemini: {e}"
                sector_concentration_warning = f"⚠️ Kļūda sazinoties ar Gemini: {e}"
        else:
            reply_fallback = "⚠️ Trūkst API atslēgas vai neizdevās nolasīt datus."
            sector_concentration_warning = "⚠️ Trūkst API atslēgas vai neizdevās nolasīt datus."

        def _ui():
            self.btn_scan_all_growth.configure(state="normal", text="🚀 Medīt Jaunos Vienradžus")
            
            # Clear all widgets in the container
            for widget in self.growth_results_container.winfo_children():
                widget.destroy()
            
            if sector_concentration_warning:
                warning_label = ctk.CTkLabel(self.growth_results_container, text=f"🚩 Sarkanais Karogs: {sector_concentration_warning}", 
                                             font=ctk.CTkFont(size=12, weight="bold"), text_color="#dc2626", wraplength=400)
                warning_label.pack(fill="x", padx=10, pady=5)
            
            if unicorn_candidates:
                for idx, candidate in enumerate(unicorn_candidates):
                    card = ctk.CTkFrame(self.growth_results_container, fg_color="#1e293b", corner_radius=10, border_width=2, border_color="#334155")
                    card.pack(fill="x", padx=10, pady=8)

                    ctk.CTkLabel(card, text=f"{idx+1}. {candidate.get('ticker', 'N/A')} - {candidate.get('name', 'N/A')}", 
                                 font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=15, pady=(10, 2))
                    
                    ctk.CTkLabel(card, text=f"Slepenais ierocis (Niša):", 
                                 font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(5, 0))
                    ctk.CTkLabel(card, text=candidate.get('niche', 'N/A'), 
                                 font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=450, justify="left").pack(anchor="w", padx=15)
                    
                    ctk.CTkLabel(card, text=f"Kāpēc 10x-100x:", 
                                 font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(5, 0))
                    ctk.CTkLabel(card, text=candidate.get('why_10x_100x', 'N/A'), 
                                 font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=450, justify="left").pack(anchor="w", padx=15)
                    
                    ctk.CTkLabel(card, text=f"Melnais Gulbis (risks):", 
                                 font=ctk.CTkFont(size=11, weight="bold"), text_color="#94a3b8").pack(anchor="w", padx=15, pady=(5, 0))
                    ctk.CTkLabel(card, text=candidate.get('black_swan_risk', 'N/A'), 
                                 font=ctk.CTkFont(size=11), text_color="#cbd5e1", wraplength=450, justify="left").pack(anchor="w", padx=15)
                    
                    add_button = ctk.CTkButton(card, text="➕ Pievienot manam eToro sarakstam", 
                                               width=220, height=30, fg_color="#059669", hover_color="#10b981", 
                                               font=ctk.CTkFont(size=12, weight="bold"), 
                                               command=lambda t=candidate.get('ticker'): self._add_ticker_to_portfolio(t))
                    add_button.pack(anchor="w", padx=15, pady=(10, 10))
            else:
                fallback_label = ctk.CTkLabel(self.growth_results_container, text=reply_fallback,
                                              font=ctk.CTkFont(size=12), text_color="#f8fafc", wraplength=400)
                fallback_label.pack(fill="x", padx=10, pady=10)

        self.after(0, _ui)

    # ==========================================
    # 4. CILNE: DIVIDENDES & NAUDAS PLŪSMA
    # ==========================================
    def _setup_dividends_tab(self):
        f = self.tab_dividends
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        ctk.CTkLabel(
            left, 
            text="💰 NAUDAS PLŪSMAS RADARS", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            text_color="#38bdf8"
        ).pack(anchor="w", padx=15, pady=(15, 4))

        ctk.CTkLabel(
            left, 
            text="Aprēķina pasīvo dividenžu ienesīgumu no portfeļa akcijām.", 
            font=ctk.CTkFont(size=11), 
            text_color="#94a3b8", 
            wraplength=280, 
            justify="left"
        ).pack(anchor="w", padx=15, pady=(0, 12))

        self.btn_calc_dividends = ctk.CTkButton(
            left,
            text="⚡ Aprēķināt Dividenžu Plūsmu",
            height=34,
            fg_color="#0284c7",
            hover_color="#0369a1",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.start_dividend_scan
        )
        self.btn_calc_dividends.pack(fill="x", padx=15, pady=(0, 15))

        self.div_summary_box = ctk.CTkTextbox(
            left, 
            fg_color="#0b1329", 
            font=ctk.CTkFont(family="Consolas", size=11), 
            corner_radius=8
        )
        self.div_summary_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.div_summary_box.insert("1.0", "Nospiediet pogu, lai nolasītu dividenžu datus...")
        self.div_summary_box.configure(state="disabled")

        ctk.CTkLabel(
            right, 
            text="📅 KVARKA DIVIDENŽU KALENDĀRS & REZOLŪCIJA", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            text_color="#f59e0b"
        ).pack(anchor="w", padx=15, pady=(15, 6))

        self.div_analysis_box = ctk.CTkTextbox(
            right, 
            fg_color="#0b1329", 
            text_color="#f8fafc", 
            font=ctk.CTkFont(size=12), 
            wrap="word", 
            corner_radius=8
        )
        self.div_analysis_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.div_analysis_box.insert("1.0", "Gaidu dividenžu skenēšanas palaišanu...")
        self.div_analysis_box.configure(state="disabled")

    def start_dividend_scan(self):
        raw = self.txt_my_portfolio.get("1.0", "end").strip()
        tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip()]
        if not tickers:
            return

        self.btn_calc_dividends.configure(state="disabled", text="⏳ Vāc dividenžu datus...")
        self.div_summary_box.configure(state="normal")
        self.div_summary_box.delete("1.0", "end")
        self.div_summary_box.insert("1.0", f"Pārbauda {len(tickers)} aktīvus tirgū...\n")
        self.div_summary_box.configure(state="disabled")

        self.div_analysis_box.configure(state="normal")
        self.div_analysis_box.delete("1.0", "end")
        self.div_analysis_box.insert("1.0", "Kvarks apkopo naudas plūsmas kalendāru...\n")
        self.div_analysis_box.configure(state="disabled")

        threading.Thread(target=self._dividend_worker, args=(tickers,), daemon=True).start()

    def _dividend_worker(self, tickers):
        from datetime import datetime

        div_data = []
        summary_lines = ["=== DIVIDENŽU PĀRSKATS ==="]

        for t in tickers:
            try:
                stock = yf.Ticker(t)
                info = stock.info or {}
                div_yield = info.get("dividendYield")
                div_rate = info.get("dividendRate", 0.0)
                ex_div = info.get("exDividendDate")

                # 1. Droša datuma apstrāde un pārvēršana uz DD.MM.YYYY
                ex_date_str = "Nav norādīts"
                if ex_div:
                    try:
                        # Ja atsūtīts kā skaitlis (Unix timestamp sekundēs vai milisekundēs)
                        if isinstance(ex_div, (int, float)):
                            ts = ex_div / 1000 if ex_div > 1e11 else ex_div
                            ex_date_str = datetime.fromtimestamp(ts).strftime("%d.%m.%Y")
                        else:
                            ex_date_str = str(ex_div).split()[0]
                    except Exception:
                        ex_date_str = str(ex_div)

                # 2. Dividenžu ienesīguma un korekciju apstrāde
                if div_yield and div_yield > 0:
                    # Ja nāk kā 0.015 (1.5%) vai jau kā 1.5
                    yield_val = div_yield * 100 if div_yield < 1 else div_yield
                    yield_pct = round(yield_val, 2)

                    # Kļūdu filtrs (piem., AVGO pēc split rāda 73%)
                    if yield_pct > 35 and t.upper() == "AVGO": # Specific correction for AVGO, based on domain knowledge
                        yield_pct = 1.35 # Corrected yield after split for AVGO

                    div_data.append({
                        "ticker": t,
                        "name": info.get("shortName", t),
                        "yield_pct": f"{yield_pct}%",
                        "rate": div_rate,
                        "ex_date": ex_date_str
                    })
                    summary_lines.append(f"{t:6} | {yield_pct:5.2f}% | Izmaksa: {div_rate} | Dat: {ex_date_str}")
                else:
                    summary_lines.append(f"{t:6} | 0.00%  | Dividendes nemaksā")
            except Exception as e:
                summary_lines.append(f"{t:6} | Kļūda: {e}")

        # Gemini rezolūcija
        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        reply = ""
        if api_key and HAS_GENAI and div_data:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")
                prompt = f"""Tu esi dividenžu un pasīvās naudas plūsmas eksperts Komandants Kvarks.
Šeit ir lietotāja portfeļa dividenču dati:
{json.dumps(div_data, ensure_ascii=False, indent=2)}

Sagatavo latviešu valodā strukturētu analīzi:
1. 💵 PASĪVĀS NAUDAS ČEMPIONI: Kuri no šiem uzņēmumiem nodrošina stabilāko un drošāko dividenžu atdevi.
2. ⚠️ ILGTSPĒJAS RISKI: Vai kādam uzņēmumam dividende nav mākslīgi uzpūsta vai apdraudēta (piemēram, naftas vai cikliskajos sektoros).
3. 🔄 REINVESTĒŠANAS STRATĒĢIJA: Kā visefektīvāk izmantot ienākošās dividendes, lai portfelis augtu eksponenciāli."""
                resp = model.generate_content(prompt)
                reply = resp.text.strip()
            except Exception as e:
                reply = f"⚠️ Gemini kļūda: {e}"
        elif not div_data:
            reply = "Portfelī nav atrasts neviens uzņēmums, kas maksā dividendes."
        else:
            reply = "Nav pieejama API atslēga."

        def _ui():
            self.btn_calc_dividends.configure(state="normal", text="⚡ Aprēķināt Dividenžu Plūsmu")
            self.div_summary_box.configure(state="normal")
            self.div_summary_box.delete("1.0", "end")
            self.div_summary_box.insert("1.0", "\n".join(summary_lines))
            self.div_summary_box.configure(state="disabled")

            self.div_analysis_box.configure(state="normal")
            self.div_analysis_box.delete("1.0", "end")
            self.div_analysis_box.insert("1.0", reply)
            self.div_analysis_box.configure(state="disabled")

        self.after(0, _ui)

    # ==========================================
    # 5. CILNE: TIRGUS KLIMATS & INSAIDERI
    # ==========================================
    def _setup_market_tab(self):
        f = self.tab_market
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        left.grid(row=0, column=0, padx=(5, 5), pady=5, sticky="nsew")

        right = ctk.CTkFrame(f, fg_color="#131b2e", corner_radius=10)
        right.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="nsew")

        ctk.CTkLabel(left, text="🌡️ TIRGUS NOSKAŅOJUMS", font=ctk.CTkFont(size=14, weight="bold"), text_color="#38bdf8").pack(anchor="w", padx=15, pady=(15, 6))

        self.lbl_fng = ctk.CTkLabel(left, text="Fear & Greed: Ielādē...", font=ctk.CTkFont(size=13, weight="bold"), text_color="#f59e0b")
        self.lbl_fng.pack(anchor="w", padx=15, pady=(0, 10))

        self.btn_refresh_market = ctk.CTkButton(
            left, 
            text="⚡ Atjaunot Klimatu & Insaiderus", 
            height=34, 
            fg_color="#0284c7", 
            hover_color="#0369a1", 
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.start_market_intel_scan
        )
        self.btn_refresh_market.pack(fill="x", padx=15, pady=(0, 15))

        self.insider_box = ctk.CTkTextbox(left, fg_color="#0b1329", font=ctk.CTkFont(family="Consolas", size=11), corner_radius=8)
        self.insider_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.insider_box.insert("1.0", "Nospiediet pogu, lai ielādētu insaideru datus un noskaņojumu...")
        self.insider_box.configure(state="disabled")

        ctk.CTkLabel(right, text="📰 JAUNĀKĀS ZIŅAS & KVARKA TAKTIKA", font=ctk.CTkFont(size=14, weight="bold"), text_color="#f59e0b").pack(anchor="w", padx=15, pady=(15, 6))
        
        self.market_news_box = ctk.CTkTextbox(right, fg_color="#0b1329", text_color="#f8fafc", font=ctk.CTkFont(size=12), wrap="word", corner_radius=8)
        self.market_news_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.market_news_box.insert("1.0", "Gaidu datu pieprasījumu...")
        self.market_news_box.configure(state="disabled")

    def start_market_intel_scan(self):
        import requests
        self.btn_refresh_market.configure(state="disabled", text="⏳ Vāc makro datus & Kvarks analizē...")
        
        def _worker():
            # 1. Fear & Greed indekss
            fng_val = "50"
            fng_status = "Neutral"
            fng_text = "Fear & Greed: Nav pieejams"
            try:
                res = requests.get("https://api.alternative.me/fng/?limit=1", timeout=5).json()
                item = res["data"][0]
                fng_val = item['value']
                fng_status = item['value_classification']
                fng_text = f"Fear & Greed: {fng_val}/100 ({fng_status})"
            except Exception:
                pass

            # 2. Makro & Likviditātes rādītāji (yfinance)
            macro_summary = []
            macro_lines = ["=== MAKRO & LIKVIDITĀTE ==="]
            try:
                macro_tickers = {"^TNX": "ASV 10Y Yield", "GC=F": "Zelts (Gold)", "BTC-USD": "Bitcoin (BTC)"}
                for sym, label in macro_tickers.items():
                    try:
                        tk = yf.Ticker(sym)
                        p = tk.fast_info.get("lastPrice") or tk.info.get("regularMarketPrice") or 0.0
                        if sym == "^TNX":
                            macro_lines.append(f"• {label:14}: {p:.2f}%")
                            macro_summary.append(f"{label}: {p:.2f}%")
                        else:
                            curr = "USD"
                            macro_lines.append(f"• {label:14}: {p:,.2f} {curr}")
                            macro_summary.append(f"{label}: {p:,.2f} {curr}")
                    except Exception:
                        macro_lines.append(f"• {label:14}: N/A")
            except Exception:
                pass

            # 3. Portfeļa tikeri
            raw = self.txt_my_portfolio.get("1.0", "end").strip()
            tickers = [x.strip().upper() for x in raw.replace("\n", ",").split(",") if x.strip() and "." not in x][:5]
            if not tickers:
                tickers = ["MSFT", "ARM", "VOO"]
            
            # 4. Finnhub insaideri un ziņas
            cfg_file = BASE_DIR / "config.json"
            finnhub_key = ""
            if cfg_file.exists():
                try:
                    with open(cfg_file, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                        finnhub_key = cfg.get("finnhub_api_key", "")
                except Exception:
                    pass

            insider_lines = ["\n=== INSAIDERU TRANSAKCIJAS (Finnhub) ==="]
            top_headlines = []

            if finnhub_key:
                for t in tickers:
                    try:
                        u = f"https://finnhub.io/api/v1/stock/insider-sentiment?symbol={t}&token={finnhub_key}"
                        r = requests.get(u, timeout=5).json()
                        d = r.get("data", [])
                        if d:
                            lat = d[-1]
                            m = lat.get("mspr", 0)
                            sig = "🟢 PĒRK" if m > 0 else "🔴 PĀRDOD" if m < 0 else "⚪ NEITRĀLS"
                            insider_lines.append(f"{t:5} | {sig:8} | MSPR: {m:6.2f} | Net: {lat.get('change', 0)}")
                        else:
                            insider_lines.append(f"{t:5} | ⚪ Mierīgs (nav lielu darījumu)")
                    except Exception:
                        insider_lines.append(f"{t:5} | Kļūda datos")

                from datetime import timedelta
                d_to = datetime.now().strftime("%Y-%m-%d")
                d_from = (datetime.now() - timedelta(days=14)).strftime("%Y-%m-%d")
                for t in tickers[:2]:
                    try:
                        nu = f"https://finnhub.io/api/v1/company-news?symbol={t}&from={d_from}&to={d_to}&token={finnhub_key}"
                        nlist = requests.get(nu, timeout=5).json()
                        if isinstance(nlist, list):
                            for n in nlist[:2]:
                                top_headlines.append(f"[{t}] {n.get('headline', '')}")
                    except Exception:
                        pass
            else:
                insider_lines.append("⚠️ config.json trūkst 'finnhub_api_key'!")

            # 5. Apvienojam kreisās puses datus (Makro + Insaideri)
            left_display_text = "\n".join(macro_lines) + "\n" + "\n".join(insider_lines)

            # 6. Komandanta Kvarka (Gemini) stratēģiskais slēdziens
            news_output = []
            if top_headlines:
                news_output.append("=== SVARĪGĀKĀS ZIŅAS ===")
                for h in top_headlines:
                    news_output.append(f"• {h}")
                news_output.append("\n" + "="*30 + "\n")

            api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
            if api_key and HAS_GENAI:
                try:
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel("gemini-2.5-flash")
                    prompt = f"""Tu esi militāri ass investīciju virspavēlnieks Komandants Kvarks (resnais runcis no Liepājas štāba).
Makro un likviditātes rādītāji: {', '.join(macro_summary)}
Tirgus baiļu/alkatības noskaņojums: {fng_text}
Insaideru signāli: {'; '.join(insider_lines[1:])}
Jaunākās ziņas: {'; '.join(top_headlines) if top_headlines else 'Nav agresīvu ziņu triecienu'}

Dod 3-5 teikumu zibenīgu operatīvo atskaiti un taktisko rīkojumu:
1. Makro fons: ko saka 10 gadu likmes, zelts un BTC par naudas plūsmu tirgū.
2. Insaideru un sentimenta vērtējums.
3. Galīgā kaujas pavēle šodienai: uzbrukt (pirkt), nocietināties (turēt) vai atkāpties skaidrā naudā."""
                    res = model.generate_content(prompt)
                    news_output.append("🐾 KOMANDANTA KVARKA STRATĒĢISKĀ REZOLŪCIJA:\n")
                    news_output.append(res.text.strip())
                except Exception as e:
                    news_output.append(f"⚠️ Kvarka kļūda: {e}")
            else:
                news_output.append("⚠️ Gemini API atslēga nav pieejama.")

            def _ui():
                self.btn_refresh_market.configure(state="normal", text="⚡ Atjaunot Klimatu & Insaiderus")
                self.lbl_fng.configure(text=fng_text)
                
                self.insider_box.configure(state="normal")
                self.insider_box.delete("1.0", "end")
                self.insider_box.insert("1.0", left_display_text)
                self.insider_box.configure(state="disabled")

                self.market_news_box.configure(state="normal")
                self.market_news_box.delete("1.0", "end")
                self.market_news_box.insert("1.0", "\n".join(news_output))
                self.market_news_box.configure(state="disabled")

            self.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    # ==========================================
    # DATU VĀKŠANA UN GEMINI METODES
    # ==========================================
    def _fetch_stock_data(self, ticker):
        stats_text = ""
        payload = {}
        try:
            stock = yf.Ticker(ticker)
            info = stock.info or {}

            short_name = info.get("shortName", ticker)
            currency = info.get("currency", "USD")
            curr_price = info.get("currentPrice") or info.get("regularMarketPrice", "N/A")
            pe = info.get("trailingPE", "N/A")
            fwd_pe = info.get("forwardPE", "N/A")
            peg = info.get("pegRatio", "N/A")
            high52 = info.get("fiftyTwoWeekHigh", "N/A")
            low52 = info.get("fiftyTwoWeekLow", "N/A")
            rec = info.get("recommendationKey", "N/A").upper()
            target_mean = info.get("targetMeanPrice", "N/A")
            div_yield = info.get("dividendYield")
            div_str = f"{round(div_yield * 100, 2)}%" if div_yield else "0.0%"
            
            stats_text = (
                f"=== {short_name} ({ticker}) ===\n"
                f"Cena:             {curr_price} {currency}\n"
                f"52 ned. diapazons:{low52} - {high52} {currency}\n"
                f"P/E (faktiskais): {pe}\n"
                f"P/E (prognoze):   {fwd_pe}\n"
                f"PEG koeficients:  {peg}\n"
                f"Dividendes:       {div_str}\n"
                f"Analītiķu vērt.:  {rec}\n"
                f"Mērķa vid. cena:  {target_mean} {currency}\n"
            )

            payload = {
                "ticker": ticker,
                "name": short_name,
                "currency": currency,
                "price": curr_price,
                "pe": pe,
                "forward_pe": fwd_pe,
                "peg": peg,
                "52w_high": high52,
                "52w_low": low52,
                "recommendation": rec,
                "target_price": target_mean,
                "dividend_yield": div_str,
                "sector": info.get("sector", "Nezināms"),
                "industry": info.get("industry", "Nezināms"), # Added industry for more specific analysis
                "summary": info.get("longBusinessSummary", "")
            }
        except Exception as e:
            stats_text = f"⚠️ Kļūda nolasot {ticker}: {e}"
            payload = {"ticker": ticker, "name": ticker, "error": str(e)} # Minimal payload to avoid breaking if fetch fails

        return stats_text, payload

    def _ask_gemini_single(self, summary_payload, stats_fallback):
        api_key = self.get_gemini_key() if self.get_gemini_key else os.environ.get("GEMINI_API_KEY", "")
        if api_key and HAS_GENAI and summary_payload:
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel("gemini-2.5-flash")
                prompt = f"""Tu esi stingrs, profesionāls investīciju analītiķis un Komandants Kvarks.
Sniedz kodolīgu, praktisku analīzi latviešu valodā par šo uzņēmumu/fondu, pamatojoties uz šādiem rādītājiem:
{json.dumps(summary_payload, ensure_ascii=False, indent=2)}

Strukturē atbildi precīzi šādās 3 sadaļās:
1. 🟢 Biznesa veselība un stiprās puses.
2. 🚩 Sarkanie karogi un riski (vai nav pārvērtēts burbulis).
3. ⚖️ Komandanta Kvarka rezolūcija (konkrēts kopsavilkums ilgtermiņa ieguldītājam: Turēt / Pārdot / Nogaidīt)."""
                resp = model.generate_content(prompt)
                return resp.text.strip()
            except Exception as e:
                return f"⚠️ Gemini analīzes kļūda: {e}"
        return stats_fallback