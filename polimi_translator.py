import os
import json
import re
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

# ==========================================
# 0. KONFIGURĀCIJA (VALODAS IZVĒLE)
# ==========================================
TARGET_LANG = "en"  # "en" priekš Valtera, "lv" priekš latviešu valodas

LANG_NAME = "English" if TARGET_LANG == "en" else "Latvian"
LANG_COL_LABEL = "ENGLISH TRANSLATION" if TARGET_LANG == "en" else "LATVIEŠU TULKOJUMS"

# ==========================================
# 1. SHĒMA
# ==========================================
class OptionPair(BaseModel):
    opt_key: str = Field(description="Option letter: A, B, C, D, or E")
    text_it: str = Field(description="Option text in Italian")
    text_trans: str = Field(description="Option text translated to target language")

class SingleQuestion(BaseModel):
    id: str = Field(description="Question ID, e.g. I, II, III, IV, V")
    topic_it: str = Field(description="Topic in Italian")
    topic_trans: str = Field(description="Topic translated to target language")
    question_it: str = Field(description="Question text in Italian")
    question_trans: str = Field(description="Question text translated to target language")
    options: list[OptionPair] = Field(description="List of 5 options A through E")
    solution_it: str = Field(description="Detailed solution in Italian")
    solution_trans: str = Field(description="Detailed solution translated with LaTeX math")
    correct_option: str = Field(description="Correct option letter (A, B, C, D, or E)")
    keywords: list[str] = Field(description="Key concepts in target language")

# ==========================================
# 2. CEĻI UN INSTRUKCIJA
# ==========================================
STORAGE_DIR = Path("storage")
INBOX_DIR = STORAGE_DIR / "polimi_inbox"
OUTBOX_DIR = STORAGE_DIR / "outbox"
TEMP_DIR = STORAGE_DIR / "temp_slices"

for d in [INBOX_DIR, OUTBOX_DIR, TEMP_DIR]:
    d.mkdir(parents=True, exist_ok=True)

PROMPT_EN = """
You are a professional Politecnico di Milano (PoliMi) engineering physics professor and translator.
Translate the physics entrance exam questions and their step-by-step solutions from Italian into precise academic English.

STRICT MATHEMATICS & LATEX RULES:
- DO NOT use plain unicode symbols like '√', '≈', '×', or '^' in text.
- ALWAYS wrap all mathematical, numerical, unit, and physical expressions inside LaTeX dollar signs: $...$.
- Example: write $\\sqrt{8}$ instead of √8, write $\\approx$ instead of ≈, write $\\frac{1}{2}mv^2$ instead of 1/2mv^2.
- Preserve all physics variables: $m$, $v$, $g$, $h$, $R$, $I$, $V$, $c$, $\\Delta U$.

STANDARD TERMINOLOGY (IT -> EN):
- 'forza peso' -> 'gravitational force' (or 'weight')
- 'caduta libera' -> 'free fall'
- 'quantità di moto' -> 'linear momentum'
- 'energia cinetica' -> 'kinetic energy'
- 'energia potenziale' -> 'potential energy'
- 'rendimento' -> 'efficiency'
- 'piano inclinato' -> 'inclined plane'
- 'spinta di Archimede' -> 'buoyant force (Archimedes principle)'
- 'calore latente' -> 'latent heat'
- 'tempo di dimezzamento' -> 'half-life'
"""

PROMPT_LV = """
Tu esi profesionāls Politecnico di Milano (PoliMi) inženierfizikas tulkotājs.
Iztulko fizikas iestājeksāmenu uzdevumus un atrisinājumus no itāļu valodas akadēmiskā latviešu valodā.

STINGRIE LATEX NOTEIKUMI:
- NEKAD nelieto parastos unicode simbolus '√', '≈', '×'.
- VISAS formulas un skaitliskās izteiksmes obligāti iekļauj LaTeX dolārzīmēs: $...$.
- Piemēram: raksti $\\sqrt{8}$ nevis √8, raksti $\\approx$ nevis ≈, raksti $\\frac{1}{2}mv^2$ nevis 1/2mv^2.
"""

ACTIVE_PROMPT = PROMPT_EN if TARGET_LANG == "en" else PROMPT_LV

# ==========================================
# 3. PDF UN MI APSTRĀDE
# ==========================================
def extract_pdf_page_range(src_pdf: Path, start_page: int, end_page: int, out_pdf: Path):
    reader = PdfReader(src_pdf)
    writer = PdfWriter()
    for i in range(start_page - 1, min(end_page, len(reader.pages))):
        writer.add_page(reader.pages[i])
    with open(out_pdf, "wb") as f:
        writer.write(f)

def clean_latex_escapes(raw_text: str) -> str:
    """Salabo vadības rakstzīmju kropļojumus (\\f -> \\frac)"""
    # Aizstāj vadības rakstzīmi form-feed (\x0c) ar \f
    raw_text = raw_text.replace('\x0c', '\\f')
    return raw_text

def translate_single_question(client: genai.Client, pdf_bytes: bytes, q_num: str) -> dict:
    pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
    prompt = f"""{ACTIVE_PROMPT}

TASK:
Extract and translate ONLY Question {q_num} and its corresponding solution from the provided text.
Fill the JSON schema accurately for Question {q_num}.
"""
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[prompt, pdf_part],
        config=types.GenerateContentConfig(
            temperature=0.1,
            response_mime_type="application/json",
            response_schema=SingleQuestion
        )
    )
    cleaned = clean_latex_escapes(response.text.strip())
    return json.loads(cleaned)

# ==========================================
# 4. HTML ĢENERATORS
# ==========================================
def build_interactive_html(test_title: str, questions: list, output_html_path: Path):
    cards_html = ""
    for q in questions:
        opt_rows = ""
        for opt in q.get("options", []):
            key = opt.get("opt_key", "")
            is_correct = (key == q.get("correct_option"))
            cls = "opt-correct" if is_correct else ""
            opt_rows += f"""
            <tr class="{cls}">
                <td class="opt-key"><b>{key})</b></td>
                <td class="col-it">{opt.get('text_it', '')}</td>
                <td class="col-trans">{opt.get('text_trans', '')}</td>
            </tr>
            """

        keywords_str = ", ".join(q.get("keywords", []))
        
        cards_html += f"""
        <div class="q-card">
            <div class="q-header">
                <span class="q-num">Question {q.get('id')}</span>
                <span class="q-topic">{q.get('topic_trans')} <i>({q.get('topic_it')})</i></span>
            </div>
            
            <div class="q-body-grid">
                <div class="col-it-box">
                    <div class="lang-tag">ITALIAN ORIGINAL</div>
                    <p>{q.get('question_it')}</p>
                </div>
                <div class="col-trans-box">
                    <div class="lang-tag">{LANG_COL_LABEL}</div>
                    <p>{q.get('question_trans')}</p>
                </div>
            </div>

            <table class="options-table">
                {opt_rows}
            </table>

            <div class="solution-container">
                <button class="toggle-btn" onclick="toggleSolution('{q.get('id')}')">
                    💡 Show / Hide Solution (Correct: {q.get('correct_option')})
                </button>
                <div id="sol-{q.get('id')}" class="solution-content" style="display:none;">
                    <div class="sol-grid">
                        <div class="sol-it"><b>IT:</b> {q.get('solution_it')}</div>
                        <div class="sol-trans"><b>{TARGET_LANG.upper()}:</b> {q.get('solution_trans')}</div>
                    </div>
                    <div class="sol-keys"><b>Keywords:</b> {keywords_str}</div>
                </div>
            </div>
        </div>
        """

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>{test_title} - PoliMi Prep</title>
    <script src="https://polyfill.io/v3/polyfill.min.js?features=es6"></script>
    <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
    <style>
        body {{ font-family: 'Segoe UI', -apple-system, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 24px; }}
        .container {{ max-width: 1100px; margin: 0 auto; }}
        h1 {{ color: #38bdf8; border-bottom: 2px solid #334155; padding-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }}
        .print-btn {{ background: #0284c7; color: white; border: none; padding: 8px 16px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: bold; }}
        .q-card {{ background: #1e293b; border-radius: 12px; padding: 20px; margin-bottom: 24px; border: 1px solid #334155; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.3); }}
        .q-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; border-bottom: 1px solid #334155; padding-bottom: 8px; }}
        .q-num {{ font-size: 18px; font-weight: bold; color: #38bdf8; }}
        .q-topic {{ font-size: 13px; color: #94a3b8; }}
        .q-body-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }}
        .col-it-box, .col-trans-box {{ background: #0f172a; padding: 14px; border-radius: 8px; font-size: 15px; line-height: 1.6; border: 1px solid #1e293b; }}
        .lang-tag {{ font-size: 10px; font-weight: bold; color: #64748b; letter-spacing: 1px; margin-bottom: 6px; }}
        .options-table {{ width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 14px; }}
        .options-table td {{ padding: 8px 10px; border-top: 1px solid #334155; }}
        .opt-key {{ width: 30px; color: #94a3b8; }}
        .col-it {{ width: 45%; color: #cbd5e1; }}
        .col-trans {{ width: 50%; color: #f1f5f9; }}
        .opt-correct {{ background: rgba(16, 185, 129, 0.1); border-left: 3px solid #10b981; }}
        .toggle-btn {{ background: #334155; color: #e2e8f0; border: none; padding: 10px 16px; border-radius: 6px; cursor: pointer; font-size: 13px; width: 100%; text-align: left; transition: 0.2s; }}
        .toggle-btn:hover {{ background: #475569; }}
        .solution-content {{ background: #0b1329; border: 1px solid #1e293b; border-radius: 6px; padding: 16px; margin-top: 10px; }}
        .sol-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; font-size: 14px; line-height: 1.6; margin-bottom: 10px; }}
        .sol-it {{ color: #94a3b8; border-right: 1px solid #1e293b; padding-right: 12px; }}
        .sol-trans {{ color: #e2e8f0; }}
        .sol-keys {{ font-size: 12px; color: #38bdf8; border-top: 1px solid #1e293b; padding-top: 8px; }}
        @media print {{
            body {{ background: white; color: black; }}
            .q-card {{ background: white; color: black; border: 1px solid #ccc; page-break-inside: avoid; }}
            .col-it-box, .col-trans-box {{ background: #f8fafc; color: black; border: 1px solid #eee; }}
            .solution-content {{ display: block !important; background: #fafafa; color: black; }}
            .print-btn, .toggle-btn {{ display: none; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>
            <span>🎓 {test_title} (PoliMi Ingegneria)</span>
            <button class="print-btn" onclick="window.print()">🖨 Print / Save as PDF</button>
        </h1>
        {cards_html}
    </div>
    <script>
        function toggleSolution(id) {{
            var el = document.getElementById('sol-' + id);
            el.style.display = (el.style.display === 'none') ? 'block' : 'none';
        }}
    </script>
</body>
</html>
"""
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"✨ Gatavs: {output_html_path}")

# ==========================================
# 5. TESTU IZPILDE AR PRECIZU LAPU KARTI
# ==========================================
def process_single_test(src_pdf: Path, test_num: int, start_pg: int, end_pg: int, client: genai.Client):
    suffix = "EN" if TARGET_LANG == "en" else "LV"
    out_html = OUTBOX_DIR / f"POLIMI_Fizika_Test_{test_num}_{suffix}.html"
    
    if out_html.exists():
        print(f"⏩ Test n. {test_num} ({suffix}) jau eksistē, izlaižam...")
        return

    slice_pdf = TEMP_DIR / f"Test_{test_num}_slice.pdf"
    print(f"\n==========================================")
    print(f"📄 Griežam Test n. {test_num} (lpp. {start_pg}-{end_pg})...")
    extract_pdf_page_range(src_pdf, start_page=start_pg, end_page=end_pg, out_pdf=slice_pdf)

    with open(slice_pdf, "rb") as f:
        pdf_bytes = f.read()

    parsed_data = []
    for q_id in ["I", "II", "III", "IV", "V"]:
        print(f"🧠 [Test {test_num}] Apstrādājam {q_id} ({LANG_NAME})...")
        try:
            q_res = translate_single_question(client, pdf_bytes, q_id)
            parsed_data.append(q_res)
        except Exception as e:
            print(f"⚠️ Kļūda pie uzdevuma {q_id}: {e}")

    if parsed_data:
        build_interactive_html(f"Physics - Test n. {test_num}", parsed_data, out_html)

if __name__ == "__main__":
    src_pdf = Path("Politest_FISICA.pdf")
    if not src_pdf.exists():
        src_pdf = INBOX_DIR / "Politest_FISICA.pdf"

    if not src_pdf.exists():
        print("⚠️ Politest_FISICA.pdf nav atrasts!")
        exit(1)

    api_key = os.getenv("AQ_AI_API_KEY")
    if not api_key:
        raise ValueError("AQ_AI_API_KEY nav definēts sistēmas mainīgajos!")
    client = genai.Client(api_key=api_key)

    # Grāmatas reālais indeksējums: Test 1 jautājumi sākas 18. lpp. failā (grāmatas 1. lpp.)
    # Katrs tests aizņem 4 lappuses
    START_OFFSET = 18

    print(f"🚀 Generējam PoliMi testus valodā: {LANG_NAME.upper()}")
    for t_idx in range(1, 21):
        p_start = START_OFFSET + (t_idx - 1) * 4
        p_end = p_start + 3
        process_single_test(src_pdf, test_num=t_idx, start_pg=p_start, end_pg=p_end, client=client)

    print("\n🎉 Apstrāde pabeigta! Pārbaudi storage/outbox.")