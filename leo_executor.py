# -*- coding: utf-8 -*-
"""
LEO (MI Koderis) — Autonomais Koda Izpildītājs.
Nolasa uzdevumu, izmaina mērķa failu un pārbauda sintaksi caur Bruno auditu.
"""
import os
import sys
import json
import ast
import subprocess
from pathlib import Path
import google.generativeai as genai

BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
CONTEXT_FILE = BASE_DIR / "SYSTEM_CONTEXT.md"


def load_keys():
    if not CONFIG_FILE.exists():
        raise FileNotFoundError("config.json netika atrasts!")
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    gemini_key = cfg.get("gemini_api_key") or cfg.get("aq_ai_api_key")
    if not gemini_key:
        raise ValueError("Nav norādīta Gemini API atslēga config.json failā.")
    return gemini_key


def load_system_context():
    if CONTEXT_FILE.exists():
        with open(CONTEXT_FILE, "r", encoding="utf-8") as f:
            return f.read()
    return "Tu esi Leo, pieredzējis Python koderis."


def run_leo_task(target_filename: str, task_description: str):
    target_path = BASE_DIR / target_filename
    if not target_path.exists():
        print(f"❌ Fails {target_filename} neeksistē!")
        return False

    print(f"🐾 [LEO] Uzsāk darbu pie: {target_filename}")
    print(f"📋 Uzdevums: {task_description}")

    # 1. Nolasām esošo faila kodu
    with open(target_path, "r", encoding="utf-8") as f:
        current_code = f.read()

    # 2. Sagatavojam Gemini instrukciju
    gemini_key = load_keys()
    system_context = load_system_context()
    genai.configure(api_key=gemini_key)
    model = genai.GenerativeModel("gemini-2.5-flash")

    bt = chr(96) * 3  # Trīskāršais backtick drošai formatēšanai

    prompt = (
        f"{system_context}\n\n"
        f"Tu esi LEO — elitārs Python programmētājs.\n"
        f"Tev ir precīzi jāatjaunina šis fails: `{target_filename}`.\n\n"
        f"UZDEVUMS:\n{task_description}\n\n"
        f"ESOŠAIS KODS:\n{bt}python\n{current_code}\n{bt}\n\n"
        f"Dzelžaini noteikumi:\n"
        f"1. Atgriez TIKAI un VIENĪGI pilnu atjaunināto Python kodu.\n"
        f"2. Neiekļauj nekādus paskaidrojumus, piezīmes vai ievadvārdus.\n"
        f"3. Kodu iekļauj {bt}python un {bt} blokā.\n"
        f"4. Nelauz esošo arhitektūru un loģiku."
    )

    print("🧠 [LEO] Gemini analizē un pārraksta kodu...")
    response = model.generate_content(prompt)
    raw_text = response.text.strip()

    # Iztīrām koda marķējumus
    if raw_text.startswith("```python"):
        new_code = raw_text[len("```python"):].strip()
    elif raw_text.startswith("```"):
        new_code = raw_text[len("```"):].strip()
    else:
        new_code = raw_text

    if new_code.endswith("```"):
        new_code = new_code[:-3].strip()

    # 3. Bruno drošības pārbaude (Python sintakses audits)
    try:
        ast.parse(new_code)
        print("✅ [BRUNO] Sintakses audits iziets: Kods ir sintaktiski derīgs Python!")
    except SyntaxError as e:
        print(f"❌ [BRUNO VETO] Kļūda Leo ģenerētajā kodā! Sintakse nederīga: {e}")
        return False

    # 4. Faila ierakstīšana diskā
    with open(target_path, "w", encoding="utf-8") as f:
        f.write(new_code)
    print(f"💾 [LEO] Fails {target_filename} veiksmīgi atjaunināts!")

    # 5. Git komits
    try:
        subprocess.run(["git", "add", target_filename], cwd=BASE_DIR, check=True)
        commit_msg = f"refactor(leo): {task_description[:50]}"
        subprocess.run(["git", "commit", "-m", commit_msg], cwd=BASE_DIR, check=True)
        print(f"🚀 [LEO] Git commit izpildīts: '{commit_msg}'")
    except Exception as e:
        print(f"ℹ️ Git commit solis izlaists: {e}")

    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Lietošana: python leo_executor.py <fails> <uzdevums>")
        print("Piemērs: python leo_executor.py investor_studio.py 'Pievieno jaunu pogu'")
        sys.exit(1)

    target = sys.argv[1]
    task = " ".join(sys.argv[2:])
    run_leo_task(target, task)
