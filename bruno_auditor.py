# -*- coding: utf-8 -*-
"""
BRUNO (MI Arhitekts) — Sistēmas Integritātes & Drošības Audits.
Pārbauda sintaksi, aizliegtās API atslēgas un arhitektūras tīrību.
"""
import ast
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# Drošības filtri: meklējam nejauši aizmirstas atslēgas
SUSPICIOUS_PATTERNS = [
    r"AIzaSy[A-Za-z0-9_-]{33}",           # Gemini / Google API Key
    r"sk-[A-Za-z0-9]{32,}",                # OpenAI / citi API tokeni
    r"[0-9]{8,10}:[A-Za-z0-9_-]{35}",      # Telegram Bot Token
]

def audit_file(file_path: Path):
    rel_path = file_path.relative_to(BASE_DIR)
    print(f"🧐 [BRUNO] Pārbauda: {rel_path}")

    try:
        content = file_path.read_text(encoding="utf-8")
    except Exception as e:
        print(f"⚠️ Nevarēja nolasīt {rel_path}: {e}")
        return True

    # 1. Drošības audits (Secret Leak Check)
    for pattern in SUSPICIOUS_PATTERNS:
        if re.search(pattern, content):
            print(f"❌ [BRUNO VETO] Drošības risks failā {rel_path}! Atrasta aizdomīga API atslēga vai parole.")
            return False

    # 2. Sintakses audits Python failiem
    if file_path.suffix == ".py":
        try:
            ast.parse(content)
        except SyntaxError as e:
            print(f"❌ [BRUNO VETO] Sintakses kļūda failā {rel_path}: {e}")
            return False

    return True

def run_full_audit():
    print("🛡️ [BRUNO] Sākas AuraQuadroOS pilnais arhitektūras audits...")
    passed = True
    py_and_json = list(BASE_DIR.glob("*.py")) + list(BASE_DIR.glob("*.json"))

    for item in py_and_json:
        if item.name in ["config.json", "radar_state.json"]:
            continue  # Lokālos konfigus nelamājam
        if not audit_file(item):
            passed = False

    if passed:
        print("✅ [BRUNO AUDIT PASSED] Viss ir tīrs, stabils un gatavs Kvarka un Lāča vadībai!")
        return 0
    else:
        print("⛔ [BRUNO AUDIT FAILED] Konstatēti riski! Izmaiņas noraidītas.")
        return 1

if __name__ == "__main__":
    sys.exit(run_full_audit())
