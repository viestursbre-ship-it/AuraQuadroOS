# -*- coding: utf-8 -*-
import os
import json
import random
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
STATE_FILE = BASE_DIR / "radar_state.json"
BRIEFINGS_DIR = BASE_DIR / "storage" / "briefings"
BRIEFINGS_DIR.mkdir(parents=True, exist_ok=True)

KVARK_VERDICTS = [
    "🐾 Kvarka verdikts: Trauciņš bija tukšs 17:42, enerģētiskā stabilitāte 98%. Visi rītdienas termiņi tiek stingri novēroti ar vienu pievērtu aci.",
    "🐾 Kvarka verdikts: Visuma Padome atzinīgi novērtē šodienas progresu. Pieprasu papildu zivju konservu devu par nakts maiņas uzraudzību.",
    "🐾 Kvarka verdikts: Gravitācijas lauks uz dīvāna šodien bija spēcīgs, bet procesi rit pēc plāna. Steidzamie jautājumi nodoti rītdienas rītam.",
    "🐾 Kvarka verdikts: 100% gatavība nakts cēlienam. Rītdienas termiņi izskatās draudīgi, bet es esmu gatavs aizmigt tiem virsū."
]

def generate_daily_report():
    if not os.path.exists(STATE_FILE):
        return None

    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return None

    today_str = datetime.now().strftime("%d.%m")
    
    total_active = 0
    total_done = 0
    urgent_tomorrow = []
    
    for eid, item in data.items():
        status = item.get("status", "ACTIVE")
        deadline = item.get("deadline", "")
        subj = item.get("subject", "(Bez temata)")
        
        if status == "DONE":
            total_done += 1
        else:
            total_active += 1
            if "rīt" in deadline.lower() or "rītdien" in deadline.lower() or "steidzam" in deadline.lower():
                urgent_tomorrow.append(f"  • [🔥 {deadline}] {subj[:55]}")

    verdict = random.choice(KVARK_VERDICTS)

    report_lines = [
        "=" * 55,
        f"📋 VAKARA KOMANDANTA ATSKAITE — {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "=" * 55,
        f"✅ Pabeigti / nokārtoti uzdevumi: {total_done}",
        f"⏳ Aktīvi uzdevumi rindā:        {total_active}",
    ]

    if urgent_tomorrow:
        report_lines.append("\n⚠️ RĪT VAI TŪLĪT DEG:")
        report_lines.extend(urgent_tomorrow)
    else:
        report_lines.append("\n✨ Rītdienai kritisku degpunktu nav!")

    report_lines.append("\n" + "-" * 55)
    report_lines.append(verdict)
    report_lines.append("=" * 55)

    report_text = "\n".join(report_lines)

    # Saglabājam arhīvā
    file_path = BRIEFINGS_DIR / f"briefing_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(report_text)
    except Exception:
        pass

    return report_text
