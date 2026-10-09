import os
import json
from datetime import datetime, timedelta
import win32com.client
from google import genai
from google.genai import types

STATE_FILE = "radar_state.json"

def load_radar_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_radar_state(state):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

def analyze_email_with_gemini(subject, body, sender):
    api_key = os.getenv("AQ_AI_API_KEY")
    if not api_key:
        print("⚠️ Trūkst AQ_AI_API_KEY mainīgā!")
        return None

    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    Analizē šo ienākošo e-pastu un nosaki, vai tajā ir konkrēts pieprasījums, uzdevums, rēķins, apdrošināšana vai termiņš, kas prasa rīcību.
    
    No: {sender}
    Temats: {subject}
    Saturs:
    {body[:2500]}
    
    Atgriez TIKAI tīru JSON ar šādu struktūru:
    {{
      "is_actionable": true/false,
      "category": "Cenu pieprasījums" / "Apdrošināšana / Termiņš" / "Līgums" / "Rēķins" / "Cits",
      "deadline": "YYYY-MM-DD" vai null,
      "core_request": "1-2 teikumi par galveno uzdevumu vai pieprasījumu",
      "urgency": "CRITICAL" / "HIGH" / "NORMAL"
    }}
    """
    
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                response_mime_type="application/json"
            )
        )
        return json.loads(response.text.strip())
    except Exception as e:
        print(f"⚠️ Gemini kļūda: {e}")
        return None

def run_radar_scan(hours_back=48):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    since_time = datetime.now() - timedelta(hours=hours_back)
    state = load_radar_state()

    print(f"📡 Skenējam pastkasti kopš {since_time.strftime('%Y-%m-%d %H:%M')}...")

    for msg in messages:
        try:
            received_time = msg.ReceivedTime.replace(tzinfo=None)
            if received_time < since_time:
                break

            entry_id = msg.EntryID
            # Ja šis e-pasts jau ir apstrādāts vai noņemts no radara, ejam tālāk
            if entry_id in state:
                continue

            subject = msg.Subject or "(Bez temata)"
            sender = msg.SenderName or "Nezināms"
            body = msg.Body or ""

            # Ātrā analīze ar Gemini
            analysis = analyze_email_with_gemini(subject, body, sender)
            if not analysis:
                continue

            if analysis.get("is_actionable"):
                state[entry_id] = {
                    "subject": subject,
                    "sender": sender,
                    "received": received_time.strftime("%Y-%m-%d %H:%M"),
                    "category": analysis.get("category"),
                    "deadline": analysis.get("deadline"),
                    "core_request": analysis.get("core_request"),
                    "urgency": analysis.get("urgency"),
                    "status": "ACTIVE"  # Statuss: ACTIVE, DISMISSED, DONE
                }
                print(f"🎯 [RADARS] Atrasts uzdevums: {subject} ({analysis.get('category')})")
            else:
                # Piefiksējam kā fona troksni, lai vairs nepārlasītu
                state[entry_id] = {"status": "IGNORED", "subject": subject}

        except Exception:
            continue

    save_radar_state(state)
    print("✅ Skenēšana pabeigta un Radara stāvoklis saglabāts!")

if __name__ == "__main__":
    run_radar_scan(hours_back=48)