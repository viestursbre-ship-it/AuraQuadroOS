# -*- coding: utf-8 -*-
import os
import sys
import json
import urllib.request
from pathlib import Path


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def load_alert_config():
    cfg_file = get_base_dir() / "config.json"
    if not cfg_file.exists():
        return {}
    try:
        with open(cfg_file, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️ [AlertService]: Neizdevās nolasīt config.json: {e}")
        return {}


def send_telegram_alert(text: str, cfg: dict = None) -> bool:
    """Nosūta paziņojumu caur Telegram Bot API (JSON POST)."""
    if cfg is None:
        cfg = load_alert_config()

    tg = cfg.get("telegram", {})
    if not tg.get("enabled", False):
        return False

    token = tg.get("bot_token", "").strip()
    chat_id = tg.get("chat_id", "").strip()

    if not token or not chat_id:
        print("⚠️ [Telegram]: Nav norādīts bot_token vai chat_id!")
        return False

    url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = json.dumps({
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML"
    }).encode("utf-8")

    try:
        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
            }
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            if resp.status == 200:
                print("📲 [Telegram]: Trauksmes ziņa veiksmīgi nosūtīta!")
                return True
    except Exception as e:
        print(f"⚠️ [Telegram Kļūda]: {e}")
    return False


def notify_urgent_request(category: str, subject: str, sender: str, deadline: str = "Steidzams"):
    """Universāls zibens paziņojums uz Telegram."""
    cfg = load_alert_config()

    tg_text = (
        f"🚨 <b>AuraQuadro RADARS: {category}</b>\n\n"
        f"👤 <b>No:</b> {sender}\n"
        f"📋 <b>Temats:</b> {subject}\n"
        f"⏳ <b>Termiņš:</b> {deadline}\n"
        f"⚡ <i>Reaģēt nekavējoties!</i>"
    )

    if cfg.get("telegram", {}).get("enabled"):
        send_telegram_alert(tg_text, cfg)


if __name__ == "__main__":
    print("🔔 Testējam Telegram trauksmes kanālu...")
    notify_urgent_request(
        category="Cenu pieprasījums",
        subject="Testa pieprasījums caur Python skriptu",
        sender="AuraQuadro Sentinel",
        deadline="Tūlītējs"
    )
