import json
import os
import threading
import customtkinter as ctk
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item
import win32com.client

STATE_FILE = "radar_state.json"


def open_email_in_outlook(entry_id):
  """Atver konkrēto e-pastu tieši Outlook programmā."""
  try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
        "MAPI"
    )
    msg = outlook.GetItemFromID(entry_id)
    msg.Display()
  except Exception as e:
    print(f"⚠️ Kļūda atverot e-pastu: {e}")


class RadarApp(ctk.CTk):

  def __init__(self):
    super().__init__()

    self.title("AQ Termiņu Radars")
    self.geometry("420x600")
    self.resizable(False, False)
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    # Novietojam logu apakšējā labajā stūrī virs pulksteņa
    sw = self.winfo_screenwidth()
    sh = self.winfo_screenheight()
    self.geometry(f"420x600+{sw - 440}+{sh - 680}")

    # Aizverot tikai paslēpjam sistēmas joslā
    self.protocol("WM_DELETE_WINDOW", self.hide_window)

    # Galvene
    header_frame = ctk.CTkFrame(self, fg_color="transparent")
    header_frame.pack(fill="x", padx=12, pady=(12, 6))

    lbl_title = ctk.CTkLabel(
        header_frame,
        text="⚡ AQ Termiņu Radars",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    lbl_title.pack(side="left")

    btn_reload = ctk.CTkButton(
        header_frame,
        text="🔄",
        width=30,
        height=28,
        command=self.load_cards,
        fg_color="#1e293b",
    )
    btn_reload.pack(side="right")

    # Ritināmais kartīšu laukums
    self.scroll_frame = ctk.CTkScrollableFrame(self, width=390, height=520)
    self.scroll_frame.pack(padx=10, pady=5, fill="both", expand=True)

    self.load_cards()

  def hide_window(self):
    self.withdraw()

  def show_window(self):
    self.deiconify()
    self.lift()
    self.load_cards()

  def mark_as_done(self, entry_id, card_widget):
    if os.path.exists(STATE_FILE):
      try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
          data = json.load(f)
        if entry_id in data:
          data[entry_id]["status"] = "DONE"
          with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
      except Exception as e:
        print(f"Kļūda saglabājot statusu: {e}")
    card_widget.destroy()

  def load_cards(self):
    for w in self.scroll_frame.winfo_children():
      w.destroy()

    if not os.path.exists(STATE_FILE):
      ctk.CTkLabel(
          self.scroll_frame,
          text="Nav atrasts radar_state.json.\nPalaidiet radar_engine.py!",
      ).pack(pady=30)
      return

    try:
      with open(STATE_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    except Exception:
      data = {}

    active_items = {
        k: v for k, v in data.items() if v.get("status") == "ACTIVE"
    }

    if not active_items:
      ctk.CTkLabel(
          self.scroll_frame,
          text="✨ Visi uzdevumi nokārtoti!",
          font=ctk.CTkFont(size=14),
      ).pack(pady=40)
      return

    for entry_id, item_data in active_items.items():
      card = ctk.CTkFrame(self.scroll_frame, corner_radius=8, fg_color="#131b2e")
      card.pack(pady=6, padx=4, fill="x")

      cat = item_data.get("category", "Cits")
      badge_bg = (
          "#ef4444"
          if cat in ["Līgums", "Kritisks"]
          else "#f59e0b"
          if "pieprasījums" in cat.lower()
          else "#3b82f6"
      )

      top_bar = ctk.CTkFrame(card, fg_color="transparent")
      top_bar.pack(fill="x", padx=10, pady=(8, 2))

      badge = ctk.CTkLabel(
          top_bar,
          text=f" {cat} ",
          fg_color=badge_bg,
          corner_radius=4,
          text_color="white",
          font=ctk.CTkFont(size=10, weight="bold"),
      )
      badge.pack(side="left")

      if item_data.get("received"):
        rec_lbl = ctk.CTkLabel(
            top_bar,
            text=item_data.get("received"),
            text_color="#64748b",
            font=ctk.CTkFont(size=10),
        )
        rec_lbl.pack(side="right")

      lbl_subj = ctk.CTkLabel(
          card,
          text=item_data.get("subject", ""),
          font=ctk.CTkFont(size=12, weight="bold"),
          wraplength=350,
          justify="left",
      )
      lbl_subj.pack(anchor="w", padx=10, pady=(2, 0))

      lbl_sender = ctk.CTkLabel(
          card,
          text=f"No: {item_data.get('sender', '')}",
          font=ctk.CTkFont(size=11),
          text_color="#94a3b8",
          wraplength=350,
          justify="left",
      )
      lbl_sender.pack(anchor="w", padx=10)

      req = item_data.get("core_request")
      if req:
        lbl_req = ctk.CTkLabel(
            card,
            text=req,
            font=ctk.CTkFont(size=11),
            text_color="#cbd5e1",
            wraplength=350,
            justify="left",
        )
        lbl_req.pack(anchor="w", padx=10, pady=4)

      dl = item_data.get("deadline")
      if dl:
        lbl_dl = ctk.CTkLabel(
            card,
            text=f"⏰ Termiņš: {dl}",
            text_color="#f87171",
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        lbl_dl.pack(anchor="w", padx=10, pady=(0, 4))

      # Pogas
      btn_row = ctk.CTkFrame(card, fg_color="transparent")
      btn_row.pack(fill="x", padx=10, pady=(4, 8))

      btn_open = ctk.CTkButton(
          btn_row,
          text="↗ Atvērt",
          width=75,
          height=24,
          fg_color="#2563eb",
          hover_color="#1d4ed8",
          font=ctk.CTkFont(size=11),
          command=lambda eid=entry_id: open_email_in_outlook(eid),
      )
      btn_open.pack(side="left", padx=(0, 6))

      btn_done = ctk.CTkButton(
          btn_row,
          text="✓ Nokārtots",
          width=85,
          height=24,
          fg_color="#059669",
          hover_color="#047857",
          font=ctk.CTkFont(size=11),
          command=lambda eid=entry_id, c=card: self.mark_as_done(eid, c),
      )
      btn_done.pack(side="left")


def run_tray(app):
  img = Image.new("RGBA", (64, 64), color=(0, 0, 0, 0))
  d = ImageDraw.Draw(img)
  d.ellipse((6, 6, 58, 58), fill="#f59e0b")

  def on_open(icon, itm):
    app.after(0, app.show_window)

  def on_quit(icon, itm):
    icon.stop()
    app.after(0, app.destroy)

  menu = pystray.Menu(
      item("Atvērt Radaru", on_open, default=True), item("Iziet", on_quit)
  )

  icon = pystray.Icon("AQRCS", img, "AQ Termiņu Radars", menu)
  icon.run()


if __name__ == "__main__":
  app = RadarApp()
  t = threading.Thread(target=run_tray, args=(app,), daemon=True)
  t.start()
  app.mainloop()