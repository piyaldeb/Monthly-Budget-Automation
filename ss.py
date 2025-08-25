import os
import math
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont
import pywhatkit

# ===== CONFIG =====
SHEET_ID = "1WIJOkAHouWkXVXhsa04dCvHkG6ESrqI91If7OkQbAXw"
GID = 1102281456                      # worksheet gid (int)
SERVICE_FILE = "odoo-automation-465010-976566cf6fbb.json"
OUTPUT_IMAGE = "snapshot.png"
WHATSAPP_NUMBER = "+8801799306165"
CAPTION = "Snapshot: B3:K3 and B40:K40"
# ==================

# ---------- Google Sheets ----------
def open_worksheet(sheet_id: str, gid: int):
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_FILE, scope)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_id)
    ws = sh.get_worksheet_by_id(gid)
    if not ws:
        raise RuntimeError(f"Worksheet with gid={gid} not found")
    return ws

def get_row_values(ws, a1_range: str, expected_cols: int = 10):
    """
    Returns a list of 'expected_cols' strings. Pads with "" if sheet has fewer.
    """
    data = ws.get(a1_range)  # e.g., [['v1','v2',...]] or []
    row = (data[0] if data else [])
    row = [str(x) for x in row]
    if len(row) < expected_cols:
        row += [""] * (expected_cols - len(row))
    else:
        row = row[:expected_cols]
    return row

# ---------- Render (Pillow) ----------
def load_font(size=18):
    # Try Arial on Windows; fallback to default
    try_paths = [
        "C:/Windows/Fonts/arial.ttf",              # Windows
        "/System/Library/Fonts/Supplemental/Arial.ttf",  # macOS
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf" # Linux
    ]
    for p in try_paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size=size)
            except Exception:
                pass
    return ImageFont.load_default()

def measure_text(draw, text, font):
    # robust measurement for multi-language text
    bbox = draw.textbbox((0, 0), text, font=font)
    w = max(1, bbox[2] - bbox[0])
    h = max(1, bbox[3] - bbox[1])
    return w, h

def render_table_row(draw, top_left, cell_w, cell_h, values, font,
                     fill=(255, 255, 0),  # yellow
                     border=(0, 0, 0),    # black
                     grid=True,
                     border_thickness=3):
    """
    Draw a single row of 10 cells with text centered.
    """
    x, y = top_left
    # Outer row rectangle
    draw.rectangle([x, y, x + sum(cell_w), y + cell_h],
                   outline=border, width=border_thickness, fill=fill)

    # Vertical grid lines + cell text
    cur_x = x
    for i, val in enumerate(values):
        # Cell box
        cell_box = [cur_x, y, cur_x + cell_w[i], y + cell_h]
        if grid:
            # inner cell border
            draw.rectangle(cell_box, outline=border, width=1)
        # Center text
        tw, th = measure_text(draw, val, font)
        tx = cur_x + (cell_w[i] - tw) / 2
        ty = y + (cell_h - th) / 2
        draw.text((tx, ty), val, font=font, fill=(0, 0, 0))
        cur_x += cell_w[i]

def build_snapshot(row_top, row_bottom, out_file):
    # Visual parameters
    font = load_font(size=18)
    padding_x = 20
    padding_y = 12
    cell_h = 48
    spacing_between_rows = 18
    cols = 10

    # Decide column widths based on text width (both rows)
    # Minimum width to keep it readable
    min_w = 110
    # temp image for measuring
    tmp = Image.new("RGB", (10, 10), "white")
    d = ImageDraw.Draw(tmp)
    cell_w = []
    for i in range(cols):
        tmax = max(
            measure_text(d, row_top[i] if i < len(row_top) else "", font)[0],
            measure_text(d, row_bottom[i] if i < len(row_bottom) else "", font)[0]
        )
        cell_w.append(max(min_w, tmax + 2 * padding_x))

    total_w = sum(cell_w) + 2  # small margin
    total_h = cell_h * 2 + spacing_between_rows + 2 * padding_y

    img = Image.new("RGB", (total_w, total_h), "white")
    draw = ImageDraw.Draw(img)

    # Draw first row
    render_table_row(draw, (1, padding_y), cell_w, cell_h, row_top, font)
    # Draw second row
    render_table_row(draw, (1, padding_y + cell_h + spacing_between_rows), cell_w, cell_h, row_bottom, font)

    img.save(out_file)
    return out_file

# ---------- WhatsApp ----------
def send_whatsapp(number, image_path, caption):
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"{image_path} not found")

    print("[INFO] Opening WhatsApp Web… ensure you're logged in (web.whatsapp.com).")
    # Give WhatsApp plenty of time to load the chat and attach the image
    pywhatkit.sendwhats_image(
        receiver=number,
        img_path=image_path,
        caption=caption,
        wait_time=90,   # seconds to wait before automation types & sends
        tab_close=False,
        close_time=10
    )
    print("[OK] Sent (verify on your phone).")

def main():
    ws = open_worksheet(SHEET_ID, GID)

    # Get B3:K3 and B40:K40 (always return 10 cells, pad with "")
    row3  = get_row_values(ws,  "B3:K3", expected_cols=10)
    row40 = get_row_values(ws, "B40:K40", expected_cols=10)

    # Build compact, readable snapshot with yellow cells & black borders
    out = build_snapshot(row3, row40, OUTPUT_IMAGE)
    print(f"[OK] Snapshot created -> {out}")

    # Send on WhatsApp Web
    send_whatsapp(WHATSAPP_NUMBER, OUTPUT_IMAGE, CAPTION)

if __name__ == "__main__":
    main()
