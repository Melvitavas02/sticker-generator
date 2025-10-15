# sticker_generator_clean.py
# Cleaned Sticker Generator (keeps original behavior)
import sys
import subprocess
import importlib

# Auto install missing packages
required = ["pandas", "openpyxl", "reportlab", "qrcode[pil]", "requests", "pillow"]
for pkg in required:
    try:
        importlib.import_module(pkg.split("[")[0])
    except ImportError:
        print(f"Installing missing package: {pkg}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

# Imports
import io
import re
import requests
import pandas as pd
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth

# Detect Colab
try:
    from google.colab import files
    COLAB = True
except Exception:
    COLAB = False

def generate_stickers():
    # Font selection
    fonts = {"1": "Helvetica-Bold", "2": "Times-Bold", "3": "Courier-Bold"}
    choice = input("Select font (1=Helvetica, 2=Times, 3=Courier): ").strip()
    FONT = fonts.get(choice, "Helvetica-Bold")

    # Load Excel
    if COLAB:
        print("Upload Excel file:")
        uploaded = files.upload()
        filename = list(uploaded.keys())[0]
    else:
        filename = input("Enter Excel file path (.xlsx): ").strip()

    # Use openpyxl engine explicitly
    raw = pd.read_excel(filename, header=None, dtype=str, engine="openpyxl")

    # Detect header row automatically
    header_row = None
    for i, row in raw.iterrows():
        line = " ".join(str(x).lower() for x in row.tolist())
        if "description" in line and ("po" in line or "po number" in line):
            header_row = i
            break
    if header_row is None:
        raise SystemExit("❌ Header row not found (no 'Description' found).")

    def safe_join(row):
        vals = [str(x).strip() for x in row if str(x).strip() and str(x).lower() != "nan"]
        return " ".join(vals)

    company, po_number = "", ""

    for i in range(header_row):
        text = safe_join(raw.iloc[i].tolist())
        if not text:
            continue
        m1 = re.search(r'(client|company)\s*[:\-]\s*(.+)', text, re.I)
        m2 = re.search(r'PO\s*Number\s*[:\-]\s*([A-Za-z0-9\-_\/]+)', text, re.I)
        if m1:
            company = m1.group(2).strip()
        if m2:
            po_number = m2.group(1).strip()

    df = pd.read_excel(filename, header=header_row, dtype=str, engine="openpyxl").fillna("")
    df.columns = [c.strip().lower() for c in df.columns]

    print(f"\n✅ Header row: {header_row+1}")
    print(f"✅ Company detected: {company}")
    print(f"✅ PO Number detected: {po_number}\n")

    # Logo & QR setup
    add_logo = input("Include MSG logo? (y/n): ").strip().lower() == "y"
    logo_img = None
    if add_logo:
        try:
            r = requests.get("https://www.msgoilfield.com/logo.png", timeout=10)
            r.raise_for_status()
            logo_img = ImageReader(io.BytesIO(r.content))
            print("✅ Logo loaded.")
        except Exception as e:
            print("⚠ Logo load failed:", e)

    add_qr = input("Include QR code? (y/n): ").strip().lower() == "y"
    qr_img = None
    if add_qr:
        qr = qrcode.QRCode(box_size=2, border=1)
        qr.add_data("https://www.msgoilfield.com")
        qr.make(fit=True)
        qr_bytes = io.BytesIO()
        qr.make_image(fill_color="black", back_color="white").save(qr_bytes, format="PNG")
        qr_bytes.seek(0)
        qr_img = ImageReader(qr_bytes)
        print("✅ QR code generated.")

    # Layout constants
    BASE_LABEL_W, BASE_LABEL_H = 58 * mm, 39 * mm
    COMPACT_LABEL_H = 37 * mm
    PAD = 2 * mm
    RIGHT_MARGIN = 3 * mm
    SPACING = 1.15
    output_pdf = "stickers_msg_final.pdf"

    def extract_dpe_code(text):
        m = re.search(r'DPE\s*Item\s*Code\s*[:\-]?\s*([A-Za-z0-9\-_/]+)', str(text), re.I)
        return m.group(1).strip() if m else ""

    def clean_description(t):
        t = str(t)
        t = re.sub(r'(?i)dpe\s*item\s*code\s*[:\-]?\s*[A-Za-z0-9\-_/]+', '', t)
        t = re.sub(r'(?i)item\s*description\s*:?','', t)
        t = " ".join(t.replace("\r", "\n").splitlines())
        t = re.sub(r'\s{2,}', ' ', t)
        t = re.sub(r'\s*,\s*', ', ', t)
        return t.strip()

    def wrap_text(text, font, size, width):
        if not text:
            return []
        words = text.split()
        lines = []
        line = ""
        for w in words:
            test = (line + " " + w).strip()
            if stringWidth(test, font, size) <= width:
                line = test
            else:
                if line:
                    lines.append(line)
                line = w
        if line:
            lines.append(line)
        return lines

    # Generate Stickers
    c = canvas.Canvas(output_pdf, pagesize=(BASE_LABEL_W, BASE_LABEL_H))

    for _, row in df.iterrows():
        get = lambda x: str(row.get(x, "")).strip()

        item_no = get("sl no") or get("item no") or ""
        dpe_code = get("dpe item code") or get("dpe code") or extract_dpe_code(get("description"))
        desc = clean_description(get("description"))
        po_qty = get("po qty") or ""
        uom = get("uom") or ""
        heat = get("heat number") or ""
        cert = get("certificate number") or ""
        make = get("make") or ""
        remarks = get("remarks") or ""

        LABEL_H = COMPACT_LABEL_H if not remarks.strip() else BASE_LABEL_H
        c.setPageSize((BASE_LABEL_W, LABEL_H))

        header_lines = []
        if company:
            header_lines.append(company)
        if po_number:
            header_lines.append(f"PO Number: {po_number}")
        if item_no:
            header_lines.append(f"ITEM NO: {item_no}")
        if dpe_code:
            header_lines.append(f"DPE ITEM CODE: {dpe_code}")

        footer_fixed = [f"PO QTY: {po_qty} {uom}".strip()]
        if heat.strip():
            footer_fixed.append(f"HEAT NUMBER: {heat}".strip())
        footer_fixed.append(f"MAKE: {make}".strip() if make.strip() else "MAKE: ")
        if cert.strip():
            footer_fixed.append(f"CERTIFICATE NO: {cert}".strip())

        available_width = BASE_LABEL_W - 2 * PAD - 15
        min_size, max_size = 2.0, 5.5

        # auto-fit font size
        chosen_font_size = min_size
        for font_size in [x / 10 for x in range(int(max_size * 10), int(min_size * 10) - 1, -1)]:
            wrapped_desc = wrap_text(desc, FONT, font_size, available_width)
            wrapped_remarks = wrap_text(f"Remarks: {remarks}", FONT, font_size, available_width) if remarks.strip() else []
            total_height = ((len(header_lines) + len(wrapped_desc) + len(footer_fixed) + len(wrapped_remarks)) * font_size * SPACING)
            if total_height <= LABEL_H - 6 * mm:
                chosen_font_size = font_size
                break

        font_size = chosen_font_size

        # Draw text
        c.setFont(FONT, font_size)
        y = LABEL_H - PAD - font_size

        for line in header_lines:
            c.drawString(PAD, y, line)
            y -= font_size * SPACING
        y -= font_size * 0.4

        if logo_img:
            try:
                c.drawImage(logo_img, BASE_LABEL_W - 12 * mm - RIGHT_MARGIN,
                            LABEL_H - 7 * mm - 1.5 * mm, width=12 * mm,
                            height=7 * mm, mask="auto")
            except Exception:
                pass

        y = min(y, LABEL_H - 7 * mm - 4 * mm)
        for line in wrap_text(desc, FONT, font_size, available_width):
            c.drawString(PAD, y, line)
            y -= font_size * SPACING
        y -= font_size * 0.5

        for line in footer_fixed:
            for sub in wrap_text(line, FONT, font_size, available_width):
                c.drawString(PAD, y, sub)
                y -= font_size * SPACING

        if remarks.strip():
            for line in wrap_text(f"Remarks: {remarks}", FONT, font_size, available_width):
                c.drawString(PAD, y, line)
                y -= font_size * SPACING

        if qr_img:
            try:
                c.drawImage(qr_img, BASE_LABEL_W - 6 * mm - RIGHT_MARGIN,
                            PAD, width=6 * mm, height=6 * mm, mask="auto")
            except Exception:
                pass

        c.showPage()

    try:
        c.save()
    except RuntimeError as e:
        if "can only be saved once" in str(e):
            print("⚠ PDF already saved, skipping duplicate save.")
        else:
            raise

    print(f"\n✅ Stickers created successfully: {output_pdf}")
    if COLAB:
        try:
            files.download(output_pdf)
        except Exception:
            print("⚠ Unable to auto-download in this environment.")
    else:
        print("Saved to:", output_pdf)

if __name__ == "__main__":
    generate_stickers()
