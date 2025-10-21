# sticker_generator_clean.py
# Cleaned Sticker Generator (Final Updated Version)

import sys
import subprocess
import importlib

# -----------------------------
# Auto install missing packages
# -----------------------------
required = ["pandas", "openpyxl", "reportlab", "qrcode[pil]", "requests", "pillow"]
for pkg in required:
    try:
        importlib.import_module(pkg.split("[")[0])
    except ImportError:
        print(f"Installing missing package: {pkg}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

# -----------------------------
# Imports
# -----------------------------
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
    # -----------------------------
    # Font selection
    # -----------------------------
    fonts = {"1": "Helvetica-Bold", "2": "Times-Bold", "3": "Courier-Bold"}
    choice = input("Select font (1=Helvetica, 2=Times, 3=Courier): ").strip()
    FONT = fonts.get(choice, "Helvetica-Bold")

    # -----------------------------
    # Load Excel
    # -----------------------------
    if COLAB:
        print("Upload Excel file:")
        uploaded = files.upload()
        filename = list(uploaded.keys())[0]
    else:
        # allow stdin piping from Flask, else fall back to prompt
        filename = sys.stdin.readline().strip() or input("Enter Excel file path (.xlsx): ").strip()

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

    def safe_join(row_vals):
        vals = [str(x).strip() for x in row_vals if str(x).strip() and str(x).lower() != "nan"]
        return " ".join(vals)

    # Read company & PO number from the pre-header region
    company, po_number = "", ""
    for i in range(header_row):
        text = safe_join(raw.iloc[i].tolist())
        if not text:
            continue

        m1 = re.search(r'(client|company)\s*[:\-]\s*(.+)', text, re.I)
        if not m1 and len(text.split()) > 2 and "po" not in text.lower():
            # sometimes the client line is just a long name line
            company = text.strip()

        m2 = re.search(r'PO\s*Number\s*[:\-]\s*([A-Za-z0-9\-_\/]+)', text, re.I)
        if m1:
            company = m1.group(2).strip()
        if m2:
            po_number = m2.group(1).strip()

    # --- Smart Guard: detect proper client/company name or fallback ---
    company_original = company  # keep what was first detected
    low = company.lower() if company else ""

    # Case 1: If it looks like an MSG reference, try to find real client name
    if company and (low.startswith("msg") or "msg" in low):
        company = ""
        for i in range(header_row):
            text = safe_join(raw.iloc[i].tolist())
            if not text:
                continue
            mclient = re.search(r'(client|customer|company|customer name)\s*[:\-]\s*(.+)', text, re.I)
            if mclient:
                company = mclient.group(2).strip()
                break

        # Fallback – first long line that’s not MSG/PO
        if not company:
            for i in range(header_row):
                text = safe_join(raw.iloc[i].tolist())
                if text and "msg" not in text.lower() and "po" not in text.lower() and len(text.split()) > 2:
                    company = text.strip()
                    break

    # Final fallback
    if not company:
        if company_original and ("msg" in company_original.lower() or "reference" in company_original.lower()):
            company = company_original  # use MSG reference as last resort
        else:
            company = "Client: N/A"

    # Build dataframe from header
    df = pd.read_excel(filename, header=header_row, dtype=str, engine="openpyxl").fillna("")
    df.columns = [c.strip().lower() for c in df.columns]

    # Remove empty rows
    def _all_blank(row):
        return all(str(v).strip() == "" for v in row.values)

    df = df.loc[~df.apply(_all_blank, axis=1)].reset_index(drop=True)

    print(f"\n✅ Header row: {header_row+1}")
    print(f"✅ Company detected: {company}")
    print(f"✅ PO Number detected (header): {po_number}\n")

    # -----------------------------
    # Logo & QR setup
    # -----------------------------
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

    # -----------------------------
    # Layout constants
    # -----------------------------
    BASE_LABEL_W, BASE_LABEL_H = 58 * mm, 39 * mm
    COMPACT_LABEL_H = 37 * mm
    PAD = 0.5 * mm
    RIGHT_MARGIN = 3 * mm
    SPACING = 1.15
    output_pdf = "stickers_msg_final.pdf"

    # -----------------------------
    # Helpers
    # -----------------------------
    def extract_dpe_code(text):
        m = re.search(r'DPE\s*Item\s*Code\s*[:\-]?\s*([A-Za-z0-9\-_/]+)', str(text), re.I)
        return m.group(1).strip() if m else ""

    def clean_description(t):
        t = str(t)
        t = re.sub(r'(?i)dpe\s*item\s*code\s*[:\-]?\s*[A-Za-z0-9\-_/]+', '', t)
        t = re.sub(r'(?i)item\s*description\s*:?', '', t)
        t = " ".join(t.replace("\r", "\n").splitlines())
        t = re.sub(r'\s{2,}', ' ', t)
        t = re.sub(r'\s*,\s*', ', ', t)
        return t.strip()

    def wrap_text(text, font, size, width):
        if not text:
            return []
        words, lines, line = text.split(), [], ""
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

    def get_field_from_row(row, key):
        if not key:
            return ""
        keys = key if isinstance(key, (list, tuple)) else [key]
        norm_keys = []
        for k in keys:
            if not k:
                continue
            kk = k.strip().lower()
            norm_keys.extend([
                kk,
                kk.replace("number", "no"),
                kk.replace("no", "number"),
                kk.replace("qty", "quantity"),
                kk.replace("quantity", "qty"),
            ])
        for nk in norm_keys:
            if nk in row.index:
                val = row.get(nk, "")
                if pd.isna(val):
                    continue
                s = str(val).strip()
                if s:
                    return s
        return ""

    def fmt_qty(q):
        q = str(q).strip()
        if not q:
            return ""
        try:
            q_clean = q.replace(",", "")
            n = float(q_clean)
            return str(int(n)) if n.is_integer() else str(n).rstrip("0").rstrip(".")
        except Exception:
            return q

    # ===========================================
    # Generate Stickers
    # ===========================================
    c = canvas.Canvas(output_pdf, pagesize=(BASE_LABEL_W, BASE_LABEL_H))

    for _, row in df.iterrows():
        # Extract fields (handles varying column names)
        item_no = get_field_from_row(row, ["sl no", "item no", "item number"]) or ""
        dpe_code = get_field_from_row(row, ["dpe item code", "dpe code"]) or extract_dpe_code(get_field_from_row(row, "description"))
        desc = clean_description(get_field_from_row(row, ["description", "item description", "item name", "material"]))
        if not desc.strip():
            continue
        per_row_po = get_field_from_row(row, ["po number", "po no", "po"]) or ""
        po_qty = fmt_qty(get_field_from_row(row, ["po qty", "poqty", "po quantity", "qty", "quantity"])) or ""
        uom = get_field_from_row(row, ["uom", "unit"]) or ""
        heat = get_field_from_row(row, ["heat number", "heat no", "heat"]) or ""
        cert = get_field_from_row(row, ["certificate number", "certificate no", "cert"]) or ""
        make = get_field_from_row(row, ["make", "manufacturer"]) or ""
        remarks = get_field_from_row(row, ["remarks", "remark"]) or ""

        # Height depends on remarks
        LABEL_H = COMPACT_LABEL_H if not remarks.strip() else BASE_LABEL_H

        # --- parse PO quantity into integer number of copies ---
        qty = 1
        if po_qty:
            try:
                qty = int(float(str(po_qty)))
            except Exception:
                digits = re.sub(r"[^\d]", "", str(po_qty))
                try:
                    qty = int(digits) if digits else 1
                except Exception:
                    qty = 1
        if qty < 1:
            qty = 1
        MAX_QTY = 500
        if qty > MAX_QTY:
            qty = MAX_QTY

        # Build header/footer
        header_lines = []
        if company:
            header_lines.append(company)
        chosen_po = per_row_po if per_row_po else po_number
        if chosen_po:
            header_lines.append(f"PO Number: {chosen_po}")
        if item_no:
            header_lines.append(f"ITEM NO: {item_no}")
        if dpe_code:
            header_lines.append(f"DPE ITEM CODE: {dpe_code}")

        po_qty_full = f"PO QTY: {po_qty} {uom}".strip() if po_qty or uom else "PO QTY: "
        footer_fixed = [po_qty_full]
        if heat.strip():
            footer_fixed.append(f"HEAT NUMBER: {heat}".strip())
        footer_fixed.append(f"MAKE: {make}".strip() if make.strip() else "MAKE: ")
        if cert.strip():
            footer_fixed.append(f"CERTIFICATE NO: {cert}".strip())

        available_width = BASE_LABEL_W - 2 * PAD - 15
        min_size, max_size = 2.0, 7.0

        # Choose font size that fits all content
        chosen_font_size = min_size
        for font_size in [x / 10 for x in range(int(max_size * 10), int(min_size * 10) - 1, -1)]:
            wrapped_desc = wrap_text(desc, FONT, font_size, available_width)
            wrapped_remarks = wrap_text(f"Remarks: {remarks}", FONT, font_size, available_width) if remarks.strip() else []
            total_height = (len(header_lines) + len(wrapped_desc) + len(footer_fixed) + len(wrapped_remarks)) * font_size * SPACING
            if total_height <= LABEL_H - 0.3 * mm:
                chosen_font_size = font_size
                break

        font_size = chosen_font_size

        # Draw the same sticker page qty times (one PDF page per sticker)
        for _copy in range(qty):
            c.setPageSize((BASE_LABEL_W, LABEL_H))

            # Draw header
            c.setFont(FONT, font_size)
            y = LABEL_H - PAD - font_size
            y -= font_size * 0.1

            for line in header_lines:
                c.drawString(PAD, y, line)
                y -= font_size * SPACING
            y -= font_size * 0.4

            # Logo (optional)
            if logo_img:
                try:
                    c.drawImage(
                        logo_img,
                        BASE_LABEL_W - 12 * mm - RIGHT_MARGIN,
                        LABEL_H - 7 * mm - 1.5 * mm,
                        width=12 * mm,
                        height=7 * mm,
                        mask="auto",
                    )
                except Exception:
                    pass

            # Description
            y = min(y, LABEL_H - 3 * mm)
            for line in wrap_text(desc, FONT, font_size, available_width):
                c.drawString(PAD, y, line)
                y -= font_size * SPACING
            y -= font_size * 0.5

            # Footer
            for line in footer_fixed:
                for sub in wrap_text(line, FONT, font_size, available_width):
                    c.drawString(PAD, y, sub)
                    y -= font_size * SPACING

            # Remarks
            if remarks.strip():
                for line in wrap_text(f"Remarks: {remarks}", FONT, font_size, available_width):
                    c.drawString(PAD, y, line)
                    y -= font_size * SPACING

            # QR (optional)
            if qr_img:
                try:
                    c.drawImage(
                        qr_img,
                        BASE_LABEL_W - 6 * mm - RIGHT_MARGIN,
                        PAD,
                        width=6 * mm,
                        height=6 * mm,
                        mask="auto",
                    )
                except Exception:
                    pass

            c.showPage()

    # Save
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
