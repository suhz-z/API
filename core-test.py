import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io, os, re
import numpy as np
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook

# ========== CONFIG ==========
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)
DPI = 400
MAX_WORKERS = 6
TESSERACT_CONFIG = "--psm 6"
HOUSE_RE = re.compile(r"(\d{1,3}\s*/\s*\d{1,4})")
# ============================


# 🧠 Helper: Malayalam OCR
def ocr_malayalam(pix):
    """Extract text from a PyMuPDF pixmap (Malayalam + English)."""
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
    text = pytesseract.image_to_string(img, lang="mal+eng", config=TESSERACT_CONFIG)
    return text


# 🧩 Helper: Parse lines (minimal placeholder — extend if needed)
def parse_voter_lines(lines, pix_val):
    voter_id = lines[0] if lines else ""
    name = lines[1] if len(lines) > 1 else ""
    relation = lines[2] if len(lines) > 2 else ""
    house_no = lines[3] if len(lines) > 3 else ""
    house_name = lines[4] if len(lines) > 4 else ""
    gender_age = lines[-1] if len(lines) > 5 else ""
    return voter_id, name, relation, house_no, house_name, gender_age


# 🧱 PDF Export
def save_filtered_pdf(collected, output_path):
    """Save filtered voter boxes as 2x10 grid PDF."""
    out_doc = fitz.open()
    page_width, page_height = 595, 842
    boxes_per_row, boxes_per_page = 2, 20
    voter_width = page_width // boxes_per_row
    voter_height = page_height // (boxes_per_page // boxes_per_row)

    count = 0
    for crop in collected:
        if count % boxes_per_page == 0:
            out_page = out_doc.new_page(width=page_width, height=page_height)
            x_pt, y_pt = 0, 0

        pil = Image.fromarray(crop).resize((int(voter_width - 10), int(voter_height - 10)))
        buf = io.BytesIO()
        pil.save(buf, format="PNG")
        rect = fitz.Rect(x_pt + 5, y_pt + 5, x_pt + voter_width - 5, y_pt + voter_height - 5)
        out_page.insert_image(rect, stream=buf.getvalue())

        count += 1
        if count % boxes_per_row == 0:
            x_pt = 0
            y_pt += voter_height
        else:
            x_pt += voter_width

    out_doc.save(output_path)
    out_doc.close()


# 🧮 Excel Export
def save_filtered_xlsx(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)


# 🧠 OCR Worker
def process_box(page, r, scale, img_arr, house_no_input, page_idx, rect_index):
    """Processes one voter box with value-area cropping."""
    try:
        # --- Full image crop (for saving) ---
        x0, y0, x1, y1 = int(r.x0 * scale), int(r.y0 * scale), int(r.x1 * scale), int(r.y1 * scale)
        full_crop = img_arr[y0:y1, x0:x1]

        # --- Initial half-box OCR check ---
        sub_crop = full_crop[:, :int(full_crop.shape[1] * 0.5)]
        ocr_text = pytesseract.image_to_string(sub_crop, lang="mal+eng", config=TESSERACT_CONFIG)
        matches = [m.replace(" ", "") for m in HOUSE_RE.findall(ocr_text)]
        if house_no_input not in matches:
            return None

        # --- Value area crop (same as your original) ---
        width, height = r.x1 - r.x0, r.y1 - r.y0
        value_area = fitz.Rect(r.x0 + width * 0.26, r.y0, r.x1 - width * 0.2, r.y1)
        pix_val = page.get_pixmap(clip=value_area, dpi=DPI)
        text = ocr_malayalam(pix_val)
        if not text.strip():
            return None

        # --- Parse extracted lines ---
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        voter_id, name, relation, house_no, house_name, gender_age = parse_voter_lines(lines, pix_val)

        return {
            "crop": full_crop,
            "data": {
                "VoterID": voter_id,
                "Name": name,
                "Relation": relation,
                "HouseNo": house_no.strip(),
                "HouseName": house_name.strip(),
                "GenderAge": gender_age.strip(),
                "Page": page_idx + 1,
                "RectIndex": rect_index,
            },
        }
    except Exception as e:
        print(f"[Page {page_idx} | Rect {rect_index}] Error: {e}")
        return None


# 🚀 Core processing logic (used by FastAPI)
def process_voter_pdf(pdf_bytes, house_no_input, mode="xlsx"):
    """Filter voter boxes and export as PDF or XLSX."""
    temp_pdf = os.path.join(OUTPUT_DIR, "temp_input.pdf")
    with open(temp_pdf, "wb") as f:
        f.write(pdf_bytes)

    doc = fitz.open(temp_pdf)
    collected, all_voters = [], []

    for page_idx in tqdm(range(1, len(doc)), desc="📄 Pages", unit="page"):
        page = doc.load_page(page_idx)
        prefix = house_no_input.split("/")[0] + "/"
        if prefix not in page.get_text():
            continue

        pix = page.get_pixmap(dpi=DPI)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        img_arr = np.array(img)
        rects = [d["rect"] for d in page.get_drawings() if d.get("rect")]
        scale = DPI / 72.0

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = [
                executor.submit(process_box, page, r, scale, img_arr, house_no_input, page_idx, i)
                for i, r in enumerate(rects)
            ]
            for future in futures:
                res = future.result()
                if res:
                    collected.append(res["crop"])
                    all_voters.append(res["data"])

    # === Export based on mode ===
    safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", house_no_input)
    if mode == "pdf":
        pdf_path = os.path.join(OUTPUT_DIR, f"filtered_{safe_name}.pdf")
        save_filtered_pdf(collected, pdf_path)
        return pdf_path
    else:
        xlsx_path = os.path.join(OUTPUT_DIR, f"filtered_{safe_name}.xlsx")
        save_filtered_xlsx(all_voters, xlsx_path)
        return xlsx_path
