import fitz 
import pytesseract
from PIL import Image
import io, re, os
import numpy as np
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage



# config
DPI = 400
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

HOUSE_RE = re.compile(r"(\d{1,3}\s*/\s*\d{1,4})")
TESSERACT_CONFIG = r"--oem 3 --psm 6"



def ocr_malayalam(pixmap):
    """Run OCR on image pixmap with Malayalam + English accuracy."""
    img = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("L")
    img = img.point(lambda x: 0 if x < 165 else 255, "1")  # binarize

    text = pytesseract.image_to_string(img, lang="mal+eng", config=TESSERACT_CONFIG)

    text = re.sub(r"[^0-9A-Za-z\u0D00-\u0D7F/\s]", "", text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

   
    if lines:
        top_crop = img.crop((0, 0, img.width, int(img.height * 0.18)))
        id_text = pytesseract.image_to_string(
            top_crop,
            lang="eng",
            config="--psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/",
        )
        id_text = re.sub(r"[^A-Za-z0-9/]", "", id_text).strip()
        if id_text:
            lines[0] = id_text

    return "\n".join(lines)


def parse_voter_lines(lines, pixmap=None):
    """Parse OCR lines into structured voter data."""
    lines = [l.strip() for l in lines if l.strip()]
    while len(lines) < 6:
        lines.append("")

   
    voter_id = ""
    id_match = re.search(r"[A-Z0-9/]{5,}", "".join(lines))
    if id_match:
        voter_id = id_match.group(0).strip(".")

    name = lines[1] if len(lines) > 1 else ""
    relation = lines[2] if len(lines) > 2 else ""
    house_no = lines[3] if len(lines) > 3 else ""
    house_name = ""
    gender_age = ""

    relation = re.sub(r"^(വ്‌|വ്|\u0D35\u0D4D\u200C)", "", relation).strip()

    
    if len(lines) > 4:
        if not re.search(r"\d{1,3}", lines[4]):
            house_name = lines[4].strip()
            if len(lines) > 5 and not re.search(r"\d{1,3}", lines[5]):
                house_name += " " + lines[5].strip()

    # OCR last line again (Malayalam only) for gender
    if pixmap is not None:
        img = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("L")
        h = img.height
        last_line_crop = img.crop((0, int(h * 0.82), img.width, h))  # bottom ~18%

        # Malayalam OCR
        mal_text = pytesseract.image_to_string(last_line_crop, lang="mal", config="--psm 7")
        mal_text = re.sub(r"[^0-9A-Za-z\u0D00-\u0D7F/\s]", "", mal_text).strip()

        # English OCR - age 
        eng_text = pytesseract.image_to_string(last_line_crop, lang="eng", config="--psm 7")
        age_match = re.search(r"(\d{2,3})", eng_text)
        age_str = age_match.group(1) if age_match else ""

        
        if re.search(r"\d{2,3}", mal_text):
            gender_age = mal_text  
        else:
            gender_age = mal_text
            if age_str:
                gender_age += f" / {age_str}"

    return voter_id, name, relation, house_no, house_name, gender_age



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


def process_voter_pdf(pdf_bytes, house_no_input, mode):
    """Core logic used by FastAPI to filter voters."""
    temp_pdf = os.path.join(OUTPUT_DIR, "temp_input.pdf")
    with open(temp_pdf, "wb") as f:
        f.write(pdf_bytes)

    doc = fitz.open(temp_pdf)
    collected, all_voters = [], []

    def process_page(page_idx):
        results_local = []
        crops_local = []
        try:
            page = doc.load_page(page_idx)
            prefix = house_no_input.split("/")[0] + "/"
            if prefix not in page.get_text():
                return [], []

            pix = page.get_pixmap(dpi=DPI)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            img_arr = np.array(img)
            rects = [d["rect"] for d in page.get_drawings() if d.get("rect")]
            scale = DPI / 72.0

            for rect_index, r in enumerate(rects):
                x0, y0, x1, y1 = int(r.x0 * scale), int(r.y0 * scale), int(r.x1 * scale), int(r.y1 * scale)
                full_crop = img_arr[y0:y1, x0:x1]
                sub_crop = full_crop[:, :int(full_crop.shape[1] * 0.5)]
                ocr_text = pytesseract.image_to_string(sub_crop, lang="mal+eng", config=TESSERACT_CONFIG)
                matches = [m.replace(" ", "") for m in HOUSE_RE.findall(ocr_text)]
                if house_no_input not in matches:
                    continue

                width, height = r.x1 - r.x0, r.y1 - r.y0
                value_area = fitz.Rect(r.x0 + width * 0.26, r.y0, r.x1 - width * 0.2, r.y1)
                pix_val = page.get_pixmap(clip=value_area, dpi=DPI)
                text = ocr_malayalam(pix_val)
                if not text.strip():
                    continue

                lines = [l.strip() for l in text.split("\n") if l.strip()]
                voter_id, name, relation, house_no, house_name, gender_age = parse_voter_lines(lines, pix_val)

                results_local.append({
                    "VoterID": voter_id,
                    "Name": name,
                    "Relation": relation,
                    "HouseNo": house_no.strip(),
                    "HouseName": house_name.strip(),
                    "GenderAge": gender_age.strip(),
                    "Page": page_idx + 1,
                    "RectIndex": rect_index,
                })
                crops_local.append(full_crop)
        except Exception as e:
            print(f"[Page {page_idx}] Error: {e}")
        return results_local, crops_local


    # Run in parallel threads
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(process_page, i): i for i in range(1, len(doc))}
        for future in tqdm(as_completed(futures), total=len(futures), desc="🔄 OCR Pages", unit="page"):
            res, crops = future.result()
            all_voters.extend(res)
            collected.extend(crops)




    if not collected:
        raise Exception(f"No matching voters found for house number {house_no_input}")

    results = {}
    if mode == "xlsx":
        xlsx_path = os.path.join(OUTPUT_DIR, f"voter_data_{house_no_input.replace('/', '-')}.xlsx")

        # Save temp images + record their paths
        image_paths = []
        for idx, crop in enumerate(collected):
            img_path = os.path.join(OUTPUT_DIR, f"voter_{house_no_input.replace('/', '-')}_{idx+1}.png")
            Image.fromarray(crop).save(img_path)
            image_paths.append(img_path)

        # Combine data with image path column
        df = pd.DataFrame(all_voters)
        df.insert(0, "Image", image_paths)  # new first column
        df.to_excel(xlsx_path, index=False, engine="openpyxl")

        # Embed images into Excel cells
        wb = load_workbook(xlsx_path)
        ws = wb.active
        for i, img_path in enumerate(image_paths, start=2):  # header row is 1
            try:
                img = XLImage(img_path)
                img.height = 90
                img.width = 120
                ws.row_dimensions[i].height = 70
                ws.add_image(img, f"A{i}")  # insert in column A
            except Exception as e:
                print(f"Failed to embed image for row {i}: {e}")

        wb.save(xlsx_path)
        results["xlsx"] = xlsx_path

    elif mode == "pdf":
        pdf_path = os.path.join(OUTPUT_DIR, f"filtered_{house_no_input.replace('/', '-')}.pdf")
        save_filtered_pdf(collected, pdf_path)
        results["pdf"] = pdf_path

    doc.close()
    return results
