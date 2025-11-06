from fastapi import FastAPI, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
from core import process_voter_pdf
from enum import Enum
from io import BytesIO

app = FastAPI(title="Voter OCR & Filter API")

class OutputMode(str, Enum):
    pdf = "pdf"
    xlsx = "xlsx"


@app.post("/process")
async def process_pdf(
    file: UploadFile,
    house_no: str = Form(...),
    mode: OutputMode = Form(...)
):
    if mode not in ("pdf", "xlsx"):
        raise HTTPException(status_code=400, detail="Invalid mode. Must be 'pdf' or 'xlsx'.")

    # Read uploaded PDF bytes
    pdf_bytes = await file.read()

    # Process it (your custom OCR/filter logic)
    results = process_voter_pdf(pdf_bytes, house_no, mode)

    # For PDF output
    if mode == "pdf" and "pdf_bytes" in results:
        pdf_stream = BytesIO(results["pdf_bytes"])
        pdf_stream.seek(0)
        return StreamingResponse(
            pdf_stream,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="filtered_{house_no}.pdf"'}
        )

    # For Excel output
    elif mode == "xlsx" and "xlsx_bytes" in results:
        xlsx_stream = BytesIO(results["xlsx_bytes"])
        xlsx_stream.seek(0)
        return StreamingResponse(
            xlsx_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="voter_data_{house_no}.xlsx"'}
        )

    else:
        raise HTTPException(status_code=500, detail="Output not generated.")
