from fastapi import FastAPI, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from core import process_voter_pdf
from enum import Enum
import os

app = FastAPI(title="Voter OCR & Filter API")

class OutputMode(str, Enum):
    pdf = "pdf"
    xlsx = "xlsx"


def delete_file(path: str):
    """Delete a file safely."""
    try:
        if os.path.exists(path):
            os.remove(path)
            print(f" Deleted temp file: {path}")
    except Exception as e:
        print(f" Error deleting {path}: {e}")


@app.post("/process")
async def process_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile,
    house_no: str = Form(...),
    mode: OutputMode = Form(...)
):
    if mode not in ("pdf", "xlsx"):
        raise HTTPException(status_code=400, detail="Invalid mode. Must be 'pdf' or 'xlsx'.")

    pdf_bytes = await file.read()
    results = process_voter_pdf(pdf_bytes, house_no, mode)

    if mode == "pdf" and "pdf" in results:
        file_path = results["pdf"]
        background_tasks.add_task(delete_file, file_path)
        return FileResponse(
            path=file_path,
            filename=f"filtered_{house_no}.pdf",
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="filtered_{house_no}.pdf"'}
        )

    elif mode == "xlsx" and "xlsx" in results:
        file_path = results["xlsx"]
        background_tasks.add_task(delete_file, file_path)
        return FileResponse(
            path=file_path,
            filename=f"voter_data_{house_no}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="voter_data_{house_no}.xlsx"'}
        )

    else:
        raise HTTPException(status_code=500, detail="Output file not generated.")
