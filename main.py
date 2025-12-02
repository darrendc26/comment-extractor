from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
import json, io
from extractor import extract_pdf_comments

app = FastAPI()

@app.post("/extract-comments")
async def extract_comments(
    comment_csv: str = Form(...),      # now a single string
    company_codes: str = Form(...),      # now a single string
    pdfs: list[UploadFile] = File(...)
):
    # Split by comma and strip spaces
    codes = [c.strip() for c in company_codes.split(",")]

    if len(codes) != len(pdfs):
        return {
            "error": "Company string count and PDF count mismatch",
        }

    file_data = []

    for i in range(len(pdfs)):
        pdf = pdfs[i]
        company = codes[i]

        bytes_data = await pdf.read()
        file_data.append((company, bytes_data, pdf.filename))

    csv_string = extract_pdf_comments(file_data)

    return StreamingResponse(
        io.BytesIO(csv_string.encode("utf-8")),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename={}".format(comment_csv)}
    )


@app.post("/test")
async def test(
    company_codes: list[str] = Form(...),
    pdfs: list[UploadFile] = File(...)
):
    return {
        "received_company_codes": company_codes,
        "num_company_codes": len(company_codes),
        "received_pdfs": [f.filename for f in pdfs],
        "num_pdfs": len(pdfs)
    }