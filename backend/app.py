import io
import json
import math
import re
import zipfile
from datetime import date, datetime

import numpy as np
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook

app = FastAPI(title="POD Generator API", version="2.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

REQUIRED_COLUMNS = [
    "POD No", "Address", "Shipper", "Concern Receiver Name",
    "Description", "New POD Reff No",
]


def clean(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        return str(int(value)) if float(value).is_integer() else str(value)
    if isinstance(value, (datetime, date)):
        return value
    text = str(value).strip()
    try:
        number = float(text)
        if math.isfinite(number) and number.is_integer():
            return str(int(number))
    except (ValueError, TypeError):
        pass
    return text


def split_words(value, limit):
    words = str(value or "").split()
    lines, current = [], ""
    for word in words:
        if current and len(current) + len(word) + 1 > limit:
            lines.append(current)
            current = word
        else:
            current = f"{current} {word}".strip()
    if current:
        lines.append(current)
    return lines


def split_chars(value, limit=14):
    text = str(value or "").strip()
    return [text[i:i + limit] for i in range(0, len(text), limit)] if text else []


def parts(value, limit, character_split=False):
    return split_chars(value, limit) if character_split else split_words(value, limit)


def set_cell(ws, address, value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        value = ""
    ws[address] = value


def fill_template(template_bytes, row):
    workbook = load_workbook(io.BytesIO(template_bytes))
    ws = workbook.worksheets[0]
    pod_no = clean(row.get("POD No"))
    service = clean(row.get("Service")).upper()
    for key, cell in (("AIR", "K5"), ("TRAIN", "L5"), ("SURFACE", "M5")):
        if key in service:
            set_cell(ws, cell, "✔")

    set_cell(ws, "K3", " ".join(pod_no))
    ws["K3"].font = ws["K3"].font.copy(bold=True, sz=18)
    set_cell(ws, "N3", row.get("DATE"))

    shipper_lines = split_words(clean(row.get("Shipper")), 41)
    for i in range(1, 6):
        set_cell(ws, f"C{7+i}", shipper_lines[i-1] if len(shipper_lines) >= i else "")
    if len(shipper_lines) > 5:
        set_cell(ws, "C13", " ".join(shipper_lines[5:]))

    receiver_lines = split_words(clean(row.get("Concern Receiver Name")), 41)
    set_cell(ws, "K8", receiver_lines[0] if receiver_lines else "")
    address_lines = split_words(clean(row.get("Address")), 41)
    for i in range(1, 6):
        set_cell(ws, f"K{8+i}", address_lines[i-1] if len(address_lines) >= i else "")

    ref_lines = split_chars(clean(row.get("New POD Reff No")), 14)
    for i in range(1, 6):
        set_cell(ws, f"N{4+i}", ref_lines[i-1] if len(ref_lines) >= i else "")

    phone = row.get("Phone N.o")
    if phone is None or pd.isna(phone):
        phone = row.get("Phone No")
    contact = clean(row.get("Contact Person"))
    if phone is not None and not pd.isna(phone):
        phone_text = clean(phone)
        ws["K14"].number_format = "@"
        set_cell(ws, "K14", f"{contact} - {phone_text}" if contact else f"MOBILE - {phone_text}")

    description = split_words(clean(row.get("Description")), 27)
    for i in range(1, 6):
        set_cell(ws, f"B{17+i}", description[i-1] if len(description) >= i else "")

    for key, cell in (("L", "F23"), ("B", "G23"), ("H", "I23"),
                      ("Box", "G16"), ("Qty", "G18"),
                      ("Actual Weight", "J16"), ("Volume Weight", "J18")):
        value = row.get(key, "")
        if key == "Volume Weight" and value not in (None, "") and not pd.isna(value):
            value = math.floor(float(value) + 0.5)
        set_cell(ws, cell, value)
    set_cell(ws, "C10", "Mumbai")

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def safe_filename(value):
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", str(value)).strip(" ._")
    return cleaned or "POD"


@app.get("/")
def root():
    return {"status": "OK", "message": "POD Generator API is running"}


@app.post("/generate-pods")
async def generate_pods(
    data_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    data_bytes = await data_file.read()
    template_bytes = await template_file.read()
    if not data_bytes or not template_bytes:
        raise HTTPException(400, "Both Excel files must contain data.")
    if len(data_bytes) > 30 * 1024 * 1024 or len(template_bytes) > 30 * 1024 * 1024:
        raise HTTPException(413, "Each Excel file must be smaller than 30 MB.")
    try:
        dataframe = pd.read_excel(io.BytesIO(data_bytes), dtype=object)
        dataframe = dataframe.dropna(how="all").reset_index(drop=True)
    except Exception as exc:
        raise HTTPException(400, f"Could not read the data workbook: {exc}") from exc
    missing = [column for column in REQUIRED_COLUMNS if column not in dataframe.columns]
    if missing:
        raise HTTPException(400, "Missing required columns: " + ", ".join(missing))
    if dataframe.empty:
        raise HTTPException(400, "The data workbook has no delivery rows.")
    if dataframe["POD No"].isna().any():
        rows = (dataframe.index[dataframe["POD No"].isna()] + 2).tolist()
        raise HTTPException(400, "POD No is blank on spreadsheet row(s): " + ", ".join(map(str, rows[:20])))

    try:
        load_workbook(io.BytesIO(template_bytes), read_only=True).close()
    except Exception as exc:
        raise HTTPException(400, f"Could not read the POD template workbook: {exc}") from exc

    zip_buffer = io.BytesIO()
    used_names = set()
    failed_rows = []
    generated = 0
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for index, (_, row) in enumerate(dataframe.iterrows(), start=2):
            pod_no = clean(row.get("POD No"))
            try:
                workbook_bytes = fill_template(template_bytes, row)
                base_name = safe_filename(pod_no)
                filename = f"{base_name}.xlsx"
                duplicate = 1
                while filename.lower() in used_names:
                    filename = f"{base_name}_{duplicate}.xlsx"
                    duplicate += 1
                used_names.add(filename.lower())
                archive.writestr(filename, workbook_bytes)
                generated += 1
            except Exception as exc:
                failed_rows.append({"row": index, "pod": pod_no, "error": f"{type(exc).__name__}: {exc}"})
        manifest = {
            "generated": generated,
            "failed_count": len(failed_rows),
            "failed_rows": failed_rows,
            "total_rows": len(dataframe),
        }
        archive.writestr("generation-report.json", json.dumps(manifest, ensure_ascii=False, indent=2))

    if not generated:
        raise HTTPException(422, "No PODs could be generated. Check the template and source rows.")
    zip_buffer.seek(0)
    headers = {
        "Content-Disposition": 'attachment; filename="generated-pods.zip"',
        "X-POD-Generated": str(generated),
        "X-POD-Failed": str(len(failed_rows)),
        "Access-Control-Expose-Headers": "X-POD-Generated, X-POD-Failed, Content-Disposition",
    }
    return StreamingResponse(zip_buffer, media_type="application/zip", headers=headers)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
