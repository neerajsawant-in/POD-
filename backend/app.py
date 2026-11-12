import os
import time
import shutil
import tempfile
import traceback
import re
from pathlib import Path

import pandas as pd
import openpyxl
import numpy as np
import uvicorn
import pythoncom
import win32com.client as win32

from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware

# ---------- Configuration ----------
DOWNLOADS = str(Path.home() / "Downloads")
TEMP_DIR = tempfile.gettempdir()
DATA_PATH = os.path.join(TEMP_DIR, "data.xlsx")
TEMPLATE_PATH = os.path.join(TEMP_DIR, "template.xlsx")

app = FastAPI(title="POD Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def root():
    return {"status": "OK", "message": "POD Generator API is running"}


# ---------- Helpers ----------
def split_text_41(text, limit=41):
    """Word-aware split (used for Address / Shipper / Receiver / Description)."""
    words = str(text).split()
    if not words:
        return []
    lines = []
    current = words[0]
    for word in words[1:]:
        if len(current) + 1 + len(word) <= limit:
            current += " " + word
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def split_text_chars(text, limit=14):
    """
    Pure CHARACTER-based split. No word logic.
    Every `limit` characters become a new line.
    Example: "Reff-11018214397400" with limit=14
             -> ["Reff-110182143", "97400"]
    """
    text = str(text).strip()
    if not text or text.lower() == "nan":
        return []
    return [text[i:i + limit] for i in range(0, len(text), limit)]


def safe_pod_string(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        return str(int(value)) if float(value).is_integer() else str(value)
    s = str(value).strip()
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except (ValueError, TypeError):
        pass
    return s


def safe_filename(text):
    return re.sub(r'[\\/:*?"<>|]', '_', str(text))


def process_splits(dataframe, col_name, prefix, limit):
    """Word-aware split."""
    if col_name not in dataframe.columns:
        return
    all_rows = []
    max_p = 0
    for text in dataframe[col_name].fillna(""):
        parts = split_text_41(str(text), limit)
        max_p = max(max_p, len(parts))
        all_rows.append(parts)
    for i in range(max_p):
        dataframe[f"{prefix}_{i+1}"] = ""
    for i, parts in enumerate(all_rows):
        for j, part in enumerate(parts):
            dataframe.at[i, f"{prefix}_{j+1}"] = part


def process_splits_chars(dataframe, col_name, prefix, limit=14):
    """Character-based split (new function — hard cut every N chars)."""
    if col_name not in dataframe.columns:
        return
    all_rows = []
    max_p = 0
    for text in dataframe[col_name].fillna(""):
        parts = split_text_chars(str(text), limit)
        max_p = max(max_p, len(parts))
        all_rows.append(parts)
    for i in range(max_p):
        dataframe[f"{prefix}_{i+1}"] = ""
    for i, parts in enumerate(all_rows):
        for j, part in enumerate(parts):
            dataframe.at[i, f"{prefix}_{j+1}"] = part


# ---------- MAIN API ----------
@app.post("/generate-pods")
async def generate_pods(
    data_file: UploadFile = File(...),
    template_file: UploadFile = File(...)
):
    try:
        with open(DATA_PATH, "wb") as f:
            f.write(await data_file.read())
        with open(TEMPLATE_PATH, "wb") as f:
            f.write(await template_file.read())
    except Exception as e:
        traceback.print_exc()
        return {"status": "ERROR", "message": f"File save failed: {e}"}

    try:
        df = pd.read_excel(DATA_PATH)
    except Exception as e:
        traceback.print_exc()
        return {"status": "ERROR", "message": f"Data Excel read failed: {e}"}

    required_cols = [
        "POD No", "Address", "Shipper", "Concern Receiver Name",
        "Description", "New POD Reff No",
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        return {
            "status": "ERROR",
            "message": f"Missing columns: {missing}. Available: {list(df.columns)}",
        }

    if df["POD No"].isnull().any():
        bad = int(df["POD No"].isnull().sum())
        return {"status": "ERROR", "message": f"{bad} rows mein 'POD No' khali hai."}

    df["POD_ST"] = df["POD No"].apply(safe_pod_string)
    df = df[df["POD_ST"] != ""].reset_index(drop=True)
    if df.empty:
        return {"status": "ERROR", "message": "Koi valid POD row nahi mili."}

    # Word-aware splits (as before)
    process_splits(df, "Address", "ADDR", 41)
    process_splits(df, "Shipper", "Shipper", 41)
    process_splits(df, "Concern Receiver Name", "Receiver", 41)
    process_splits(df, "Description", "Desc", 27)

    # Character-aware split — hard cut every 14 chars
    process_splits_chars(df, "New POD Reff No", "New Reff", 14)

    if "Volume Weight" in df.columns:
        df["Volume Weight"] = np.floor(
            pd.to_numeric(df["Volume Weight"], errors="coerce") + 0.5
        ).astype("Int64")

    pod_nums = pd.to_numeric(df["POD No"], errors="coerce").dropna()
    if not pod_nums.empty:
        pod_min = int(pod_nums.min())
        pod_max = int(pod_nums.max())
    else:
        pod_min, pod_max = df["POD_ST"].min(), df["POD_ST"].max()

    RUN_FOLDER = os.path.join(DOWNLOADS, f"PODs_{pod_min}_to_{pod_max}")

    if os.path.exists(RUN_FOLDER):
        try:
            shutil.rmtree(RUN_FOLDER)
        except Exception:
            RUN_FOLDER = f"{RUN_FOLDER}_{int(time.time())}"

    os.makedirs(RUN_FOLDER, exist_ok=True)

    PROCESSED_FILE = os.path.join(RUN_FOLDER, "Processed_POD_Data.xlsx")
    try:
        df.to_excel(PROCESSED_FILE, index=False)
    except Exception as e:
        traceback.print_exc()
        return {"status": "ERROR", "message": f"Processed file save failed: {e}"}

    # ---------- EXCEL GENERATION ----------
    pythoncom.CoInitialize()
    excel = None
    pod_count = 0
    failed_rows = []
    data_wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        data_wb = openpyxl.load_workbook(PROCESSED_FILE, data_only=True)
        data_ws = data_wb.worksheets[0]
        headers = [cell.value for cell in data_ws[1]]

        for excel_row in data_ws.iter_rows(min_row=2, values_only=True):
            row = dict(zip(headers, excel_row))
            pod_no = str(row.get("POD_ST", "")).strip()
            if not pod_no:
                continue

            wb = None
            try:
                wb = excel.Workbooks.Open(TEMPLATE_PATH)
                ws = wb.Worksheets(1)

                service = str(row.get("Service", "")).strip().upper()
                if service == "AIR":
                    ws.Range("K5").Value = "✔"
                elif service == "TRAIN":
                    ws.Range("L5").Value = "✔"
                elif service == "SURFACE":
                    ws.Range("M5").Value = "✔"

                ws.Range("K3").Value = " ".join(pod_no)
                ws.Range("K3").Font.Bold = True
                ws.Range("K3").Font.Size = 18
                ws.Range("N3").Value = row.get("DATE")

                for i in range(1, 6):
                    ws.Range(f"C{7+i}").Value = row.get(f"Shipper_{i}", "")

                extra_shipper = row.get("Shipper_Extra")
                if extra_shipper not in (None, "") and pd.notna(extra_shipper):
                    ws.Range("C13").Value = extra_shipper

                ws.Range("K8").Value = row.get("Receiver_1", "")
                for i in range(1, 6):
                    ws.Range(f"K{8+i}").Value = row.get(f"ADDR_{i}", "")

                # New Reff lines -> N5, N6, N7, N8, N9
                for i in range(1, 6):
                    ws.Range(f"N{4+i}").Value = row.get(f"New Reff_{i}", "")

                # --- PHONE + CONTACT PERSON ---
                phone = row.get("Phone N.o")
                if phone is None or (isinstance(phone, float) and pd.isna(phone)):
                    phone = row.get("Phone No")

                contact = row.get("Contact Person")
                contact_str = "" if contact is None or (isinstance(contact, float) and pd.isna(contact)) else str(contact).strip()

                if phone is not None and not (isinstance(phone, float) and pd.isna(phone)):
                    phone_str = str(phone).strip()
                    ws.Range("K14").NumberFormat = "@"
                    if contact_str:
                        ws.Range("K14").Value = f"{contact_str} - {phone_str}"
                    else:
                        ws.Range("K14").Value = f"MOBILE - {phone_str}"

                for i in range(1, 6):
                    ws.Range(f"B{17+i}").Value = row.get(f"Desc_{i}", "")

                ws.Range("F23").Value = row.get("L", "")
                ws.Range("G23").Value = row.get("B", "")
                ws.Range("I23").Value = row.get("H", "")
                ws.Range("G16").Value = row.get("Box", "")
                ws.Range("G18").Value = row.get("Qty", "")
                ws.Range("J16").Value = row.get("Actual Weight", "")
                ws.Range("J18").Value = row.get("Volume Weight", "")

                ws.Range("C10").Value = "Mumbai"

                base_name = safe_filename(pod_no)
                excel_path = os.path.join(RUN_FOLDER, f"{base_name}.xlsx")

                dup = 1
                while os.path.exists(excel_path):
                    excel_path = os.path.join(RUN_FOLDER, f"{base_name}_{dup}.xlsx")
                    dup += 1

                wb.SaveAs(excel_path, FileFormat=51)
                pod_count += 1

            except Exception as row_err:
                err_msg = f"{type(row_err).__name__}: {row_err}"
                failed_rows.append({"pod": pod_no, "error": err_msg})
                print(f"[ROW FAIL] POD={pod_no} -> {err_msg}")
                traceback.print_exc()
            finally:
                if wb is not None:
                    try:
                        wb.Close(False)
                    except Exception:
                        pass

    except Exception as e:
        traceback.print_exc()
        return {
            "status": "ERROR",
            "message": f"Excel COM failed: {type(e).__name__}: {e}",
            "generated_excels": pod_count,
            "folder": RUN_FOLDER,
            "failed_rows": failed_rows,
        }
    finally:
        if data_wb is not None:
            try:
                data_wb.close()
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    status = "SUCCESS" if pod_count > 0 else "ERROR"
    msg = None
    if pod_count == 0:
        msg = "Koi bhi POD generate nahi hua. Template/Excel check karo."
    elif failed_rows:
        msg = f"{len(failed_rows)} rows fail hui."

    return {
        "status": status,
        "message": msg,
        "generated_excels": pod_count,
        "failed_count": len(failed_rows),
        "failed_rows": failed_rows,
        "folder": RUN_FOLDER,
        "processed_file": PROCESSED_FILE,
    }


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)