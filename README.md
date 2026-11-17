# POD Studio

A local web app for generating one Excel Proof of Delivery workbook per delivery row.

## Start the backend

From the project directory, install the Python packages once and start FastAPI:

```bash
cd backend
python -m pip install -r requirements.txt
python app.py
```

The API runs at `http://127.0.0.1:8000`.

## Start the frontend

In a second terminal:

```bash
cd frontend
npm install
npm run dev
```

Open the local URL printed by Vite (usually `http://localhost:5173`).

## Generate PODs

Upload an `.xlsx` delivery workbook and an `.xlsx` POD template. The data workbook must have these columns: `POD No`, `Address`, `Shipper`, `Concern Receiver Name`, `Description`, and `New POD Reff No`. Other supported columns are `DATE`, `Service`, `Phone N.o` (or `Phone No`), `Contact Person`, `Box`, `Qty`, `Actual Weight`, `Volume Weight`, `L`, `B`, and `H`.

The generator fills the first worksheet of the template using the existing cell layout, then downloads a ZIP containing one `.xlsx` per valid row and a `generation-report.json` summary. The template itself is not included in the project; supply your usual POD template when generating a batch.
