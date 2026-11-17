# POD Studio

POD Studio turns delivery records in an Excel spreadsheet into individual Proof of Delivery workbooks. Upload a delivery data file and your existing POD template, generate the documents, and download them together as a ZIP.

## Features

- Upload `.xlsx` data and template files by browsing or drag and drop.
- Generate one POD workbook for each delivery row.
- Keep the template's existing layout and formatting while filling its first worksheet.
- Download generated workbooks and a `generation-report.json` summary in one ZIP.
- Review validation feedback in the app and row-level generation errors in the report.

## Requirements

- Python 3.10 or newer
- Node.js and npm
- An `.xlsx` delivery spreadsheet and an `.xlsx` POD template

## Run locally

Start the backend in one terminal:

```bash
cd backend
python -m pip install -r requirements.txt
python app.py
```

The API starts at `http://127.0.0.1:8000`.

Start the frontend in a second terminal:

```bash
cd frontend
npm install
npm run dev
```

Open the local address printed by Vite, usually `http://localhost:5173`.

Run the frontend checks with:

```bash
cd frontend
npm run build
npm run lint
```

## Prepare the delivery spreadsheet

The first worksheet must contain a header row and these required columns:

| Column | Use |
| --- | --- |
| `POD No` | POD number printed on the document; every delivery row needs one. |
| `Address` | Delivery address. |
| `Shipper` | Shipper name and address. |
| `Concern Receiver Name` | Receiver name. |
| `Description` | Shipment description. |
| `New POD Reff No` | Reference number printed on the POD. |

Optional columns recognized by the generator:

- `DATE`, `Service` (`AIR`, `TRAIN`, or `SURFACE`)
- `Phone N.o` or `Phone No`, `Contact Person`
- `Box`, `Qty`, `Actual Weight`, `Volume Weight`
- `L`, `B`, `H`

Completely blank rows are ignored. Each row with a POD number generates a separate workbook. Files larger than 30 MB are rejected.

## POD template and field placement

The template is not included because each organization uses its own POD form. Upload an `.xlsx` template whose first worksheet follows the cell layout expected by the generator. The current layout fills POD number and date, service selection, shipper and receiver details, address, reference number, contact details, description, dimensions, quantity, and weights.

## Output

The browser downloads `generated-pods.zip`. It contains one `.xlsx` file per successfully processed delivery row and `generation-report.json`, which lists the total, generated, and failed rows. Duplicate POD numbers receive unique filenames.

## Project layout

```text
backend/
  app.py              FastAPI upload and Excel generation API
  requirements.txt    Python dependencies
frontend/
  src/                React interface and styles
  package.json        Frontend scripts and dependencies
Trail Dataset.xlsx    Sample delivery spreadsheet
```
