# SKUSheet

Google Apps Script that takes a list of SKUs and lays them out nicely on a Google Sheet for printing — generates barcodes, product info, and page-friendly layouts inside the spreadsheet.

## Setup

1. Open a Google Sheet
2. **Extensions → Apps Script**
3. Paste [`code.gs`](code.gs) and [`appscript.json`](appscript.json) into the script editor
4. Save and refresh the sheet — a custom "Barcode Tools" menu will appear

## Features

- Checkbox-driven per-row controls in column A
- Configurable box dimensions (width × height), boxes-per-row, rows-per-page
- Optional product image insertion
- Color-coded status cells (idle / queued / busy / success / error)
- Amazon product lookup via SerpAPI (optional — set `SERPAPI_KEY` in Script Properties)

## Configuration

Open the sheet's Apps Script project → **Project Settings → Script Properties**, and set:

| Key | Purpose |
|---|---|
| `SERPAPI_KEY` | Optional — enables Amazon product lookup for each SKU |

## License

MIT.
