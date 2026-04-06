---
name: gsheets-edit
description: Edit a Google Sheets spreadsheet using natural language instructions
user-invocable: true
---

Edit a Google Sheets spreadsheet using the Google Sheets REST API authenticated via `gcloud auth print-access-token`.

**Usage:** `/gsheets-edit <spreadsheet_id> <instruction>`

**Example:** `/gsheets-edit 1NNxWj4WByNfCPNPbc-G2o-GwaMHG99Qy2tBqMPqQhDs "Bold the header row and add a SUM formula in the totals row"`

---

## How to execute

### Step 1 — Get auth token
Run:
```bash
TOKEN=$(gcloud auth print-access-token)
```
Store this token. It expires after ~1 hour.

### Step 2 — Read the spreadsheet metadata
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID?fields=properties.title,sheets.properties"
```
This returns the spreadsheet title and all sheet (tab) names, IDs, and row/column counts.

### Step 3 — Read sheet data
Read values from a specific range:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values/Sheet1!A1:Z100"
```

Read multiple ranges at once:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values:batchGet?ranges=Sheet1!A1:Z1&ranges=Sheet1!A2:Z100"
```

Read formulas (not computed values):
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values/Sheet1!A1:Z100?valueRenderOption=FORMULA"
```

**Important:** URL-encode tab names with special characters:
```bash
# Tab named "Total projs/spend" → URL encode the slash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values/Total%20projs%2Fspend!A1:Z10"
```

**Important:** When tab names contain spaces, `curl` handles URL encoding automatically, but Python `urllib` does not. For programmatic reads/writes with special tab names, use `urllib.parse.quote()` on the range but preserve `!` and `:`:
```python
import urllib.parse
range_str = urllib.parse.quote("How to Read!A1:D50", safe="!:")
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SSID}/values/{range_str}?valueInputOption=RAW"
```
Alternatively, use `curl` via subprocess for simpler URL handling.

### Step 4 — Plan the edits
Based on the instruction and spreadsheet content, decide which API operations to use:

| Goal | API method |
|------|-----------|
| Write values to cells | `values.update` or `values.batchUpdate` |
| Append rows after last data | `values.append` |
| Clear cell values (keep formatting) | `values.clear` |
| Bold / italic / font / colors | `batchUpdate` → `repeatCell` or `updateCells` |
| Number format (currency, %, date) | `batchUpdate` → `repeatCell` with `numberFormat` |
| Merge cells | `batchUpdate` → `mergeCells` |
| Set column width / row height | `batchUpdate` → `updateDimensionProperties` |
| Freeze rows/columns | `batchUpdate` → `updateSheetProperties` with `gridProperties` |
| Add a new sheet tab | `batchUpdate` → `addSheet` |
| Delete a sheet tab | `batchUpdate` → `deleteSheet` |
| Rename a sheet tab | `batchUpdate` → `updateSheetProperties` |
| Sort a range | `batchUpdate` → `sortRange` |
| Auto-resize columns to fit | `batchUpdate` → `autoResizeDimensions` |
| Add conditional formatting | `batchUpdate` → `addConditionalFormatRule` |
| Create a new spreadsheet | `spreadsheets.create` (POST) |

**Important rules:**
- Use A1 notation for ranges (e.g., `Sheet1!A1:D10`)
- When writing values, use `valueInputOption=USER_ENTERED` so formulas (=SUM, etc.) are parsed
- Sheet IDs (integers) are needed for formatting; get them from the metadata call in Step 2
- Row/column indices in formatting requests are 0-based (row 1 = index 0)
- Always read before editing to understand the current structure

### Step 5 — Write values
**Update a range:**
```bash
curl -s -X PUT \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"values": [["Header1", "Header2"], ["val1", "val2"]]}' \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values/Sheet1!A1:B2?valueInputOption=USER_ENTERED"
```

**Append rows:**
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"values": [["new1", "new2"], ["new3", "new4"]]}' \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values/Sheet1!A:B:append?valueInputOption=USER_ENTERED"
```

**Batch update multiple ranges:**
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "valueInputOption": "USER_ENTERED",
    "data": [
      {"range": "Sheet1!A1:B1", "values": [["Name", "Score"]]},
      {"range": "Sheet1!A10:B10", "values": [["Total", "=SUM(B2:B9)"]]}
    ]
  }' \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values:batchUpdate"
```

### Step 6 — Format cells via batchUpdate
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"requests": [...]}' \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID:batchUpdate"
```

Example requests:

**Bold the header row (row 1):**
```json
{
  "repeatCell": {
    "range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 1},
    "cell": {"userEnteredFormat": {"textFormat": {"bold": true}}},
    "fields": "userEnteredFormat.textFormat.bold"
  }
}
```

**Set number format (currency):**
```json
{
  "repeatCell": {
    "range": {"sheetId": 0, "startRowIndex": 1, "endRowIndex": 100, "startColumnIndex": 2, "endColumnIndex": 3},
    "cell": {"userEnteredFormat": {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}}},
    "fields": "userEnteredFormat.numberFormat"
  }
}
```

**Set background color:**
```json
{
  "repeatCell": {
    "range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 1},
    "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.0, "green": 0.5, "blue": 0.68}}},
    "fields": "userEnteredFormat.backgroundColor"
  }
}
```

**Freeze the first row:**
```json
{
  "updateSheetProperties": {
    "properties": {"sheetId": 0, "gridProperties": {"frozenRowCount": 1}},
    "fields": "gridProperties.frozenRowCount"
  }
}
```

**Merge cells:**
```json
{
  "mergeCells": {
    "range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 3},
    "mergeType": "MERGE_ALL"
  }
}
```

**Auto-resize columns to fit content:**
```json
{
  "autoResizeDimensions": {
    "dimensions": {"sheetId": 0, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 5}
  }
}
```

**Add a new sheet tab:**
```json
{
  "addSheet": {
    "properties": {"title": "New Tab"}
  }
}
```

**Delete a sheet tab:**
```json
{
  "deleteSheet": {
    "sheetId": 123456789
  }
}
```

### Step 7 — Create a new spreadsheet (if needed)
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "properties": {"title": "My New Spreadsheet"},
    "sheets": [
      {"properties": {"title": "Summary"}},
      {"properties": {"title": "Details"}}
    ]
  }' \
  "https://sheets.googleapis.com/v4/spreadsheets"
```

### Step 8 — Verify
Re-read the affected range to confirm the edits:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://sheets.googleapis.com/v4/spreadsheets/SPREADSHEET_ID/values/Sheet1!A1:Z10"
```

### Step 9 — Report
Tell the user what was changed, with a link to the spreadsheet:
`https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`

### Troubleshooting
- **403 Forbidden:** Token expired — re-run `TOKEN=$(gcloud auth print-access-token)`
- **403 quota project error:** Add header `-H "x-goog-user-project: YOUR_PROJECT_ID"` to all requests
- **404 sheet not found:** Check the exact sheet name (case-sensitive) and URL-encode special characters
- **400 invalid range:** Ensure A1 notation is correct and the sheet name matches exactly
- **Formatting not applied:** Verify you're using the integer `sheetId` (from metadata), not the sheet name
