---
name: gdocs-edit
description: Edit a Google Doc using natural language instructions
user-invocable: true
---

Edit a Google Doc using the Google Docs REST API authenticated via `gcloud auth print-access-token`.

**Usage:** `/gdocs-edit <document_id> <instruction>`

**Example:** `/gdocs-edit 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms "Add a summary section at the end"`

---

## How to execute

### Step 1 — Get auth token
Run:
```bash
TOKEN=$(gcloud auth print-access-token)
```
Store this token. It expires after ~1 hour.

### Step 2 — Read the document
Documents may have multiple tabs. Always use `includeTabsContent=true` to get all content:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://docs.googleapis.com/v1/documents/DOCUMENT_ID?includeTabsContent=true"
```
Parse the response to understand the document structure:
- `tabs[]` — array of document tabs, each with `tabProperties.tabId` and `tabProperties.title`
- `tabs[].documentTab.body.content` is an array of structural elements
- Each element has `startIndex` and `endIndex`
- Paragraphs have `paragraph.elements[].textRun.content` for text
- Paragraph styles: `paragraph.paragraphStyle.namedStyleType` (TITLE, HEADING_1, HEADING_2, etc.)
- Indices are 1-based character positions

Use a Python one-liner to extract text quickly:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://docs.googleapis.com/v1/documents/DOCUMENT_ID?includeTabsContent=true" \
  | python3 -c "
import json, sys
doc = json.loads(sys.stdin.read())
for tab in doc.get('tabs', []):
    props = tab.get('tabProperties', {})
    print(f'Tab: {props.get(\"title\",\"\")} (id: {props.get(\"tabId\",\"\")})')
    body = tab.get('documentTab', {}).get('body', {}).get('content', [])
    for elem in body:
        para = elem.get('paragraph', {})
        style = para.get('paragraphStyle', {}).get('namedStyleType', '')
        texts = []
        for e in para.get('elements', []):
            t = e.get('textRun', {}).get('content', '').strip()
            if t: texts.append(t)
        if texts:
            print(f'  [{style}] {\" \".join(texts)[:120]}')
"
```

### Step 3 — Plan the edits
Based on the instruction and document content, decide which API operations to use:

| Goal | API request type |
|------|-----------------|
| Find & replace text | `replaceAllText` |
| Insert text at index | `insertText` |
| Delete a range | `deleteContentRange` |
| Bold / italic / underline | `updateTextStyle` |
| Change paragraph style (Heading, Normal) | `updateParagraphStyle` |
| Append to end | `insertText` at `endIndex - 1` of last element |
| Insert table | `insertTable` |
| Set cell background color | `updateTableCellStyle` |
| Set column widths | `updateTableColumnProperties` |
| Rename document | Use Drive API `files.update` |

**Important rules:**
- When making multiple index-based edits, execute them **bottom-to-top** (highest index first) to avoid index shifting
- Prefer `replaceAllText` for updating dates, names, phrases — it's index-safe
- Each paragraph's `endIndex` includes the trailing `\n`
- For documents with tabs, include `"tabId": "TAB_ID"` in location objects (e.g., `{"index": 42, "tabId": "t.0"}`)
- When inserting large blocks: insert the text first in one call, then apply formatting in a second call (indices shift after insertion)

### Step 4 — Apply edits via batchUpdate
```bash
curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"requests": [...]}' \
  "https://docs.googleapis.com/v1/documents/DOCUMENT_ID:batchUpdate"
```

Example requests:

**replaceAllText:**
```json
{
  "replaceAllText": {
    "containsText": {"text": "old text", "matchCase": false},
    "replaceText": "new text"
  }
}
```

**insertText:**
```json
{
  "insertText": {
    "location": {"index": 42, "tabId": "t.0"},
    "text": "Hello world\n"
  }
}
```

**deleteContentRange:**
```json
{
  "deleteContentRange": {
    "range": {"startIndex": 10, "endIndex": 25, "tabId": "t.0"}
  }
}
```

**updateTextStyle (bold):**
```json
{
  "updateTextStyle": {
    "range": {"startIndex": 5, "endIndex": 20, "tabId": "t.0"},
    "textStyle": {"bold": true},
    "fields": "bold"
  }
}
```

**updateParagraphStyle:**
```json
{
  "updateParagraphStyle": {
    "range": {"startIndex": 5, "endIndex": 20, "tabId": "t.0"},
    "paragraphStyle": {"namedStyleType": "HEADING_2"},
    "fields": "namedStyleType"
  }
}
```

**insertTable:**
```json
{
  "insertTable": {
    "location": {"index": 42, "tabId": "t.0"},
    "rows": 5,
    "columns": 4
  }
}
```

**updateTableCellStyle (background color):**
```json
{
  "updateTableCellStyle": {
    "tableRange": {
      "tableCellLocation": {"tableStartLocation": {"index": 43, "tabId": "t.0"}, "rowIndex": 0, "columnIndex": 0},
      "rowSpan": 1, "columnSpan": 4
    },
    "tableCellStyle": {"backgroundColor": {"color": {"rgbColor": {"red": 0.9, "green": 0.9, "blue": 0.9}}}},
    "fields": "backgroundColor"
  }
}
```

### Step 5 — Verify
Re-read the document to confirm the edits were applied correctly:
```bash
curl -s -H "Authorization: Bearer $TOKEN" \
  "https://docs.googleapis.com/v1/documents/DOCUMENT_ID?includeTabsContent=true"
```

### Step 6 — Report
Tell the user what was changed, with a link to the document:
`https://docs.google.com/document/d/DOCUMENT_ID/edit`

### Troubleshooting
- **403 Forbidden:** Token expired — re-run `gcloud auth print-access-token`
- **Tab content is empty:** You forgot `?includeTabsContent=true` on the read call
- **Indices are wrong after edit:** You applied edits top-to-bottom instead of bottom-to-top
- **createTab not supported:** Google Docs API does not support creating tabs programmatically yet — add content as a new section in an existing tab instead
