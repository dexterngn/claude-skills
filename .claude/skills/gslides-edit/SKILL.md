---
name: gslides-edit
description: Create or edit a Google Slides presentation using natural language instructions
user-invocable: true
---

Create or edit a Google Slides presentation — either directly via a Google Slides URL or from a local `.pptx` file.

**Usage:**
- `/gslides-edit <google_slides_url> <instruction>` — edit a Google Slides deck directly
- `/gslides-edit create <filename> <description>` — create a new local deck
- `/gslides-edit edit <filepath> <instruction>` — edit an existing local .pptx file

**Examples:**
- `/gslides-edit https://docs.google.com/presentation/d/ABC123/edit "Add a summary slide at the end"`
- `/gslides-edit https://docs.google.com/presentation/d/ABC123/edit "Review this deck and suggest improvements"`
- `/gslides-edit create q1_review.pptx "Q1 business review with title slide, metrics slide, and next steps"`
- `/gslides-edit edit ~/Downloads/deck.pptx "Add a slide at the end with our top 3 risks"`

**Prerequisites:**
```bash
pip install python-pptx lxml
```

---

## How to execute

### Step 0 — Detect mode

Parse the user's input to determine the mode:

1. **Google Slides URL mode** — input contains `docs.google.com/presentation/d/`
   - Extract the presentation ID from the URL (the part between `/d/` and the next `/`)
   - Proceed to Step 1A
2. **Local create mode** — input starts with `create`
   - Proceed to Step 1B (skip to Step 3)
3. **Local edit mode** — input starts with `edit` or references a local `.pptx` path
   - Proceed to Step 1B

### Step 1A — Google Slides: Download via Drive API

Use `gcloud auth print-access-token` for authentication. Write and run a Python script:

```python
import subprocess, urllib.request, json

PRES_ID = "EXTRACTED_PRESENTATION_ID"
LOCAL_PPTX = "/tmp/gslides_edit.pptx"

def get_token() -> str:
    return subprocess.run(
        ["gcloud", "auth", "print-access-token"],
        capture_output=True, text=True, check=True,
    ).stdout.strip()

token = get_token()

# Download as PPTX via Drive export API
url = (
    f"https://www.googleapis.com/drive/v3/files/{PRES_ID}/export"
    "?mimeType=application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
with urllib.request.urlopen(req, timeout=60) as r:
    data = r.read()
with open(LOCAL_PPTX, "wb") as f:
    f.write(data)
print(f"Downloaded {len(data):,} bytes to {LOCAL_PPTX}")
```

Then proceed to Step 2 using `LOCAL_PPTX` as the file path.

### Step 1B — Local file: use the path directly

No download needed. Use the provided file path.

### Step 2 — Read the existing deck structure

Run a Python snippet to inspect the deck:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = Presentation("PATH_TO_FILE")
print(f"Slide dimensions: {prs.slide_width} x {prs.slide_height} EMU")
print(f"  = {prs.slide_width / 914400:.2f} x {prs.slide_height / 914400:.2f} inches")
print(f"Total slides: {len(prs.slides)}\n")

for i, slide in enumerate(prs.slides):
    print(f"--- Slide {i}: layout='{slide.slide_layout.name}' ---")
    for shape in slide.shapes:
        if shape.has_table:
            tbl = shape.table
            print(f"  [TABLE] {shape.name}: {tbl.rows.__len__()} rows x {len(tbl.columns)} cols")
            headers = [tbl.cell(0, c).text.strip() for c in range(len(tbl.columns))]
            print(f"    Headers: {headers}")
        elif shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = "".join(r.text for r in para.runs).strip()
                if text:
                    print(f"  [{shape.shape_type}] {shape.name}: {text[:120]}")
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            print(f"  [PICTURE] {shape.name}: {shape.width}x{shape.height} at ({shape.left},{shape.top})")
        elif shape.has_chart:
            print(f"  [CHART] {shape.name}: {shape.chart.chart_type}")
```

Review the output to understand the deck structure before making changes.

### Step 3 — Write the Python script to create/modify the deck

Use `python-pptx` to implement the changes. Key patterns:

**Create a new presentation:**
```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)
```

**Add a title slide:**
```python
slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide layout
slide.shapes.title.text = "My Title"
slide.placeholders[1].text = "Subtitle text"
```

**Add a content slide with bullet points:**
```python
slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content
slide.shapes.title.text = "Slide Title"
tf = slide.placeholders[1].text_frame
tf.text = "First bullet"
p = tf.add_paragraph()
p.text = "Second bullet"
p.level = 1  # indent level (0=top, 1=sub-bullet)
```

**Add a blank slide with custom text box:**
```python
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
tf = txBox.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Custom text"
p.font.size = Pt(24)
p.font.bold = True
p.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
```

**Style a text run:**
```python
run = paragraph.runs[0]
run.font.name  = "Calibri"
run.font.size  = Pt(18)
run.font.bold  = True
run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
```

**Add a table:**
```python
rows, cols = 3, 4
table = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(10), Inches(3)).table
table.cell(0, 0).text = "Header"
```

**Find & replace text across all slides:**
```python
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "old text" in run.text:
                        run.text = run.text.replace("old text", "new text")
```

**Set slide background color:**
```python
fill = slide.background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(0x1F, 0x49, 0x7D)
```

**Copy a slide (deep copy):**
```python
from copy import deepcopy

def copy_slide(prs, source_slide):
    """Deep-copy a slide and append it to the presentation."""
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    for ph in list(new_slide.placeholders):
        sp = ph._element
        sp.getparent().remove(sp)
    for shape in source_slide.shapes:
        el = deepcopy(shape._element)
        new_slide.shapes._spTree.append(el)
    return new_slide
```

**Move a slide to a specific position:**
```python
def move_slide_to(prs, from_idx, to_idx):
    slides = prs.slides._sldIdLst
    el = slides[from_idx]
    slides.remove(el)
    slides.insert(to_idx, el)
```

**Delete a slide:**
```python
xml_slides = prs.slides._sldIdLst
xml_slides.remove(xml_slides[slide_index])
```

**Save the file:**
```python
prs.save("/tmp/deck.pptx")
```

### Step 4 — Execute the script
Write the full Python script to a temp file and run it:
```bash
python3 /tmp/build_deck.py
```
Fix any errors and re-run until it succeeds.

### Step 5 — For Google Slides mode: Upload back via Drive API

After saving the modified PPTX locally, **strip comments** to prevent duplication on re-upload, then upload:

```python
import subprocess, urllib.request, zipfile, os, json, re

PRES_ID = "EXTRACTED_PRESENTATION_ID"
LOCAL_PPTX = "/tmp/gslides_edit.pptx"
CLEAN_PPTX = "/tmp/gslides_edit_clean.pptx"

def get_token() -> str:
    return subprocess.run(
        ["gcloud", "auth", "print-access-token"],
        capture_output=True, text=True, check=True,
    ).stdout.strip()

def strip_comments(src, dst):
    """Remove embedded comment XML to prevent duplication on re-upload."""
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/comments/"):
                continue
            if item.filename.endswith(".rels") and b"comment" in data.lower():
                data = re.sub(
                    rb'<Relationship[^>]*Target="[^"]*comment[^"]*"[^>]*/>\s*',
                    b"", data, flags=re.IGNORECASE
                )
            zout.writestr(item, data)

strip_comments(LOCAL_PPTX, CLEAN_PPTX)

token = get_token()
upload_url = f"https://www.googleapis.com/upload/drive/v3/files/{PRES_ID}?uploadType=media"
with open(CLEAN_PPTX, "rb") as f:
    file_data = f.read()

req = urllib.request.Request(
    upload_url,
    data=file_data,
    method="PATCH",
    headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    },
)
with urllib.request.urlopen(req, timeout=120) as r:
    resp = json.loads(r.read())
    print(f"Uploaded successfully: {resp.get('name', PRES_ID)}")

os.remove(CLEAN_PPTX)
```

For new local files, drag the `.pptx` into [drive.google.com](https://drive.google.com) → right-click → **Open with Google Slides**.

### Step 6 — Report
Tell the user:
- What slides were read/created/modified
- A summary of changes made
- For Google Slides mode: confirm the upload was successful and provide the deck URL: `https://docs.google.com/presentation/d/PRESENTATION_ID/edit`
- For local mode: the output file path

### Troubleshooting
- **ModuleNotFoundError: python-pptx:** Run `pip install python-pptx lxml`
- **Comments duplicated after re-upload:** You skipped `strip_comments()` — always strip before uploading
- **Token expired:** Re-run `gcloud auth print-access-token`
- **Upload fails:** Check that the file is valid PPTX (`python3 -c "from pptx import Presentation; Presentation('/tmp/deck.pptx')"`)
- **Formatting lost on upload to Google Slides:** Some advanced PPTX features (animations, custom fonts) don't convert perfectly — keep formatting simple
- **`AttributeError: no .rgb property on color type '_SchemeColor'`:** Wrap `font.color.rgb` access in try/except — scheme colors don't expose RGB directly
