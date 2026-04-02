# hoichoi Film Script Production Breakdown

AI-powered film script breakdown for pre-production planning.

## What it does

- Ingests film scripts (PDF or DOCX) and uses GPT-4 to extract structured production data
- Identifies all unique locations, scene types (INT/EXT), and time-of-day requirements
- Extracts props, costumes, and special equipment per scene
- Supports scanned scripts via Mistral OCR for image-based PDFs
- Exports complete production breakdown to Excel and PDF for department heads

## Tech Stack

Python · Streamlit · OpenAI GPT-4 · Mistral OCR (`mistral-ocr-latest`) · PyPDF2 · pdfplumber · openpyxl · reportlab

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Secrets setup

Create `.streamlit/secrets.toml` (never commit this file):

```toml
OPENAI_API_KEY  = "sk-..."
MISTRAL_API_KEY = "..."

[users]
"yourname@hoichoi.tv" = ""   # paste hash generated below
```

Generate a password hash:
```bash
python -c "from auth import hash_password; print(hash_password('your_password'))"
```

## Sample output

After processing a script the tool produces:

- **Location list** — unique shooting locations with INT/EXT and DAY/NIGHT tags
- **Scene breakdown** — per-scene summary with location, cast, props, and notes
- **Props inventory** — consolidated props list grouped by department
- **Excel export** — one sheet per breakdown category, ready for production scheduling
- **PDF report** — formatted breakdown document for distribution to department heads

## Access

Login requires a `@hoichoi.tv` email and a hashed password stored in `secrets.toml`.
Sessions expire after 8 hours. Accounts are locked for 15 minutes after 5 failed attempts.

---

Built by **Alokananda Sengupta**, Product & Business Analyst @ Hoichoi
