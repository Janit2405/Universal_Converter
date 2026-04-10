# Universal Excel Converter

Drop any file → auto-detect format → download as Excel (.xlsx).

## Supported Formats

| Format | Detection |
|--------|-----------|
| EDI 834 | `ISA*` header, `~` terminators |
| FLX Fixed-Width | Filename contains "FLX", fixed column positions |
| CSV | Comma-separated values |
| TSV | Tab-separated values |
| Pipe-Delimited | `\|` separated fields |
| Generic Fixed-Width | Consistent line lengths, auto-detected column boundaries |
| GZIP | `.gz` extension — decompresses transparently |

## Run Locally

```bash
pip install -r requirements.txt
python app.py
```

Open [http://localhost:5000](http://localhost:5000).

## Deploy on Render

1. Push this repo to GitHub
2. Go to [render.com](https://render.com) → New → Web Service
3. Connect your GitHub repo
4. Settings:
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn --bind 0.0.0.0:$PORT app:app`
5. Deploy

## Output

- Professional Excel formatting with headers, filters, freeze panes
- Auto-width columns, alternate row shading
- Summary sheet for EDI 834 and FLX files with record counts and financial totals

---

Quantify Specialty Care — Data Team
