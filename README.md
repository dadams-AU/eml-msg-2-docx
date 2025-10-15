# EML/MSG → DOCX/PDF (Streamlit)

Upload `.eml` or `.msg` emails. The app extracts HTML/text (and inline `cid:` images) and uses Pandoc to generate DOCX and/or PDF.

## Deploy (Streamlit Community Cloud)
1. Push this repo to GitHub.
2. In Streamlit Community Cloud → "New app" → pick this repo (`app.py` as entrypoint).
3. The build will install `requirements.txt` (Python deps) and `packages.txt` (apt packages).
4. Open the app and upload files.

## Local run
```bash
python -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
# Linux:
sudo apt-get install pandoc wkhtmltopdf
# Arch:
sudo pacman -S pandoc wkhtmltopdf

streamlit run app.py
