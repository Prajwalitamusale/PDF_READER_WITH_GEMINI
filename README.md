# JDFit

India-first resume vs job-description fit check. Upload a PDF/DOCX resume, paste a JD, get a score, skill gaps, an ATS-friendly rewrite (copy or download TXT/DOCX), and interview questions.

Not a real ATS. Gemini estimate only. Files are not stored.

## Streamlit Cloud

This repo is meant to deploy on [Streamlit Community Cloud](https://share.streamlit.io).

Secrets (App settings → Secrets):

```toml
GOOGLE_API_KEY = "your-key"
APP_PASSWORD = "optional-shared-tester-password"
```

If `APP_PASSWORD` is set, testers must enter it before using the app.

Main file: `app.py`

## Local

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export GOOGLE_API_KEY="your-key"
streamlit run app.py
```
