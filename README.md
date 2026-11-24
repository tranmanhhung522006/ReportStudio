# Report Studio

Local-first DOCX formatting workspace. Automatically detect headings, review hierarchy, apply formatting and export a new Word document.

## Run locally

Python 3.10+ required.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python server.py
```

Open http://127.0.0.1:8765 . No AI API key required. Use a copy of your document and verify the result in Word. This initial version does not perform full Word rendering, PDF export or EndNote refresh.
