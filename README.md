# Report Studio — v1.0 preview

Local-first formatting workspace for academic and business reports. No account or database required.

## Run on Windows

```powershell
git clone https://github.com/tranmanhhung522006/ReportStudio.git
cd ReportStudio
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe server.py
```

Open http://127.0.0.1:8765. On macOS/Linux use .venv/bin/python instead.

## Workflow

1. Select a DOCX report. Optionally paste an outline with numbered headings.
2. Import requirement PDF/DOCX/TXT or paste text. Extract rules and review evidence before applying.
3. Adjust formatting settings, analyze report, and review heading levels.
4. Run quality audit on original; export a formatted DOCX or render PDF preview.
5. Open output in Word and review fields, page numbers, references and layout. Update TOC/caption fields selectively (F9), not citation fields.

## PDF preview

Install LibreOffice locally. The app invokes soffice using an isolated temporary profile. LibreOffice layout can differ from Word; PDF preview is not a Word-native renderer. Scanned requirement PDFs need OCR externally.

## Optional AI

Set OPENAI_API_KEY in your shell, optionally OPENAI_MODEL, before starting. AI sends requirement text only after explicit consent. API billing is separate from ChatGPT Plus. Never commit credentials. Local parsing is default.

## Safety and scope

Original DOCX is never overwritten. Existing complex fields, hyperlinks and objects are protected conservatively; integrity checks compare selected XML objects and embedded assets before export. This is not a certification of EndNote compatibility or all document contents. Test on a backup. Existing captions with fields are preserved; the app renumbers recognizable plain-text captions, does not invent captions, and does not update cross-references automatically. APA/Harvard/IEEE reference section audit is advisory, not automatic reference conversion. Quality score is based on supported checks, not an academic grade. Page numbering currently targets primary footers. Existing complex numbering and sections require review. AI extraction and OCR are not required for basic use.

## Tests

```bash
python -m unittest discover -s tests -v
```

GitHub Actions runs the regression suite on pushes and pull requests. Do not expose the local HTTP server to the internet.
