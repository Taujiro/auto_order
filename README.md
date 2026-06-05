# Pedidos Email Auto

A local Python and Streamlit application for automating supplier order emails, pending-balance requests, and quote checking workflows.

## What This App Does

The project brings together three workflows that used to require repetitive manual work:

- sending supplier orders by Gmail with an Excel attachment and signature;
- sending batch pending-balance request emails without attachments;
- checking received quote PDFs against the sent order quantity and the supplier reference price sheet.

It was designed to run locally and fit into an existing purchasing workflow based on spreadsheets, PDF quotes, supplier folders, and Gmail. The goal is to reduce repeated work, standardize messages, keep a useful history of sent emails, and make supplier quote analysis faster.

## Why I Built It

This application came from a real operational need. I was already creating small local tools to solve practical problems in my daily work, but I noticed a risk: if everything stayed only on my computer, I could lose important automations if the machine failed.

I decided to organize this project on GitHub both as a backup and as a professional portfolio piece. It shows how I use automation, Python, spreadsheets, and AI-assisted development to turn manual processes into usable internal tools.

## Technologies

- Python
- Streamlit
- Pandas
- OpenPyXL
- pypdf
- Gmail API / OAuth
- requests
- unittest

## Main Features

- Supplier registration through `fornecedores.xlsx`.
- Order email generation with Excel attachment.
- Automatic copy of sent orders to the supplier folder.
- Separate history logs for orders and pending-balance requests.
- Email normalization for accidental spaces in addresses.
- OAuth reauthentication when the Gmail token expires or is revoked.
- Batch pending-balance emails without attachments.
- Quote checking from PDF against sent order and supplier price reference.
- Excel export for the quote checking report.

## Project Structure

```text
.
├── app.py                    # Streamlit app and core business rules
├── test_app.py               # Automated tests
├── requirements.txt          # Python dependencies
├── rodar_app.bat             # Local Windows shortcut to start the app
├── fornecedores.example.csv  # Safe example of the supplier sheet structure
├── .env.example              # Local configuration example
└── README.md
```

Local files and folders used during real operation are intentionally not versioned:

- `credentials.json`
- `token.json`
- `fornecedores.xlsx`
- `pedidos/`
- `enviados/`
- `logs/`
- `assinatura.png`
- `referencias_precos/`

## Installation

Create and activate a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```powershell
pip install -r requirements.txt
```

Set up Gmail OAuth:

1. Create a project in Google Cloud.
2. Enable the Gmail API.
3. Create an OAuth credential for a desktop application.
4. Download the credential file and save it as `credentials.json` in the project root.

Create `fornecedores.xlsx` with these columns:

- `fornecedor`
- `email`
- `cc`
- `pasta`

Use `fornecedores.example.csv` as a safe visual reference for the expected structure.

## Running The App

With the virtual environment active:

```powershell
streamlit run app.py
```

You can also use the local Windows shortcut:

```powershell
.\rodar_app.bat
```

On the first email send, the app opens Google's authentication flow and creates `token.json` locally.

## Price Reference Configuration

The quote checking feature uses supplier reference price workbooks. By default, the app looks for them in:

```text
referencias_precos/
```

You can point to another folder with an environment variable:

```powershell
$env:PRICE_REFERENCE_DIR="C:\path\to\price_references"
```

The currently expected file names are:

- `Autobras.xlsx`
- `Tuba.xlsx`
- `Tsa.xlsx`

## Tests

Run:

```powershell
python -m unittest test_app.py
python -m py_compile app.py test_app.py
```

Some PDF parser tests use real supplier quote PDFs. In public environments or machines without those files, they are skipped automatically. To run them locally, define:

```powershell
$env:AUTOBRAS_QUOTE_PDF="C:\path\to\autobras.pdf"
$env:TUBA_QUOTE_PDF="C:\path\to\tuba.pdf"
$env:TSA_QUOTE_PDF="C:\path\to\tsa.pdf"
```

## Next Steps

- Add a settings screen for local paths.
- Support more supplier PDF layouts.
- Improve the exported quote checking reports.
- Split `app.py` into smaller modules as the project grows.
- Add synthetic PDF samples for complete public test coverage.

## AI-Assisted Development Note

This project was developed with support from AI/Codex as a development, review, and acceleration tool. The problem definition, workflow decisions, practical validation, and project direction were led by me, while AI helped turn a real business need into a working local application.
