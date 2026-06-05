# Auto Order
![Python](https://img.shields.io/badge/Python-3.13-blue?logo=python)
![Streamlit](https://img.shields.io/badge/Streamlit-red?logo=streamlit)
![Pandas](https://img.shields.io/badge/Pandas-purple?logo=pandas)
![Gmail API](https://img.shields.io/badge/Gmail%20API-orange?logo=gmail)
![Local App](https://img.shields.io/badge/app-local-blue)
![Status](https://img.shields.io/badge/status-active-brightgreen)
![AI Assisted](https://img.shields.io/badge/AI-assisted-purple)

A local Python and Streamlit application for automating supplier order emails, pending-balance requests, and supplier quote checking workflows.

## Overview

Pedidos Email Auto was built to support a real purchasing workflow based on spreadsheets, supplier folders, PDF quotes, and Gmail.

The application brings together three repetitive processes:

* sending supplier order emails through Gmail with an Excel attachment and signature;
* sending batch pending-balance request emails without attachments;
* checking received supplier quote PDFs against the sent order quantities and supplier reference price sheets.

The goal is to reduce manual work, standardize supplier communication, keep a clear history of sent emails, and make quote analysis faster and more reliable.

## Why I Built This

This project came from a practical operational need.

I was creating small local tools to solve repetitive tasks in my daily work, but I realized that keeping these automations only on my computer created a risk: if the machine failed or was replaced, the project could be lost or become difficult to recover.

I decided to organize this application on GitHub both as a backup strategy and as a portfolio project. It represents how I use Python, automation, spreadsheets, and AI-assisted development to transform manual business processes into usable internal tools.

## Technologies

* Python
* Streamlit
* Pandas
* OpenPyXL
* pypdf
* Gmail API / OAuth
* requests
* unittest

## Main Features

* Supplier registration through `fornecedores.xlsx`.
* Order email generation with Excel attachment.
* Automatic copy of sent order files to the supplier folder.
* Separate history logs for sent orders and pending-balance requests.
* Email normalization to handle accidental spaces in addresses.
* OAuth reauthentication when the Gmail token expires or is revoked.
* Batch pending-balance request emails without attachments.
* Quote checking from PDF against sent order quantities and supplier reference prices.
* Excel export for quote checking reports.

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

The following files and folders are used during real local operation and are intentionally not versioned:

* `credentials.json`
* `token.json`
* `fornecedores.xlsx`
* `pedidos/`
* `enviados/`
* `logs/`
* `assinatura.png`
* `referencias_precos/`

## Security and Privacy

This project is designed to run locally.

Sensitive files such as Gmail credentials, OAuth tokens, real supplier spreadsheets, sent order files, logs, signatures, and price reference files are not included in the repository.

Example files are provided only to document the expected structure without exposing real business data.

## Installation

Create and activate a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install the dependencies:

```powershell
pip install -r requirements.txt
```

Set up Gmail OAuth:

1. Create a project in Google Cloud.
2. Enable the Gmail API.
3. Create an OAuth credential for a desktop application.
4. Download the credential file and save it as `credentials.json` in the project root.

Create a `fornecedores.xlsx` file with these columns:

* `fornecedor`
* `email`
* `cc`
* `pasta`

Use `fornecedores.example.csv` as a safe reference for the expected structure.

## Running the App

With the virtual environment active, run:

```powershell
streamlit run app.py
```

You can also use the local Windows shortcut:

```powershell
.\rodar_app.bat
```

On the first email send, the app opens Google's authentication flow and creates `token.json` locally.

## Price Reference Configuration

The quote checking feature uses supplier reference price workbooks.

By default, the app looks for them in:

```text
referencias_precos/
```

You can also define a custom folder using an environment variable:

```powershell
$env:PRICE_REFERENCE_DIR="C:\path\to\price_references"
```

For new suppliers, the reference workbook should use the supplier name from `fornecedores.xlsx`.
For example:

* `GONEL.xlsx`
* `WISA.xlsx`

Some legacy suppliers can also use the explicit names already mapped by the app, such as `Autobras.xlsx`, `Tuba.xlsx`, and `Tsa.xlsx`.

## Tests

Run the automated tests with:

```powershell
python -m unittest test_app.py
python -m py_compile app.py test_app.py
```

Some PDF parser tests use real supplier quote PDFs. In public environments or machines without those files, these tests are skipped automatically.

To run them locally, define the following environment variables:

```powershell
$env:AUTOBRAS_QUOTE_PDF="C:\path\to\autobras.pdf"
$env:TUBA_QUOTE_PDF="C:\path\to\tuba.pdf"
$env:TSA_QUOTE_PDF="C:\path\to\tsa.pdf"
```

## Next Steps

* Add a settings screen for local paths.
* Support more supplier PDF layouts.
* Improve the exported quote checking reports.
* Split `app.py` into smaller modules as the project grows.
* Add synthetic PDF samples for complete public test coverage.

## AI-Assisted Development

This project was developed with support from AI/Codex as a development, review, and acceleration tool.

The problem definition, workflow decisions, practical validation, and project direction were led by me. AI was used to help structure, implement, review, and improve the application while turning a real business need into a working local tool.
