from __future__ import annotations

import base64
import csv
import os
import re
import shutil
import unicodedata
import ssl
from dataclasses import dataclass
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import make_msgid
from io import BytesIO
from pathlib import Path

import certifi
import pandas as pd
import requests
import streamlit as st
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import AuthorizedSession, Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from pypdf import PdfReader
from requests.adapters import HTTPAdapter


BASE_DIR = Path(__file__).resolve().parent
PEDIDOS_DIR = BASE_DIR / "pedidos"
ENVIADOS_DIR = BASE_DIR / "enviados"
LOGS_DIR = BASE_DIR / "logs"
FORNECEDORES_FILE = BASE_DIR / "fornecedores.xlsx"
LOG_FILE = LOGS_DIR / "envios.csv"
PENDENCIA_LOG_FILE = LOGS_DIR / "pendencias.csv"
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"
SIGNATURE_IMAGE_FILE = BASE_DIR / "assinatura.png"

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
EMAIL_SUBJECT_TEMPLATE = "Pedido de Cotacao - {fornecedor}"
EMAIL_BODY_TEXT = "Bom dia,\n\nSegue em anexo o pedido de {fornecedor}.\n\nAtenciosamente,"
PENDENCIA_SUBJECT_TEMPLATE = "Pendência {fornecedor}"
PENDENCIA_BODY_TEXT = "Bom dia,\n\nSolicito saldo pendente de {fornecedor}, por favor.\n\nAtenciosamente,"
SIGNATURE_NAME = "Gabriel Taujiro"
SIGNATURE_EMAIL = "contato.estoqueltda@gmail.com"
SIGNATURE_PHONE = "(62) 3771-3801"
SIGNATURE_IMAGE_WIDTH = 400
SIGNATURE_IMAGE_HEIGHT = 200
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
CA_BUNDLE_FILE = LOGS_DIR / "certifi_windows_ca.pem"
GMAIL_SEND_URL = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"
PEDIDO_LOG_COLUMNS = [
    "data",
    "fornecedor",
    "arquivo",
    "email",
    "cc",
    "data_envio",
    "data_pedido",
    "tipo",
    "assunto",
    "pasta",
    "total_destinatarios",
    "observacao",
]
PENDENCIA_LOG_COLUMNS = [
    "data",
    "fornecedor",
    "email",
    "cc",
    "assunto",
    "data_envio",
    "data_referencia",
    "tipo",
    "total_destinatarios",
    "observacao",
]
CONFERENCE_REPORT_COLUMNS = [
    "codigo",
    "quantidade_pedida",
    "quantidade_cotada",
    "preco_cotado",
    "preco_referencia",
    "diferenca_percentual",
    "status",
]
DEFAULT_REFERENCE_DIR = BASE_DIR / "referencias_precos"
REFERENCE_FILE_NAMES = {
    "AUTOBRAS": "Autobras.xlsx",
    "TUBA": "Tuba.xlsx",
    "TSA": "Tsa.xlsx",
}
PRICE_TOLERANCE_PERCENT = 1.0

def build_combined_ca_bundle() -> Path:
    ca_chunks = [Path(certifi.where()).read_text(encoding="ascii")]

    if hasattr(ssl, "enum_certificates"):
        for store_name in ("ROOT", "CA"):
            for certificate, encoding, _trust in ssl.enum_certificates(store_name):
                if encoding == "x509_asn":
                    ca_chunks.append(ssl.DER_cert_to_PEM_cert(certificate))

    unique_chunks = list(dict.fromkeys(ca_chunks))
    LOGS_DIR.mkdir(exist_ok=True)
    bundle_text = "\n".join(chunk.strip() for chunk in unique_chunks if chunk.strip()) + "\n"
    if not CA_BUNDLE_FILE.exists() or CA_BUNDLE_FILE.read_text(encoding="ascii") != bundle_text:
        CA_BUNDLE_FILE.write_text(bundle_text, encoding="ascii")
    return CA_BUNDLE_FILE


def configure_ssl_certificates() -> Path:
    try:
        ca_bundle_path = build_combined_ca_bundle()
    except Exception:
        ca_bundle_path = Path(certifi.where())

    os.environ.pop("SSLKEYLOGFILE", None)
    os.environ["SSL_CERT_FILE"] = str(ca_bundle_path)
    os.environ["REQUESTS_CA_BUNDLE"] = str(ca_bundle_path)
    return ca_bundle_path


CA_BUNDLE_PATH = configure_ssl_certificates()


def create_ssl_context() -> ssl.SSLContext:
    context = ssl.create_default_context(cafile=str(CA_BUNDLE_PATH))
    if hasattr(ssl, "VERIFY_X509_STRICT"):
        context.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return context


class GoogleSSLAdapter(HTTPAdapter):
    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        pool_kwargs["ssl_context"] = create_ssl_context()
        return super().init_poolmanager(connections, maxsize, block=block, **pool_kwargs)

    def proxy_manager_for(self, proxy, **proxy_kwargs):
        proxy_kwargs["ssl_context"] = create_ssl_context()
        return super().proxy_manager_for(proxy, **proxy_kwargs)


def configure_google_session(session: requests.Session) -> requests.Session:
    session.verify = str(CA_BUNDLE_PATH)
    session.mount("https://", GoogleSSLAdapter())
    return session


@dataclass(frozen=True)
class Pedido:
    fornecedor: str
    arquivo: Path
    email: str | None
    cc: str = ""
    pasta: str = ""


@dataclass(frozen=True)
class PendenciaEmail:
    fornecedor: str
    email: str
    cc: str = ""


def ensure_project_structure() -> None:
    PEDIDOS_DIR.mkdir(exist_ok=True)
    ENVIADOS_DIR.mkdir(exist_ok=True)
    LOGS_DIR.mkdir(exist_ok=True)


def normalize_text(value: object) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_product_code(value: object) -> str:
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    digits = re.sub(r"\D+", "", text)
    if digits:
        return digits
    return re.sub(r"[^A-Z0-9]+", "", text)


def parse_brazilian_number(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def read_pdf_text(pdf_source: Path | BytesIO | object) -> str:
    reader = PdfReader(pdf_source)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def extract_money_values(text: str) -> list[float]:
    return [
        parsed
        for parsed in (parse_brazilian_number(match) for match in re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", text))
        if parsed is not None
    ]


def parse_solida_quote_text(text: str) -> pd.DataFrame:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    item_starts: list[tuple[int, str, list[str]]] = []

    for line in lines:
        match = re.match(r"^(\d{1,3})\s+(\d{3,8})(?!,\d)(.*)$", line)
        if not match:
            continue
        item_number = int(match.group(1))
        code = match.group(2)
        if item_number > 500:
            continue
        item_starts.append((item_number, code, [line]))

    if not item_starts:
        return quote_items_dataframe([])

    blocks: list[tuple[str, str]] = []
    current_index = -1
    start_by_line = {id(block_lines[0]): index for index, (_item, _code, block_lines) in enumerate(item_starts)}

    item_iter = iter(item_starts)
    current = next(item_iter, None)
    next_item = next(item_iter, None)
    current_lines: list[str] = []
    current_code = ""
    for line in lines:
        if current and line == current[2][0]:
            if current_lines:
                blocks.append((current_code, " ".join(current_lines)))
            current_code = current[1]
            current_lines = [line]
            current = next_item
            next_item = next(item_iter, None)
        elif current_lines:
            current_lines.append(line)
    if current_lines:
        blocks.append((current_code, " ".join(current_lines)))

    rows = []
    for code, block in blocks:
        quantity = None
        unit_price = None
        for qty_text, unit_text, _total_text in re.findall(
            r"\b(\d+)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s+(\d{1,3}(?:\.\d{3})*,\d{2})\b",
            block,
        ):
            quantity = int(qty_text)
            unit_price = parse_brazilian_number(unit_text)
        if quantity is None or unit_price is None:
            rows.append(
                {
                    "codigo": code,
                    "codigo_normalizado": normalize_product_code(code),
                    "quantidade_cotada": pd.NA,
                    "preco_cotado": pd.NA,
                    "status_extracao": "ERRO_EXTRACAO",
                }
            )
            continue
        rows.append(
            {
                "codigo": code,
                "codigo_normalizado": normalize_product_code(code),
                "quantidade_cotada": quantity,
                "preco_cotado": round(float(unit_price), 4),
                "status_extracao": "",
            }
        )
    return quote_items_dataframe(rows)


def parse_tsa_quote_text(text: str) -> pd.DataFrame:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    blocks: list[tuple[str, int, list[str]]] = []
    current_code = ""
    current_quantity = 0
    current_lines: list[str] = []

    for line in lines:
        match = re.match(r"^(\d{1,3})\s+(\d+)\s+([A-Z]-?\d[\w.\-]*)\b(.*)$", line, flags=re.IGNORECASE)
        if match:
            if current_lines:
                blocks.append((current_code, current_quantity, current_lines))
            current_quantity = int(match.group(2))
            current_code = match.group(3).upper()
            current_lines = [line]
        elif current_lines:
            current_lines.append(line)
    if current_lines:
        blocks.append((current_code, current_quantity, current_lines))

    rows = []
    for code, quantity, block_lines in blocks:
        block = " ".join(block_lines)
        prices = extract_money_values(block)
        unit_price = prices[-2] if len(prices) >= 2 else None
        if unit_price is None:
            rows.append(
                {
                    "codigo": code,
                    "codigo_normalizado": normalize_product_code(code),
                    "quantidade_cotada": quantity,
                    "preco_cotado": pd.NA,
                    "status_extracao": "ERRO_EXTRACAO",
                }
            )
            continue
        rows.append(
            {
                "codigo": code,
                "codigo_normalizado": normalize_product_code(code),
                "quantidade_cotada": quantity,
                "preco_cotado": round(float(unit_price), 4),
                "status_extracao": "",
            }
        )
    return quote_items_dataframe(rows)


def quote_items_dataframe(rows: list[dict[str, object]]) -> pd.DataFrame:
    columns = ["codigo", "codigo_normalizado", "quantidade_cotada", "preco_cotado", "status_extracao"]
    return pd.DataFrame(rows, columns=columns)


def extract_quote_items_from_pdf(pdf_source: Path | BytesIO | object, fornecedor: str = "") -> pd.DataFrame:
    text = read_pdf_text(pdf_source)
    fornecedor_normalizado = normalize_text(fornecedor)
    if "item qtde codigo" in normalize_text(text) or fornecedor_normalizado == "tsa":
        return parse_tsa_quote_text(text)
    return parse_solida_quote_text(text)


def find_column_by_normalized_name(df: pd.DataFrame, candidates: set[str], contains: str | None = None) -> str:
    for column in df.columns:
        normalized = normalize_text(column)
        if normalized in candidates or (contains and contains in normalized):
            return column
    raise ValueError(f"Coluna nao encontrada. Procurado: {', '.join(sorted(candidates))}")


def read_sent_order_items(order_path: Path) -> pd.DataFrame:
    df = pd.read_excel(order_path, engine="openpyxl")
    code_column = find_column_by_normalized_name(df, {"catalogo", "catálogo"}, contains="catalog")
    quantity_column = find_column_by_normalized_name(df, {"qnt", "quantidade", "qtde"}, contains=None)
    rows = []
    for _, row in df.iterrows():
        code = str(row.get(code_column, "") or "").strip()
        quantity = parse_brazilian_number(row.get(quantity_column))
        if not code or quantity is None:
            continue
        rows.append(
            {
                "codigo": code,
                "codigo_normalizado": normalize_product_code(code),
                "quantidade_pedida": int(quantity) if float(quantity).is_integer() else quantity,
            }
        )
    return pd.DataFrame(rows, columns=["codigo", "codigo_normalizado", "quantidade_pedida"])


def find_reference_base_folder() -> Path:
    configured_path = os.environ.get("PRICE_REFERENCE_DIR", "").strip()
    if configured_path:
        return Path(os.path.expandvars(os.path.expanduser(configured_path)))
    return DEFAULT_REFERENCE_DIR


def find_reference_file(fornecedor: str) -> Path | None:
    file_name = REFERENCE_FILE_NAMES.get(str(fornecedor or "").strip().upper())
    if not file_name:
        return None
    return find_reference_base_folder() / file_name


def choose_reference_sheet(reference_path: Path) -> str | int:
    with pd.ExcelFile(reference_path, engine="openpyxl") as workbook:
        sheet_names = workbook.sheet_names
    for sheet in sheet_names:
        if normalize_text(sheet) == "precificacao":
            return sheet
    return sheet_names[0]


def detect_reference_columns(df: pd.DataFrame) -> tuple[list[str], str]:
    code_columns = [column for column in df.columns if "codigo" in normalize_text(column) or normalize_text(column) == "cod"]
    compra_columns = [column for column in df.columns if "compra" in normalize_text(column)]
    if not code_columns:
        raise ValueError("Nenhuma coluna de codigo encontrada na planilha de referencia.")
    if not compra_columns:
        raise ValueError("Nenhuma coluna de compra encontrada na planilha de referencia.")

    def purchase_priority(column: str) -> tuple[int, str]:
        normalized = normalize_text(column)
        if normalized == "compra":
            return (0, normalized)
        if "prazo" in normalized:
            return (1, normalized)
        if "vista" in normalized:
            return (2, normalized)
        return (3, normalized)

    return code_columns, sorted(compra_columns, key=purchase_priority)[0]


def read_reference_prices(reference_path: Path) -> dict[str, float]:
    sheet_name = choose_reference_sheet(reference_path)
    df = pd.read_excel(reference_path, sheet_name=sheet_name, engine="openpyxl")
    code_columns, price_column = detect_reference_columns(df)
    prices: dict[str, float] = {}
    for _, row in df.iterrows():
        price = parse_brazilian_number(row.get(price_column))
        if price is None or price <= 0:
            continue
        for code_column in code_columns:
            normalized_code = normalize_product_code(row.get(code_column))
            if normalized_code:
                prices.setdefault(normalized_code, float(price))
    return prices


def build_conference_report(
    order_items: pd.DataFrame,
    quote_items: pd.DataFrame,
    reference_prices: dict[str, float],
    tolerance_percent: float = PRICE_TOLERANCE_PERCENT,
) -> pd.DataFrame:
    quote_by_code = {
        str(row["codigo_normalizado"]): row
        for _, row in quote_items.dropna(subset=["codigo_normalizado"]).iterrows()
        if str(row.get("codigo_normalizado", "")).strip()
    }
    rows = []
    for _, order in order_items.iterrows():
        code = str(order.get("codigo", "") or "").strip()
        normalized_code = str(order.get("codigo_normalizado", "") or "").strip()
        ordered_quantity = order.get("quantidade_pedida")
        quote = quote_by_code.get(normalized_code)
        reference_price = reference_prices.get(normalized_code)

        quoted_quantity = pd.NA
        quoted_price = pd.NA
        difference_percent = pd.NA
        extraction_error = False

        if quote is None:
            status = "NAO_ENCONTRADO_COTACAO"
        else:
            extraction_error = str(quote.get("status_extracao", "") or "") == "ERRO_EXTRACAO"
            quoted_quantity = quote.get("quantidade_cotada")
            quoted_price = quote.get("preco_cotado")
            if extraction_error:
                status = "ERRO_EXTRACAO"
            elif reference_price is None:
                status = "NAO_ENCONTRADO_REFERENCIA"
            else:
                quoted_price_number = parse_brazilian_number(quoted_price)
                if quoted_price_number is not None and reference_price:
                    difference_percent = round(((quoted_price_number - reference_price) / reference_price) * 100, 2)
                price_ok = pd.notna(difference_percent) and abs(float(difference_percent)) <= tolerance_percent
                quantity_ok = parse_brazilian_number(ordered_quantity) == parse_brazilian_number(quoted_quantity)
                if price_ok and quantity_ok:
                    status = "OK"
                elif price_ok:
                    status = "QUANTIDADE_DIVERGENTE"
                elif quantity_ok:
                    status = "PRECO_DIVERGENTE"
                else:
                    status = "PRECO_E_QUANTIDADE_DIVERGENTES"

        rows.append(
            {
                "codigo": code,
                "quantidade_pedida": ordered_quantity,
                "quantidade_cotada": quoted_quantity,
                "preco_cotado": quoted_price,
                "preco_referencia": reference_price if reference_price is not None else pd.NA,
                "diferenca_percentual": difference_percent,
                "status": status,
            }
        )
    return pd.DataFrame(rows, columns=CONFERENCE_REPORT_COLUMNS)


def conference_report_to_excel_bytes(report: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report.to_excel(writer, index=False, sheet_name="Conferencia")
    return output.getvalue()


def display_name_from_file(file_path: Path) -> str:
    name = re.sub(r"[_\-]+", " ", file_path.stem).strip()
    return re.sub(r"\s+", " ", name)


def sanitize_filename_part(value: str) -> str:
    text = re.sub(r'[<>:"/\\|?*]', "", str(value or "").strip())
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .") or "Sem nome"


def standard_order_filename(fornecedor: str, suffix: str, include_time: bool = False) -> str:
    safe_supplier = sanitize_filename_part(fornecedor)
    date_part = datetime.now().strftime("%Y-%m-%d")
    time_part = datetime.now().strftime("_%H%M") if include_time else ""
    return f"Pedido {safe_supplier} {date_part}{time_part}{suffix.lower()}"


def parse_sent_at(value: object | None) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    text = str(value or "").strip()
    if text:
        try:
            return datetime.fromisoformat(text)
        except ValueError:
            pass
    return datetime.now()


def supplier_order_filename(sent_at: object | None, suffix: str, include_time: bool = False) -> str:
    sent_datetime = parse_sent_at(sent_at)
    date_part = sent_datetime.strftime("%Y-%m-%d")
    time_part = sent_datetime.strftime("_%H%M") if include_time else ""
    return f"Pedido {date_part}{time_part}{suffix.lower()}"


def resolve_unique_supplier_destination(folder: Path, source_file: Path, sent_at: object | None) -> Path:
    destination = folder / supplier_order_filename(sent_at, source_file.suffix)
    if not destination.exists():
        return destination

    destination_with_time = folder / supplier_order_filename(sent_at, source_file.suffix, include_time=True)
    if not destination_with_time.exists():
        return destination_with_time

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    return folder / f"Pedido {timestamp}{source_file.suffix.lower()}"


def resolve_unique_destination(folder: Path, fornecedor: str, suffix: str) -> Path:
    destination = folder / standard_order_filename(fornecedor, suffix)
    if not destination.exists():
        return destination

    destination_with_time = folder / standard_order_filename(fornecedor, suffix, include_time=True)
    if not destination_with_time.exists():
        return destination_with_time

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    safe_supplier = sanitize_filename_part(fornecedor)
    return folder / f"Pedido {safe_supplier} {timestamp}{suffix.lower()}"


def write_uploaded_file(uploaded_file, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as file:
        file.write(uploaded_file.getbuffer())


@st.cache_data(show_spinner=False)
def load_fornecedores() -> pd.DataFrame:
    if not FORNECEDORES_FILE.exists():
        return pd.DataFrame(columns=["fornecedor", "email", "cc", "pasta"])

    df = pd.read_excel(FORNECEDORES_FILE, engine="openpyxl")
    df.columns = [normalize_text(column).replace(" ", "_") for column in df.columns]
    df = df.rename(columns={"s": "fornecedor", "supplier": "fornecedor", "nome": "fornecedor"})

    required_columns = {"fornecedor", "email"}
    if not required_columns.issubset(df.columns):
        missing = ", ".join(sorted(required_columns - set(df.columns)))
        raise ValueError(f"Colunas ausentes em fornecedores.xlsx: {missing}")

    for column in ("cc", "pasta"):
        if column not in df.columns:
            df[column] = ""

    df = df[["fornecedor", "email", "cc", "pasta"]].dropna(how="all")
    df["fornecedor"] = df["fornecedor"].astype(str).str.strip()
    df["email"] = df["email"].astype(str).str.strip()
    df["cc"] = df["cc"].fillna("").astype(str).str.strip()
    df["pasta"] = df["pasta"].fillna("").astype(str).str.strip()
    df = df[(df["fornecedor"] != "") & (df["email"] != "")]
    df["fornecedor_normalizado"] = df["fornecedor"].map(normalize_text)
    return df


def list_excel_files() -> list[Path]:
    return sorted(
        file_path
        for file_path in PEDIDOS_DIR.iterdir()
        if file_path.is_file() and file_path.suffix.lower() in EXCEL_EXTENSIONS
    )


def find_supplier_for_file(file_path: Path, fornecedores: pd.DataFrame) -> tuple[str, str | None, str, str]:
    file_supplier = display_name_from_file(file_path)
    normalized_file = normalize_text(file_supplier)

    if fornecedores.empty:
        return file_supplier, None, "", ""

    exact_match = fornecedores[fornecedores["fornecedor_normalizado"] == normalized_file]
    if not exact_match.empty:
        row = exact_match.iloc[0]
        return str(row["fornecedor"]), str(row["email"]), str(row.get("cc", "")), str(row.get("pasta", ""))

    for _, row in fornecedores.iterrows():
        supplier_key = str(row["fornecedor_normalizado"])
        if supplier_key and (normalized_file.startswith(supplier_key) or supplier_key in normalized_file):
            return str(row["fornecedor"]), str(row["email"]), str(row.get("cc", "")), str(row.get("pasta", ""))

    return file_supplier, None, "", ""


def build_pedidos(fornecedores: pd.DataFrame) -> list[Pedido]:
    pedidos = []
    for file_path in list_excel_files():
        fornecedor, email, cc, pasta = find_supplier_for_file(file_path, fornecedores)
        pedidos.append(Pedido(fornecedor=fornecedor, arquivo=file_path, email=email, cc=cc, pasta=pasta))
    return pedidos


def get_gmail_service() -> AuthorizedSession:
    credentials = None

    if TOKEN_FILE.exists():
        credentials = Credentials.from_authorized_user_file(TOKEN_FILE, GMAIL_SCOPES)

    if credentials and credentials.expired and credentials.refresh_token:
        try:
            request_session = configure_google_session(requests.Session())
            credentials.refresh(Request(session=request_session))
        except RefreshError as exc:
            if "invalid_grant" not in str(exc):
                raise
            TOKEN_FILE.unlink(missing_ok=True)
            credentials = None
        else:
            TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")

    if not credentials or not credentials.valid:
        if not CREDENTIALS_FILE.exists():
            raise RuntimeError(
                "Arquivo credentials.json nao encontrado. Baixe o OAuth Client do Google Cloud "
                "e coloque na pasta do projeto."
            )

        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, GMAIL_SCOPES)
        credentials = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")

    if not credentials:
        raise RuntimeError("Nao foi possivel autenticar no Gmail.")

    return configure_google_session(AuthorizedSession(credentials))


def send_gmail_raw_message(session: AuthorizedSession, raw_message: dict[str, str]) -> None:
    response = session.post(GMAIL_SEND_URL, json=raw_message, timeout=30)
    response.raise_for_status()


def normalize_email_address(value: str | None) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def validate_email_address(email: str, field_name: str = "email") -> None:
    if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+", email):
        raise ValueError(f"Email invalido em {field_name}: {email}")


def parse_recipients(value: str) -> list[str]:
    recipients = re.split(r"[;,]", value or "")
    normalized_recipients = []
    for email in recipients:
        normalized_email = normalize_email_address(email)
        if not normalized_email:
            continue
        validate_email_address(normalized_email, "CC")
        normalized_recipients.append(normalized_email)
    return normalized_recipients


def format_log_date(value: object) -> str:
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value or "").strip()


def total_recipients(primary_email: str | None, cc: str = "") -> int:
    total = 1 if normalize_email_address(primary_email) else 0
    return total + len(parse_recipients(cc))


def append_text_signature(body_text: str) -> str:
    return f"{body_text}\n\n{SIGNATURE_NAME}\n{SIGNATURE_EMAIL}\n{SIGNATURE_PHONE}"


def signature_image_html(image_cid: str | None) -> str:
    if not image_cid:
        return ""

    return (
        f'<div style="margin-top: 10px;">'
        f'<img src="cid:{image_cid}" alt="Assinatura" '
        f'width="{SIGNATURE_IMAGE_WIDTH}" height="{SIGNATURE_IMAGE_HEIGHT}" '
        f'style="width: {SIGNATURE_IMAGE_WIDTH}px; max-width: 100%; height: auto;">'
        f"</div>"
    )


def signature_contact_html() -> str:
    return (
        f'<p style="margin-bottom: 8px;">'
        f'<strong style="font-size: 16px;">{SIGNATURE_NAME}</strong><br>'
        f'<a href="mailto:{SIGNATURE_EMAIL}">{SIGNATURE_EMAIL}</a><br>'
        f"{SIGNATURE_PHONE}"
        f"</p>"
    )


def build_plain_text_body(pedido: Pedido) -> str:
    return append_text_signature(EMAIL_BODY_TEXT.format(fornecedor=pedido.fornecedor))


def build_pendencia_plain_text_body(pendencia: PendenciaEmail) -> str:
    return append_text_signature(PENDENCIA_BODY_TEXT.format(fornecedor=pendencia.fornecedor))


def build_html_body(pedido: Pedido, image_cid: str | None) -> str:
    return f"""
    <html>
      <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #222222;">
        <p>Bom dia,</p>
        <p>Segue em anexo o pedido de {pedido.fornecedor}.</p>
        <p>Atenciosamente,</p>
        {signature_contact_html()}
        {signature_image_html(image_cid)}
      </body>
    </html>
    """


def build_pendencia_html_body(pendencia: PendenciaEmail, image_cid: str | None) -> str:
    return f"""
    <html>
      <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #222222;">
        <p>Bom dia,</p>
        <p>Solicito saldo pendente de {pendencia.fornecedor}, por favor.</p>
        <p>Atenciosamente,</p>
        {signature_contact_html()}
        {signature_image_html(image_cid)}
      </body>
    </html>
    """


def build_gmail_message(pedido: Pedido, cc: str = "") -> dict[str, str]:
    if not pedido.email:
        raise ValueError("Fornecedor sem email cadastrado.")

    to_email = normalize_email_address(pedido.email)
    validate_email_address(to_email, "email do fornecedor")

    message = EmailMessage()
    message["To"] = to_email
    cc_recipients = parse_recipients(cc)
    if cc_recipients:
        message["Cc"] = ", ".join(cc_recipients)
    message["Subject"] = EMAIL_SUBJECT_TEMPLATE.format(fornecedor=pedido.fornecedor)
    message.set_content(build_plain_text_body(pedido))

    signature_cid = make_msgid(domain="assinatura.local")[1:-1] if SIGNATURE_IMAGE_FILE.exists() else None
    message.add_alternative(build_html_body(pedido, signature_cid), subtype="html")

    if signature_cid:
        html_part = message.get_payload()[1]
        with SIGNATURE_IMAGE_FILE.open("rb") as signature_image:
            html_part.add_related(
                signature_image.read(),
                maintype="image",
                subtype="png",
                cid=f"<{signature_cid}>",
                filename=SIGNATURE_IMAGE_FILE.name,
            )

    with pedido.arquivo.open("rb") as attachment:
        message.add_attachment(
            attachment.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=pedido.arquivo.name,
        )

    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": encoded_message}


def build_gmail_pendencia_message(pendencia: PendenciaEmail, cc: str = "") -> dict[str, str]:
    if not pendencia.email:
        raise ValueError("Fornecedor sem email cadastrado.")

    to_email = normalize_email_address(pendencia.email)
    validate_email_address(to_email, "email do fornecedor")

    message = EmailMessage()
    message["To"] = to_email
    cc_recipients = parse_recipients(cc)
    if cc_recipients:
        message["Cc"] = ", ".join(cc_recipients)
    message["Subject"] = PENDENCIA_SUBJECT_TEMPLATE.format(fornecedor=pendencia.fornecedor)
    message.set_content(build_pendencia_plain_text_body(pendencia))

    signature_cid = make_msgid(domain="assinatura.local")[1:-1] if SIGNATURE_IMAGE_FILE.exists() else None
    message.add_alternative(build_pendencia_html_body(pendencia, signature_cid), subtype="html")

    if signature_cid:
        html_part = message.get_payload()[1]
        with SIGNATURE_IMAGE_FILE.open("rb") as signature_image:
            html_part.add_related(
                signature_image.read(),
                maintype="image",
                subtype="png",
                cid=f"<{signature_cid}>",
                filename=SIGNATURE_IMAGE_FILE.name,
            )

    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": encoded_message}


def send_email_with_gmail(pedido: Pedido, cc: str = "") -> None:
    session = get_gmail_service()
    send_gmail_raw_message(session, build_gmail_message(pedido, cc))


def send_pendencia_with_gmail(pendencia: PendenciaEmail, cc: str = "") -> None:
    session = get_gmail_service()
    send_gmail_raw_message(session, build_gmail_pendencia_message(pendencia, cc))


def supplier_folder_from_row(row: pd.Series) -> Path | None:
    folder_value = str(row.get("pasta", "") or "").strip()
    if not folder_value:
        return None
    return Path(os.path.expandvars(os.path.expanduser(folder_value)))


def save_uploaded_order(uploaded_file, fornecedor_row: pd.Series) -> tuple[Path, Path]:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix not in EXCEL_EXTENSIONS:
        allowed = ", ".join(sorted(EXCEL_EXTENSIONS))
        raise ValueError(f"Arquivo invalido. Use um arquivo Excel: {allowed}.")

    fornecedor = str(fornecedor_row["fornecedor"]).strip()
    supplier_folder = supplier_folder_from_row(fornecedor_row)
    if supplier_folder is None:
        raise ValueError("Fornecedor sem pasta cadastrada na coluna pasta do fornecedores.xlsx.")
    if not supplier_folder.exists() or not supplier_folder.is_dir():
        raise ValueError(f"Pasta do fornecedor nao encontrada: {supplier_folder}")

    pedidos_destination = resolve_unique_destination(PEDIDOS_DIR, fornecedor, suffix)
    supplier_destination = supplier_folder / pedidos_destination.name
    if supplier_destination.exists():
        supplier_destination = resolve_unique_destination(supplier_folder, fornecedor, suffix)
        pedidos_destination = PEDIDOS_DIR / supplier_destination.name
        if pedidos_destination.exists():
            pedidos_destination = resolve_unique_destination(PEDIDOS_DIR, fornecedor, suffix)
            supplier_destination = supplier_folder / pedidos_destination.name

    write_uploaded_file(uploaded_file, pedidos_destination)
    shutil.copy2(pedidos_destination, supplier_destination)
    return pedidos_destination, supplier_destination


def destination_for_sent_file(file_path: Path) -> Path:
    destination = ENVIADOS_DIR / file_path.name
    if not destination.exists():
        return destination

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return ENVIADOS_DIR / f"{file_path.stem}_{timestamp}{file_path.suffix}"


def move_to_enviados(file_path: Path) -> Path:
    destination = destination_for_sent_file(file_path)
    return Path(shutil.move(str(file_path), str(destination)))


def supplier_folder_from_pedido(pedido: Pedido) -> Path:
    folder_value = str(pedido.pasta or "").strip()
    if not folder_value:
        raise ValueError("Fornecedor sem pasta cadastrada na coluna pasta do fornecedores.xlsx.")

    supplier_folder = Path(os.path.expandvars(os.path.expanduser(folder_value)))
    if not supplier_folder.exists() or not supplier_folder.is_dir():
        raise ValueError(f"Pasta do fornecedor nao encontrada: {supplier_folder}")
    return supplier_folder


def copy_order_to_supplier_folder(pedido: Pedido, sent_at: object | None = None) -> Path:
    supplier_folder = supplier_folder_from_pedido(pedido)
    if not pedido.arquivo.exists() or not pedido.arquivo.is_file():
        raise ValueError(f"Arquivo do pedido nao encontrado: {pedido.arquivo}")

    destination = destination_for_supplier_copy(pedido.arquivo, supplier_folder, pedido.fornecedor, sent_at)
    return Path(shutil.copy2(pedido.arquivo, destination))


def finalize_sent_order(pedido: Pedido, sent_at: object | None = None) -> tuple[Path, Path]:
    supplier_copy = copy_order_to_supplier_folder(pedido, sent_at=sent_at)
    sent_file = move_to_enviados(pedido.arquivo)
    return supplier_copy, sent_file


def destination_for_supplier_copy(
    source_file: Path,
    supplier_folder: Path,
    fornecedor: str,
    sent_at: object | None = None,
) -> Path:
    return resolve_unique_supplier_destination(supplier_folder, source_file, sent_at)


def backfill_sent_orders_to_supplier_folders(data_envio_prefix: str | None = None) -> dict[str, list[str]]:
    result: dict[str, list[str]] = {"copiados": [], "ignorados": []}
    if not LOG_FILE.exists():
        result["ignorados"].append(f"Historico nao encontrado: {LOG_FILE}")
        return result

    ensure_log_columns()
    logs = pd.read_csv(LOG_FILE, encoding="utf-8-sig").fillna("")
    for _, row in logs.iterrows():
        if data_envio_prefix and not str(row.get("data_envio", "")).startswith(data_envio_prefix):
            continue

        fornecedor = str(row.get("fornecedor", "") or "").strip()
        arquivo = str(row.get("arquivo", "") or "").strip()
        pasta = str(row.get("pasta", "") or "").strip()
        if not fornecedor or not arquivo or not pasta:
            result["ignorados"].append(f"{fornecedor or 'Sem fornecedor'}: historico sem arquivo ou pasta")
            continue

        source_file = ENVIADOS_DIR / arquivo
        if not source_file.exists() or not source_file.is_file():
            result["ignorados"].append(f"{fornecedor}: arquivo nao encontrado em enviados/ ({arquivo})")
            continue

        supplier_folder = Path(os.path.expandvars(os.path.expanduser(pasta)))
        if not supplier_folder.exists() or not supplier_folder.is_dir():
            result["ignorados"].append(f"{fornecedor}: pasta nao encontrada ({supplier_folder})")
            continue

        destination = destination_for_supplier_copy(source_file, supplier_folder, fornecedor, row.get("data_envio", ""))
        shutil.copy2(source_file, destination)
        result["copiados"].append(str(destination))

    return result


def ensure_csv_columns(csv_path: Path, columns: list[str]) -> None:
    if not csv_path.exists():
        return

    logs = pd.read_csv(csv_path, encoding="utf-8-sig")
    changed = False
    for column in columns:
        if column not in logs.columns:
            logs[column] = ""
            changed = True

    if changed:
        logs = logs[columns + [column for column in logs.columns if column not in columns]]
        logs.to_csv(csv_path, index=False, encoding="utf-8-sig")


def ensure_log_columns() -> None:
    ensure_csv_columns(LOG_FILE, PEDIDO_LOG_COLUMNS)


def ensure_pendencia_log_columns() -> None:
    ensure_csv_columns(PENDENCIA_LOG_FILE, PENDENCIA_LOG_COLUMNS)


def append_log(
    pedido: Pedido,
    cc: str = "",
    data_pedido: object = "",
    observacao: str = "",
    data_envio: object | None = None,
) -> None:
    ensure_log_columns()
    file_exists = LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=PEDIDO_LOG_COLUMNS)
        if not file_exists:
            writer.writeheader()
        data_envio_text = parse_sent_at(data_envio).strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow(
            {
                "data": data_envio_text,
                "fornecedor": pedido.fornecedor,
                "arquivo": pedido.arquivo.name,
                "email": pedido.email,
                "cc": cc,
                "data_envio": data_envio_text,
                "data_pedido": format_log_date(data_pedido),
                "tipo": "pedido",
                "assunto": EMAIL_SUBJECT_TEMPLATE.format(fornecedor=pedido.fornecedor),
                "pasta": pedido.pasta,
                "total_destinatarios": total_recipients(pedido.email, cc),
                "observacao": str(observacao or "").strip(),
            }
        )


def append_pendencia_log(
    pendencia: PendenciaEmail,
    cc: str = "",
    data_referencia: object = "",
    observacao: str = "",
) -> None:
    ensure_pendencia_log_columns()
    file_exists = PENDENCIA_LOG_FILE.exists()
    with PENDENCIA_LOG_FILE.open("a", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=PENDENCIA_LOG_COLUMNS)
        if not file_exists:
            writer.writeheader()
        data_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow(
            {
                "data": data_envio,
                "fornecedor": pendencia.fornecedor,
                "email": pendencia.email,
                "cc": cc,
                "assunto": PENDENCIA_SUBJECT_TEMPLATE.format(fornecedor=pendencia.fornecedor),
                "data_envio": data_envio,
                "data_referencia": format_log_date(data_referencia),
                "tipo": "pendencia",
                "total_destinatarios": total_recipients(pendencia.email, cc),
                "observacao": str(observacao or "").strip(),
            }
        )


def format_pedido_option(pedido: Pedido) -> str:
    status = pedido.email or "email nao encontrado"
    return f"{pedido.arquivo.name} | {pedido.fornecedor} | {status}"


def build_pendencias(fornecedores: pd.DataFrame) -> list[PendenciaEmail]:
    pendencias = []
    for _, row in fornecedores.iterrows():
        fornecedor = str(row.get("fornecedor", "") or "").strip()
        email = str(row.get("email", "") or "").strip()
        cc = str(row.get("cc", "") or "").strip()
        if fornecedor and email:
            pendencias.append(PendenciaEmail(fornecedor=fornecedor, email=email, cc=cc))
    return pendencias


def format_pendencia_option(pendencia: PendenciaEmail) -> str:
    return f"{pendencia.fornecedor} | {pendencia.email}"


def render_fornecedores_status(fornecedores: pd.DataFrame) -> None:
    with st.expander("Fornecedores cadastrados", expanded=False):
        if fornecedores.empty:
            st.info("Nenhum fornecedor cadastrado em fornecedores.xlsx.")
        else:
            st.dataframe(fornecedores[["fornecedor", "email", "cc", "pasta"]], use_container_width=True, hide_index=True)


def render_upload_order(fornecedores: pd.DataFrame) -> None:
    st.subheader("Adicionar pedido")

    if "upload_success" in st.session_state:
        success = st.session_state.pop("upload_success")
        st.success("Pedido adicionado em pedidos/ e copiado para a pasta do fornecedor.")
        st.write(f"Arquivo em pedidos: `{success['pedidos']}`")
        st.write(f"Copia do fornecedor: `{success['fornecedor']}`")

    if fornecedores.empty:
        st.info("Cadastre fornecedores antes de adicionar pedidos.")
        return

    uploaded_file = st.file_uploader(
        "Planilha do pedido",
        type=[extension.removeprefix(".") for extension in sorted(EXCEL_EXTENSIONS)],
    )
    selected_supplier = st.selectbox(
        "Fornecedor",
        fornecedores["fornecedor"].tolist(),
        index=0,
        key="upload_supplier",
    )

    fornecedor_row = fornecedores[fornecedores["fornecedor"] == selected_supplier].iloc[0]
    supplier_folder = supplier_folder_from_row(fornecedor_row)
    suffix = Path(uploaded_file.name).suffix.lower() if uploaded_file else ".xlsx"
    preview_name = standard_order_filename(selected_supplier, suffix)

    st.write(f"Nome padronizado: **{preview_name}**")
    st.write(f"Destino em pedidos: `{PEDIDOS_DIR / preview_name}`")

    if supplier_folder:
        st.write(f"Destino do fornecedor: `{supplier_folder / preview_name}`")
        if not supplier_folder.exists():
            st.warning("A pasta cadastrada para este fornecedor ainda nao existe ou nao esta acessivel.")
    else:
        st.warning("Este fornecedor nao tem pasta cadastrada na coluna pasta.")

    if st.button("Adicionar pedido", type="secondary", disabled=uploaded_file is None):
        try:
            pedidos_destination, supplier_destination = save_uploaded_order(uploaded_file, fornecedor_row)
        except Exception as exc:
            st.error(f"Erro ao adicionar pedido: {exc}")
        else:
            st.session_state["upload_success"] = {
                "pedidos": str(pedidos_destination),
                "fornecedor": str(supplier_destination),
            }
            st.cache_data.clear()
            st.rerun()


def render_gmail_config_help() -> None:
    with st.expander("Configuracao do Gmail", expanded=False):
        st.write("Coloque o arquivo `credentials.json` do Google Cloud na pasta do projeto.")
        st.code(str(CREDENTIALS_FILE), language="text")


def render_email_preview(pedido: Pedido) -> None:
    st.markdown(build_html_body(pedido, image_cid=None), unsafe_allow_html=True)
    if SIGNATURE_IMAGE_FILE.exists():
        st.image(
            str(SIGNATURE_IMAGE_FILE),
            caption="Imagem de assinatura que sera anexada ao email.",
            width=SIGNATURE_IMAGE_WIDTH,
        )
    else:
        st.caption("Para incluir a imagem da assinatura, salve o arquivo como assinatura.png na pasta do projeto.")


def render_pendencia_preview(pendencia: PendenciaEmail) -> None:
    st.markdown(build_pendencia_html_body(pendencia, image_cid=None), unsafe_allow_html=True)
    if SIGNATURE_IMAGE_FILE.exists():
        st.image(
            str(SIGNATURE_IMAGE_FILE),
            caption="Imagem de assinatura que sera anexada ao email.",
            width=SIGNATURE_IMAGE_WIDTH,
        )
    else:
        st.caption("Para incluir a imagem da assinatura, salve o arquivo como assinatura.png na pasta do projeto.")


def render_logs() -> None:
    if LOG_FILE.exists():
        with st.expander("Ultimos envios", expanded=False):
            ensure_log_columns()
            logs = pd.read_csv(LOG_FILE, encoding="utf-8-sig")
            st.dataframe(logs.tail(20), use_container_width=True, hide_index=True)


def render_pendencia_logs() -> None:
    if PENDENCIA_LOG_FILE.exists():
        with st.expander("Ultimas pendencias enviadas", expanded=False):
            ensure_pendencia_log_columns()
            logs = pd.read_csv(PENDENCIA_LOG_FILE, encoding="utf-8-sig")
            st.dataframe(logs.tail(20), use_container_width=True, hide_index=True)


def list_sent_order_files_for_supplier(fornecedor: str) -> list[Path]:
    supplier_prefix = normalize_text(fornecedor)
    files = [
        file_path
        for file_path in ENVIADOS_DIR.glob("*.xlsx")
        if normalize_text(file_path.stem).startswith(supplier_prefix)
    ]
    return sorted(files, key=lambda path: path.stat().st_mtime, reverse=True)


def render_conferencia_tab() -> None:
    st.subheader("Conferencia de cotacoes")

    suppliers = list(REFERENCE_FILE_NAMES)
    selected_supplier = st.selectbox("Fornecedor", suppliers, key="conferencia_fornecedor")
    reference_file = find_reference_file(selected_supplier)
    sent_order_files = list_sent_order_files_for_supplier(selected_supplier)

    if reference_file is None or not reference_file.exists():
        st.error(f"Planilha de referencia nao encontrada para {selected_supplier}.")
        if reference_file is not None:
            st.write(f"Caminho esperado: `{reference_file}`")
        return

    selected_order = None
    if sent_order_files:
        selected_order = st.selectbox(
            "Pedido enviado",
            sent_order_files,
            format_func=lambda path: path.name,
            index=0,
            key="conferencia_pedido",
        )
    else:
        st.warning(f"Nenhum pedido enviado encontrado em enviados/ para {selected_supplier}.")

    uploaded_pdfs = st.file_uploader(
        "PDF(s) da cotacao recebida",
        type=["pdf"],
        accept_multiple_files=True,
        key="conferencia_pdfs",
    )

    if selected_order:
        st.write(f"Pedido base: `{selected_order}`")
    st.write(f"Referencia de precos: `{reference_file}`")

    if st.button(
        "Gerar conferencia",
        type="primary",
        disabled=not selected_order or not uploaded_pdfs,
        key="gerar_conferencia",
    ):
        try:
            order_items = read_sent_order_items(selected_order)
            reference_prices = read_reference_prices(reference_file)
            quote_frames = [
                extract_quote_items_from_pdf(BytesIO(uploaded_pdf.getvalue()), selected_supplier)
                for uploaded_pdf in uploaded_pdfs
            ]
            quote_items = pd.concat(quote_frames, ignore_index=True) if quote_frames else quote_items_dataframe([])
            report = build_conference_report(order_items, quote_items, reference_prices)
            st.session_state["conferencia_report"] = report
            st.session_state["conferencia_excel"] = conference_report_to_excel_bytes(report)
            st.session_state["conferencia_nome"] = (
                f"conferencia_{selected_supplier.lower()}_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
            )
        except Exception as exc:
            st.error(f"Erro ao gerar conferencia: {exc}")

    report = st.session_state.get("conferencia_report")
    if report is not None:
        st.dataframe(report, use_container_width=True, hide_index=True)
        st.download_button(
            "Baixar relatorio Excel",
            data=st.session_state.get("conferencia_excel", b""),
            file_name=st.session_state.get("conferencia_nome", "conferencia_cotacao.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_conferencia",
        )


def render_pedidos_tab(fornecedores: pd.DataFrame, pedidos: list[Pedido]) -> None:
    col_refresh, col_paths = st.columns([1, 3])
    with col_refresh:
        if st.button("Atualizar", key="refresh_pedidos"):
            st.cache_data.clear()
            st.rerun()
    with col_paths:
        st.write(f"Pedidos pendentes: **{len(pedidos)}**")

    render_upload_order(fornecedores)
    st.divider()
    render_fornecedores_status(fornecedores)
    render_gmail_config_help()

    if not FORNECEDORES_FILE.exists():
        st.warning("Crie ou preencha o arquivo fornecedores.xlsx antes de enviar pedidos.")

    if not pedidos:
        st.info("Nenhum arquivo Excel encontrado na pasta pedidos/.")
        render_logs()
        return

    selected = st.selectbox(
        "Pedido para envio",
        pedidos,
        format_func=format_pedido_option,
        index=0,
    )

    st.subheader("Email gerado")
    st.text_input("Para", value=selected.email or "", disabled=True)
    cc_value = st.text_input(
        "CC",
        value=selected.cc,
        help="Separe varios emails com virgula ou ponto e virgula.",
    )
    st.text_input(
        "Assunto",
        value=EMAIL_SUBJECT_TEMPLATE.format(fornecedor=selected.fornecedor),
        disabled=True,
    )
    data_pedido = st.date_input("Data do pedido", value=date.today(), key="pedido_data_pedido")
    observacao_pedido = st.text_input(
        "Observacao do historico",
        value="",
        placeholder="Ex.: pedido mensal, urgente, campanha...",
        key="pedido_observacao",
    )
    with st.expander("Previa do corpo do email", expanded=True):
        render_email_preview(selected)
    st.write(f"Anexo: **{selected.arquivo.name}**")

    if selected.email is None:
        st.error("Nao foi encontrado email para este fornecedor em fornecedores.xlsx.")
        render_logs()
        return

    if st.button("Enviar", type="primary"):
        try:
            send_email_with_gmail(selected, cc_value)
            sent_at = datetime.now()
            append_log(
                selected,
                cc_value,
                data_pedido=data_pedido,
                observacao=observacao_pedido,
                data_envio=sent_at,
            )
            supplier_copy, destination = finalize_sent_order(selected, sent_at=sent_at)
        except Exception as exc:
            st.error(f"Erro ao enviar: {exc}")
        else:
            st.success(f"Pedido enviado, copiado para a pasta do fornecedor e movido para {destination.name}.")
            st.write(f"Copia do fornecedor: `{supplier_copy}`")
            st.cache_data.clear()

    render_logs()


def render_pendencias_tab(fornecedores: pd.DataFrame) -> None:
    st.subheader("Enviar pedido de saldo pendente")

    if st.button("Atualizar fornecedores", key="refresh_pendencias"):
        st.cache_data.clear()
        st.rerun()

    if fornecedores.empty:
        st.info("Cadastre fornecedores antes de enviar pendencias.")
        render_pendencia_logs()
        return

    pendencias = build_pendencias(fornecedores)
    if not pendencias:
        st.info("Nenhum fornecedor com email cadastrado em fornecedores.xlsx.")
        render_pendencia_logs()
        return

    st.dataframe(fornecedores[["fornecedor", "email", "cc"]], use_container_width=True, hide_index=True)

    selected_names = st.multiselect(
        "Fornecedores para pendencia",
        [pendencia.fornecedor for pendencia in pendencias],
        key="pendencia_fornecedores",
    )
    selected_pendencias = [pendencia for pendencia in pendencias if pendencia.fornecedor in selected_names]
    st.write(f"Fornecedores selecionados: **{len(selected_pendencias)}**")

    if not selected_pendencias:
        st.info("Selecione um ou mais fornecedores para gerar a previa.")
        render_pendencia_logs()
        return

    preview = st.selectbox(
        "Fornecedor da previa",
        selected_pendencias,
        format_func=format_pendencia_option,
        index=0,
        key="pendencia_preview",
    )

    st.text_input("Para", value=preview.email, disabled=True, key="pendencia_para")
    st.text_input("CC", value=preview.cc, disabled=True, key="pendencia_cc")
    st.text_input(
        "Assunto",
        value=PENDENCIA_SUBJECT_TEMPLATE.format(fornecedor=preview.fornecedor),
        disabled=True,
        key="pendencia_assunto",
    )
    data_referencia = st.date_input(
        "Data de referencia",
        value=date.today(),
        key="pendencia_data_referencia",
    )
    observacao_pendencia = st.text_input(
        "Observacao do historico",
        value="",
        placeholder="Ex.: saldo mensal, cobranca de retorno...",
        key="pendencia_observacao",
    )
    with st.expander("Previa do corpo do email", expanded=True):
        render_pendencia_preview(preview)

    if st.button("Enviar pendencias", type="primary", key="send_pendencias"):
        sent = []
        failures = []
        try:
            session = get_gmail_service()
            for pendencia in selected_pendencias:
                try:
                    send_gmail_raw_message(session, build_gmail_pendencia_message(pendencia, pendencia.cc))
                    append_pendencia_log(
                        pendencia,
                        pendencia.cc,
                        data_referencia=data_referencia,
                        observacao=observacao_pendencia,
                    )
                    sent.append(pendencia.fornecedor)
                except Exception as exc:
                    failures.append(f"{pendencia.fornecedor}: {exc}")
        except Exception as exc:
            st.error(f"Erro ao autenticar no Gmail: {exc}")
        else:
            if sent:
                st.success(f"Pendencias enviadas: {', '.join(sent)}.")
            if failures:
                st.error("Algumas pendencias nao foram enviadas:")
                for failure in failures:
                    st.write(f"- {failure}")

    render_pendencia_logs()


def main() -> None:
    ensure_project_structure()

    st.set_page_config(page_title="Envio de Pedidos", page_icon=":material/mail:", layout="centered")
    st.title("Envio de pedidos por email")
    st.caption("Arquivos em pedidos/ sao enviados pelo Gmail e movidos para enviados/.")

    try:
        fornecedores = load_fornecedores()
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    pedidos = build_pedidos(fornecedores)
    pedidos_tab, pendencias_tab, conferencia_tab = st.tabs(
        ["Pedidos", "Pendencia", "Conferencia de cotacoes"]
    )

    with pedidos_tab:
        render_pedidos_tab(fornecedores, pedidos)

    with pendencias_tab:
        render_pendencias_tab(fornecedores)

    with conferencia_tab:
        render_conferencia_tab()


if __name__ == "__main__":
    main()
