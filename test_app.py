from __future__ import annotations

import base64
import os
import sys
import tempfile
import types
import unittest
from io import BytesIO
from email import message_from_bytes
from email.policy import default
from pathlib import Path
from unittest.mock import patch

import pandas as pd


def install_dependency_stubs() -> None:
    streamlit = types.ModuleType("streamlit")
    streamlit.cache_data = lambda **_kwargs: (lambda function: function)
    sys.modules.setdefault("streamlit", streamlit)

    google_auth_transport = types.ModuleType("google.auth.transport")
    google_auth_transport_requests = types.ModuleType("google.auth.transport.requests")
    class StubRequest:
        def __init__(self, *args, **kwargs):
            self.args = args
            self.kwargs = kwargs

    class StubAuthorizedSession:
        def __init__(self, credentials):
            self.credentials = credentials
            self.verify = True
            self.adapters = {}

        def mount(self, prefix, adapter):
            self.adapters[prefix] = adapter

    google_auth_transport_requests.Request = StubRequest
    google_auth_transport_requests.AuthorizedSession = StubAuthorizedSession
    google_auth_exceptions = types.ModuleType("google.auth.exceptions")
    google_auth_exceptions.RefreshError = Exception
    sys.modules.setdefault("google", types.ModuleType("google"))
    sys.modules.setdefault("google.auth", types.ModuleType("google.auth"))
    sys.modules.setdefault("google.auth.exceptions", google_auth_exceptions)
    sys.modules.setdefault("google.auth.transport", google_auth_transport)
    sys.modules.setdefault("google.auth.transport.requests", google_auth_transport_requests)

    google_oauth2 = types.ModuleType("google.oauth2")
    google_oauth2_credentials = types.ModuleType("google.oauth2.credentials")
    class StubCredentials:
        @staticmethod
        def from_authorized_user_file(*_args, **_kwargs):
            return None

    google_oauth2_credentials.Credentials = StubCredentials
    sys.modules.setdefault("google.oauth2", google_oauth2)
    sys.modules.setdefault("google.oauth2.credentials", google_oauth2_credentials)

    google_auth_oauthlib = types.ModuleType("google_auth_oauthlib")
    google_auth_oauthlib_flow = types.ModuleType("google_auth_oauthlib.flow")
    class StubInstalledAppFlow:
        @staticmethod
        def from_client_secrets_file(*_args, **_kwargs):
            return None

    google_auth_oauthlib_flow.InstalledAppFlow = StubInstalledAppFlow
    sys.modules.setdefault("google_auth_oauthlib", google_auth_oauthlib)
    sys.modules.setdefault("google_auth_oauthlib.flow", google_auth_oauthlib_flow)

    google_auth_httplib2 = types.ModuleType("google_auth_httplib2")

    class StubAuthorizedHttp:
        def __init__(self, credentials, http=None):
            self.credentials = credentials
            self.http = http

    google_auth_httplib2.AuthorizedHttp = StubAuthorizedHttp
    sys.modules.setdefault("google_auth_httplib2", google_auth_httplib2)

    googleapiclient = types.ModuleType("googleapiclient")
    googleapiclient_discovery = types.ModuleType("googleapiclient.discovery")
    googleapiclient_discovery.build = lambda *_args, **_kwargs: None
    sys.modules.setdefault("googleapiclient", googleapiclient)
    sys.modules.setdefault("googleapiclient.discovery", googleapiclient_discovery)


install_dependency_stubs()
import app  # noqa: E402


class PendenciaEmailTests(unittest.TestCase):
    def decode_gmail_raw_message(self, raw_message: dict[str, str]):
        return message_from_bytes(base64.urlsafe_b64decode(raw_message["raw"]), policy=default)

    def test_parse_recipients_removes_spaces_inside_email_addresses(self) -> None:
        recipients = app.parse_recipients(
            "estoqueltda@hotmail. com, contato@rboecia.com.br; vendas@rboecia.com.br"
        )

        self.assertEqual(
            recipients,
            [
                "estoqueltda@hotmail.com",
                "contato@rboecia.com.br",
                "vendas@rboecia.com.br",
            ],
        )

    def test_build_pendencia_plain_text_body_uses_fixed_template_and_signature(self) -> None:
        pendencia = app.PendenciaEmail(fornecedor="JAMAICA", email="jamaica@example.com")

        body = app.build_pendencia_plain_text_body(pendencia)

        self.assertIn("Solicito saldo pendente de JAMAICA, por favor.", body)
        self.assertIn(app.SIGNATURE_NAME, body)
        self.assertIn(app.SIGNATURE_EMAIL, body)
        self.assertIn(app.SIGNATURE_PHONE, body)

    def test_build_gmail_pendencia_message_has_no_attachment(self) -> None:
        pendencia = app.PendenciaEmail(
            fornecedor="JAMAICA",
            email="jamaica@example.com",
            cc="a@example.com, b@example.com",
        )

        with patch.object(app, "SIGNATURE_IMAGE_FILE", Path("missing-signature.png")):
            raw_message = app.build_gmail_pendencia_message(pendencia, pendencia.cc)

        message = self.decode_gmail_raw_message(raw_message)
        self.assertEqual(message["To"], "jamaica@example.com")
        self.assertEqual(message["Cc"], "a@example.com, b@example.com")
        self.assertEqual(message["Subject"], "Pendência JAMAICA")
        self.assertNotIn("attachment", message.as_string().lower())
        self.assertNotIn("filename=", message.as_string().lower())

    def test_build_gmail_pendencia_message_normalizes_to_and_cc_headers(self) -> None:
        pendencia = app.PendenciaEmail(
            fornecedor="IMA",
            email="atendimento @rboecia.com.br",
            cc="estoqueltda@hotmail. com, contato@rboecia.com.br",
        )

        with patch.object(app, "SIGNATURE_IMAGE_FILE", Path("missing-signature.png")):
            raw_message = app.build_gmail_pendencia_message(pendencia, pendencia.cc)

        message = self.decode_gmail_raw_message(raw_message)
        self.assertEqual(message["To"], "atendimento@rboecia.com.br")
        self.assertEqual(message["Cc"], "estoqueltda@hotmail.com, contato@rboecia.com.br")

    def test_signature_image_html_uses_reference_size(self) -> None:
        html = app.signature_image_html("signature-cid")

        self.assertIn('width="400"', html)
        self.assertIn('height="200"', html)
        self.assertIn("width: 400px", html)

    def test_append_pendencia_log_writes_separate_history(self) -> None:
        pendencia = app.PendenciaEmail(
            fornecedor="JAMAICA",
            email="jamaica@example.com",
            cc="cc@example.com",
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            log_file = Path(temp_dir) / "pendencias.csv"
            with patch.object(app, "PENDENCIA_LOG_FILE", log_file):
                app.append_pendencia_log(pendencia, pendencia.cc)

            contents = log_file.read_text(encoding="utf-8-sig")

        self.assertIn("data,fornecedor,email,cc,assunto", contents)
        self.assertIn("JAMAICA,jamaica@example.com,cc@example.com,Pendência JAMAICA", contents)

    def test_append_pendencia_log_writes_analysis_fields(self) -> None:
        pendencia = app.PendenciaEmail(
            fornecedor="JAMAICA",
            email="jamaica@example.com",
            cc="cc1@example.com, cc2@example.com",
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            log_file = Path(temp_dir) / "pendencias.csv"
            with patch.object(app, "PENDENCIA_LOG_FILE", log_file):
                app.append_pendencia_log(
                    pendencia,
                    pendencia.cc,
                    data_referencia="2026-06-01",
                    observacao="saldo mensal",
                )

            contents = log_file.read_text(encoding="utf-8-sig")

        self.assertIn("data_envio", contents)
        self.assertIn("data_referencia", contents)
        self.assertIn("tipo", contents)
        self.assertIn("total_destinatarios", contents)
        self.assertIn("observacao", contents)
        self.assertIn("2026-06-01", contents)
        self.assertIn("pendencia", contents)
        self.assertIn(",3,", contents)
        self.assertIn("saldo mensal", contents)


class PedidoLogTests(unittest.TestCase):
    def test_build_gmail_message_normalizes_to_and_cc_headers(self) -> None:
        pedido = app.Pedido(
            fornecedor="IMA",
            arquivo=Path("fornecedores.xlsx"),
            email="atendimento @rboecia.com.br",
            cc="estoqueltda@hotmail. com",
        )

        with patch.object(app, "SIGNATURE_IMAGE_FILE", Path("missing-signature.png")):
            raw_message = app.build_gmail_message(pedido, pedido.cc)

        message = PendenciaEmailTests().decode_gmail_raw_message(raw_message)
        self.assertEqual(message["To"], "atendimento@rboecia.com.br")
        self.assertEqual(message["Cc"], "estoqueltda@hotmail.com")

    def test_append_log_writes_analysis_fields(self) -> None:
        pedido = app.Pedido(
            fornecedor="SAMPEL",
            arquivo=Path("Pedido SAMPEL 2026-06-01.xlsx"),
            email="sampel@example.com",
            cc="cc@example.com",
            pasta="C:\\Fornecedores\\SAMPEL",
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            log_file = Path(temp_dir) / "envios.csv"
            with patch.object(app, "LOG_FILE", log_file):
                app.append_log(
                    pedido,
                    pedido.cc,
                    data_pedido="2026-06-01",
                    observacao="pedido urgente",
                )

            contents = log_file.read_text(encoding="utf-8-sig")

        self.assertIn("data_envio", contents)
        self.assertIn("data_pedido", contents)
        self.assertIn("tipo", contents)
        self.assertIn("assunto", contents)
        self.assertIn("pasta", contents)
        self.assertIn("total_destinatarios", contents)
        self.assertIn("observacao", contents)
        self.assertIn("2026-06-01", contents)
        self.assertIn("pedido", contents)
        self.assertIn("Pedido de Cotacao - SAMPEL", contents)
        self.assertIn(",2,", contents)
        self.assertIn("pedido urgente", contents)


class PedidoFolderCopyTests(unittest.TestCase):
    def test_copy_order_to_supplier_folder_copies_file_with_unique_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            base = Path(temp_dir)
            source = base / "Pedido TESTE 2026-06-05.xlsx"
            supplier_folder = base / "fornecedor"
            supplier_folder.mkdir()
            source.write_bytes(b"pedido")
            (supplier_folder / "Pedido 2026-06-05.xlsx").write_bytes(b"existente")
            pedido = app.Pedido(
                fornecedor="TESTE",
                arquivo=source,
                email="teste@example.com",
                pasta=str(supplier_folder),
            )

            copied = app.copy_order_to_supplier_folder(pedido, sent_at="2026-06-05 09:30:00")

            self.assertTrue(copied.exists())
            self.assertEqual(copied.read_bytes(), b"pedido")
            self.assertNotEqual(copied.name, "Pedido 2026-06-05.xlsx")
            self.assertTrue(copied.name.startswith("Pedido 2026-06-05_"))
            self.assertTrue(source.exists())

    def test_copy_order_to_supplier_folder_requires_valid_folder(self) -> None:
        pedido = app.Pedido(
            fornecedor="TESTE",
            arquivo=Path("pedido.xlsx"),
            email="teste@example.com",
            pasta="",
        )

        with self.assertRaisesRegex(ValueError, "Fornecedor sem pasta cadastrada"):
            app.copy_order_to_supplier_folder(pedido)

    def test_finalize_sent_order_copies_then_moves_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            base = Path(temp_dir)
            pedidos_dir = base / "pedidos"
            enviados_dir = base / "enviados"
            supplier_folder = base / "fornecedor"
            pedidos_dir.mkdir()
            enviados_dir.mkdir()
            supplier_folder.mkdir()
            source = pedidos_dir / "TESTE 05.06.xlsx"
            source.write_bytes(b"pedido")
            pedido = app.Pedido(
                fornecedor="TESTE",
                arquivo=source,
                email="teste@example.com",
                pasta=str(supplier_folder),
            )

            with patch.object(app, "ENVIADOS_DIR", enviados_dir):
                supplier_copy, sent_file = app.finalize_sent_order(pedido, sent_at="2026-06-05 09:30:00")

            self.assertTrue(supplier_copy.exists())
            self.assertEqual(supplier_copy.name, "Pedido 2026-06-05.xlsx")
            self.assertTrue(sent_file.exists())
            self.assertFalse(source.exists())
            self.assertEqual(supplier_copy.read_bytes(), b"pedido")
            self.assertEqual(sent_file.read_bytes(), b"pedido")

    def test_backfill_sent_orders_copies_from_enviados_without_moving(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            base = Path(temp_dir)
            enviados_dir = base / "enviados"
            logs_dir = base / "logs"
            supplier_folder = base / "fornecedor"
            enviados_dir.mkdir()
            logs_dir.mkdir()
            supplier_folder.mkdir()
            sent_file = enviados_dir / "TESTE 05.06.xlsx"
            sent_file.write_bytes(b"pedido")
            log_file = logs_dir / "envios.csv"
            log_file.write_text(
                "data,fornecedor,arquivo,email,cc,data_envio,data_pedido,tipo,assunto,pasta,total_destinatarios,observacao\n"
                f"2026-06-05 09:00:00,TESTE,{sent_file.name},teste@example.com,,2026-06-05 09:00:00,2026-06-05,pedido,Pedido de Cotacao - TESTE,{supplier_folder},1,\n",
                encoding="utf-8-sig",
            )

            with (
                patch.object(app, "ENVIADOS_DIR", enviados_dir),
                patch.object(app, "LOG_FILE", log_file),
            ):
                result = app.backfill_sent_orders_to_supplier_folders()

            self.assertEqual(result["copiados"], [str(supplier_folder / "Pedido 2026-06-05.xlsx")])
            self.assertEqual(result["ignorados"], [])
            self.assertTrue(sent_file.exists())
            self.assertTrue((supplier_folder / "Pedido 2026-06-05.xlsx").exists())


class FornecedoresTests(unittest.TestCase):
    def test_load_fornecedores_accepts_s_column_as_supplier_name(self) -> None:
        fornecedores = pd.DataFrame(
            [
                {
                    "S": "JAMAICA",
                    "email": "jamaica@example.com",
                    "cc": "cc@example.com",
                    "pasta": "",
                }
            ]
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            fornecedores_file = Path(temp_dir) / "fornecedores.xlsx"
            fornecedores.to_excel(fornecedores_file, index=False)
            with patch.object(app, "FORNECEDORES_FILE", fornecedores_file):
                loaded = app.load_fornecedores()

        self.assertEqual(loaded.iloc[0]["fornecedor"], "JAMAICA")
        self.assertEqual(loaded.iloc[0]["email"], "jamaica@example.com")


class ConferenciaCotacaoTests(unittest.TestCase):
    autobras_pdf = Path(sys.argv[0]).parent / "__missing_autobras_quote.pdf"
    tuba_pdf = Path(sys.argv[0]).parent / "__missing_tuba_quote.pdf"
    tsa_pdf = Path(sys.argv[0]).parent / "__missing_tsa_quote.pdf"

    @staticmethod
    def sample_pdf_from_env(env_name: str) -> Path | None:
        value = os.environ.get(env_name, "").strip()
        if not value:
            return None
        path = Path(value)
        return path if path.exists() else None

    def test_extract_quote_items_from_autobras_pdf(self) -> None:
        pdf_path = self.sample_pdf_from_env("AUTOBRAS_QUOTE_PDF")
        if pdf_path is None:
            self.skipTest("Defina AUTOBRAS_QUOTE_PDF para testar o parser com PDF real.")

        items = app.extract_quote_items_from_pdf(pdf_path, "AUTOBRAS")

        first = items.iloc[0]
        self.assertEqual(first["codigo"], "11041")
        self.assertEqual(first["quantidade_cotada"], 1)
        self.assertAlmostEqual(first["preco_cotado"], 29.48, places=2)

    def test_extract_quote_items_from_tuba_pdf(self) -> None:
        pdf_path = self.sample_pdf_from_env("TUBA_QUOTE_PDF")
        if pdf_path is None:
            self.skipTest("Defina TUBA_QUOTE_PDF para testar o parser com PDF real.")

        items = app.extract_quote_items_from_pdf(pdf_path, "TUBA")

        first = items.iloc[0]
        self.assertEqual(first["codigo"], "5032")
        self.assertEqual(first["quantidade_cotada"], 5)
        self.assertAlmostEqual(first["preco_cotado"], 24.76, places=2)

    def test_extract_quote_items_from_tsa_pdf(self) -> None:
        pdf_path = self.sample_pdf_from_env("TSA_QUOTE_PDF")
        if pdf_path is None:
            self.skipTest("Defina TSA_QUOTE_PDF para testar o parser com PDF real.")

        items = app.extract_quote_items_from_pdf(pdf_path, "TSA")

        first = items.iloc[0]
        self.assertEqual(first["codigo"], "T-010007")
        self.assertEqual(first["quantidade_cotada"], 2)
        self.assertAlmostEqual(first["preco_cotado"], 42.58, places=2)

    def test_extract_quote_items_from_gonel_pdf(self) -> None:
        pdf_path = self.sample_pdf_from_env("GONEL_QUOTE_PDF")
        if pdf_path is None:
            self.skipTest("Defina GONEL_QUOTE_PDF para testar o parser com PDF real.")

        items = app.extract_quote_items_from_pdf(pdf_path, "GONEL")

        first = items.iloc[0]
        self.assertEqual(first["codigo"], "G-1001")
        self.assertEqual(first["quantidade_cotada"], 3)
        self.assertAlmostEqual(first["preco_cotado"], 18.03, places=2)
        broken_code = items[items["codigo"] == "TG-1001"].iloc[0]
        self.assertEqual(broken_code["quantidade_cotada"], 7)
        self.assertAlmostEqual(broken_code["preco_cotado"], 6.87, places=2)

    def test_extract_quote_items_from_wisa_pdf(self) -> None:
        pdf_path = self.sample_pdf_from_env("WISA_QUOTE_PDF")
        if pdf_path is None:
            self.skipTest("Defina WISA_QUOTE_PDF para testar o parser com PDF real.")

        items = app.extract_quote_items_from_pdf(pdf_path, "WISA")

        first = items.iloc[0]
        self.assertEqual(first["codigo"], "7001")
        self.assertEqual(first["quantidade_cotada"], 1)
        self.assertAlmostEqual(first["preco_cotado"], 10.10, places=2)
        item_7032 = items[items["codigo"] == "7032"].iloc[0]
        self.assertEqual(item_7032["quantidade_cotada"], 50)
        self.assertAlmostEqual(item_7032["preco_cotado"], 23.09, places=2)

    def test_extract_quote_items_raises_clear_error_for_unsupported_format(self) -> None:
        with patch.object(app, "read_pdf_text", return_value="unknown quote format"):
            with self.assertRaisesRegex(ValueError, "Formato de cotacao nao suportado"):
                app.extract_quote_items_from_pdf(Path("unknown.pdf"), "UNKNOWN")

    def test_read_sent_order_items_uses_catalogo_and_qnt(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            order_file = Path(temp_dir) / "pedido.xlsx"
            pd.DataFrame([{"Catálogo": "F-11041", "Qnt": 1}]).to_excel(order_file, index=False)

            items = app.read_sent_order_items(order_file)

        first = items.iloc[0]
        self.assertEqual(first["codigo"], "F-11041")
        self.assertEqual(first["quantidade_pedida"], 1)
        self.assertEqual(first["codigo_normalizado"], "11041")

    def test_read_reference_prices_accepts_compra_columns(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            autobras_file = Path(temp_dir) / "autobras.xlsx"
            tuba_file = Path(temp_dir) / "tuba.xlsx"
            tsa_file = Path(temp_dir) / "tsa.xlsx"
            pd.DataFrame(
                [{"CODIGO": "11041", "CODIGO PIKI": "F-11041", "COMPRA A VISTA": 29.48}]
            ).to_excel(autobras_file, sheet_name="Precificação", index=False)
            pd.DataFrame(
                [{"CODIGO TUBA": "5032", "CODIGO PIKI": "TB5032", "compra a prazo": 24.76}]
            ).to_excel(tuba_file, sheet_name="Precificação", index=False)
            pd.DataFrame(
                [{"CÓDIGO": "T-010007", "COMPRA": 42.58}]
            ).to_excel(tsa_file, sheet_name="Precificação", index=False)

            autobras = app.read_reference_prices(autobras_file)
            tuba = app.read_reference_prices(tuba_file)
            tsa = app.read_reference_prices(tsa_file)

        self.assertEqual(autobras["11041"], 29.48)
        self.assertEqual(tuba["5032"], 24.76)
        self.assertEqual(tsa["010007"], 42.58)

    def test_find_reference_base_folder_uses_environment_variable(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch.dict(os.environ, {"PRICE_REFERENCE_DIR": temp_dir}):
                self.assertEqual(app.find_reference_base_folder(), Path(temp_dir))

    def test_find_reference_file_discovers_supplier_workbook_by_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            reference_dir = Path(temp_dir)
            reference_file = reference_dir / "GONEL.xlsx"
            reference_file.write_bytes(b"placeholder")

            with patch.dict(os.environ, {"PRICE_REFERENCE_DIR": str(reference_dir)}):
                self.assertEqual(app.find_reference_file("GONEL"), reference_file)

    def test_find_reference_file_keeps_explicit_reference_name_compatibility(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            reference_dir = Path(temp_dir)
            reference_file = reference_dir / "Autobras.xlsx"
            reference_file.write_bytes(b"placeholder")

            with patch.dict(os.environ, {"PRICE_REFERENCE_DIR": str(reference_dir)}):
                self.assertEqual(app.find_reference_file("AUTOBRAS"), reference_file)

    def test_conference_suppliers_come_from_fornecedores_dataframe(self) -> None:
        fornecedores = pd.DataFrame(
            [
                {"fornecedor": "AUTOBRAS", "email": "autobras@example.com"},
                {"fornecedor": "GONEL", "email": "gonel@example.com"},
                {"fornecedor": "WISA", "email": "wisa@example.com"},
                {"fornecedor": "", "email": "empty@example.com"},
            ]
        )

        self.assertEqual(app.conference_supplier_names(fornecedores), ["AUTOBRAS", "GONEL", "WISA"])

    def test_build_conference_report_outputs_required_columns_and_statuses(self) -> None:
        order_items = pd.DataFrame(
            [
                {"codigo": "F-11041", "codigo_normalizado": "11041", "quantidade_pedida": 1},
                {"codigo": "TB5032", "codigo_normalizado": "5032", "quantidade_pedida": 5},
                {"codigo": "T-010007", "codigo_normalizado": "010007", "quantidade_pedida": 2},
                {"codigo": "SEMREF", "codigo_normalizado": "999999", "quantidade_pedida": 1},
                {"codigo": "SEMCOT", "codigo_normalizado": "888888", "quantidade_pedida": 1},
            ]
        )
        quote_items = pd.DataFrame(
            [
                {"codigo": "11041", "codigo_normalizado": "11041", "quantidade_cotada": 1, "preco_cotado": 100.50},
                {"codigo": "5032", "codigo_normalizado": "5032", "quantidade_cotada": 5, "preco_cotado": 120.00},
                {"codigo": "T-010007", "codigo_normalizado": "010007", "quantidade_cotada": 1, "preco_cotado": 50.00},
                {"codigo": "999999", "codigo_normalizado": "999999", "quantidade_cotada": 1, "preco_cotado": 10.00},
            ]
        )
        reference_prices = {"11041": 100.00, "5032": 100.00, "010007": 100.00, "888888": 10.00}

        report = app.build_conference_report(order_items, quote_items, reference_prices)

        self.assertEqual(list(report.columns), app.CONFERENCE_REPORT_COLUMNS)
        self.assertEqual(report.loc[report["codigo"] == "F-11041", "status"].iloc[0], "OK")
        self.assertEqual(report.loc[report["codigo"] == "TB5032", "status"].iloc[0], "PRECO_DIVERGENTE")
        self.assertEqual(
            report.loc[report["codigo"] == "T-010007", "status"].iloc[0],
            "PRECO_E_QUANTIDADE_DIVERGENTES",
        )
        self.assertEqual(
            report.loc[report["codigo"] == "SEMREF", "status"].iloc[0],
            "NAO_ENCONTRADO_REFERENCIA",
        )
        self.assertEqual(
            report.loc[report["codigo"] == "SEMCOT", "status"].iloc[0],
            "NAO_ENCONTRADO_COTACAO",
        )

    def test_conference_report_to_excel_bytes_keeps_required_columns(self) -> None:
        report = pd.DataFrame(
            [
                {
                    "codigo": "F-11041",
                    "quantidade_pedida": 1,
                    "quantidade_cotada": 1,
                    "preco_cotado": 100.5,
                    "preco_referencia": 100,
                    "diferenca_percentual": 0.5,
                    "status": "OK",
                }
            ],
            columns=app.CONFERENCE_REPORT_COLUMNS,
        )

        exported = app.conference_report_to_excel_bytes(report)
        loaded = pd.read_excel(BytesIO(exported), engine="openpyxl")

        self.assertEqual(list(loaded.columns), app.CONFERENCE_REPORT_COLUMNS)
        self.assertEqual(loaded.iloc[0]["status"], "OK")


class GmailAuthTests(unittest.TestCase):
    def test_get_gmail_service_reauthenticates_when_refresh_token_is_revoked(self) -> None:
        class ExpiredCredentials:
            expired = True
            refresh_token = "refresh-token"
            valid = False

            def refresh(self, _request):
                raise app.RefreshError("invalid_grant: Token has been expired or revoked.")

        class NewCredentials:
            expired = False
            refresh_token = "new-refresh-token"
            valid = True

            def to_json(self):
                return '{"token": "new"}'

        class FakeFlow:
            def run_local_server(self, port=0):
                return NewCredentials()

        with tempfile.TemporaryDirectory() as temp_dir:
            token_file = Path(temp_dir) / "token.json"
            credentials_file = Path(temp_dir) / "credentials.json"
            token_file.write_text('{"token": "old"}', encoding="utf-8")
            credentials_file.write_text("{}", encoding="utf-8")

            with (
                patch.object(app, "TOKEN_FILE", token_file),
                patch.object(app, "CREDENTIALS_FILE", credentials_file),
                patch.object(app.Credentials, "from_authorized_user_file", return_value=ExpiredCredentials()),
                patch.object(app.InstalledAppFlow, "from_client_secrets_file", return_value=FakeFlow()),
            ):
                service = app.get_gmail_service()
                token_contents = token_file.read_text(encoding="utf-8")

        self.assertEqual(service.credentials.__class__, NewCredentials)
        self.assertEqual(service.verify, str(app.CA_BUNDLE_PATH))
        self.assertIsInstance(service.adapters["https://"], app.GoogleSSLAdapter)
        self.assertEqual(token_contents, '{"token": "new"}')


if __name__ == "__main__":
    unittest.main()
