import argparse
import json
import base64
import logging
import os
import socket
import ssl
import sys
import time
import textwrap
import unicodedata
from collections import defaultdict
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlparse

import requests
import win32event
import win32print
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth
from websocket import WebSocketConnectionClosedException, WebSocketTimeoutException, create_connection


SCRIPT_VERSION = "1.0.2"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
LOG_FILE_PATH = os.path.join(BASE_DIR, "print_server.log")
PRINT_CACHE_PATH = os.path.join(BASE_DIR, "printed_cache.json")
SERVICE_STOP_EVENT = None
TICKET_WIDTH = 40
QTY_COL_WIDTH = 8
DESC_COL_WIDTH = TICKET_WIDTH - QTY_COL_WIDTH


def load_environment() -> None:
    if os.path.exists(ENV_PATH):
        load_dotenv(ENV_PATH)
    else:
        raise FileNotFoundError(f"No se encontro el archivo .env en {ENV_PATH}")


def setup_logging() -> None:
    root_logger = logging.getLogger()
    if root_logger.handlers:
        for handler in list(root_logger.handlers):
            root_logger.removeHandler(handler)
            handler.close()

    backup_count = int(os.getenv("LOG_BACKUP_COUNT", "7"))
    handler = TimedRotatingFileHandler(
        LOG_FILE_PATH,
        when="midnight",
        interval=1,
        backupCount=backup_count,
        encoding="utf-8",
    )
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    handler.setFormatter(formatter)

    root_logger.addHandler(handler)
    root_logger.setLevel(logging.INFO)
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)


def init_service_stop_event(event: Any) -> None:
    global SERVICE_STOP_EVENT
    SERVICE_STOP_EVENT = event


def env_required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"La variable de entorno {name} es obligatoria")
    return value


def parse_json_env(name: str, default: Any) -> Any:
    raw_value = os.getenv(name, "").strip()
    if not raw_value:
        return default
    try:
        return json.loads(raw_value)
    except json.JSONDecodeError as exc:
        raise ValueError(f"La variable {name} debe contener JSON valido") from exc


def parse_hex_commands(raw_value: str, env_name: str) -> bytes:
    value = raw_value.strip()
    if not value:
        return b""

    command_parts = [part.strip().replace(" ", "") for part in value.split(",") if part.strip()]
    payload = b""
    for part in command_parts:
        try:
            payload += bytes.fromhex(part)
        except ValueError as exc:
            raise ValueError(f"{env_name} debe contener hexadecimal valido") from exc
    return payload


def normalize_center(value: Optional[str]) -> str:
    return (value or "SIN_CENTRO").strip().upper()


def has_printable_center(value: Optional[str]) -> bool:
    return normalize_center(value) != "SIN_CENTRO"


def list_installed_printers() -> List[str]:
    printer_flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(printer_flags)
    names = sorted({printer[2] for printer in printers if len(printer) > 2 and printer[2]})
    return names


def parse_printer_centers(raw_value: str) -> List[str]:
    return [normalize_center(item) for item in raw_value.split(",") if item.strip()]


def update_env_value(file_path: str, key: str, value: str) -> None:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"No se encontro el archivo {file_path}")

    with open(file_path, "r", encoding="utf-8") as env_file:
        lines = env_file.readlines()

    replacement = f"{key}={value}\n"
    key_prefix = f"{key}="
    updated = False
    for index, line in enumerate(lines):
        if line.startswith(key_prefix):
            lines[index] = replacement
            updated = True
            break

    if not updated:
        if lines and not lines[-1].endswith("\n"):
            lines[-1] = lines[-1] + "\n"
        lines.append(replacement)

    with open(file_path, "w", encoding="utf-8") as env_file:
        env_file.writelines(lines)


def load_print_cache() -> Dict[str, float]:
    if not os.path.exists(PRINT_CACHE_PATH):
        return {}
    try:
        with open(PRINT_CACHE_PATH, "r", encoding="utf-8") as cache_file:
            payload = json.load(cache_file)
        if isinstance(payload, dict):
            return {str(key): float(value) for key, value in payload.items()}
    except (OSError, ValueError, TypeError):
        logging.warning("No se pudo leer printed_cache.json. Se recreara.")
    return {}


def save_print_cache(cache: Dict[str, float]) -> None:
    with open(PRINT_CACHE_PATH, "w", encoding="utf-8") as cache_file:
        json.dump(cache, cache_file, indent=2, sort_keys=True)


def cache_detail_printed(cache: Dict[str, float], detail_id: Any) -> None:
    cache[str(detail_id)] = time.time()


def uncache_detail_printed(cache: Dict[str, float], detail_id: Any) -> None:
    cache.pop(str(detail_id), None)


def is_detail_cached(cache: Dict[str, float], detail_id: Any) -> bool:
    return str(detail_id) in cache


def generate_printer_map(centers: List[str]) -> Dict[str, str]:
    printers = list_installed_printers()
    if not printers:
        raise RuntimeError("No se encontraron impresoras instaladas en Windows")

    default_printer = printers[0]
    return {center: default_printer for center in centers}


def handle_cli(argv: List[str]) -> bool:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--list-printers", action="store_true", help="Lista las impresoras instaladas en Windows")
    parser.add_argument(
        "--sync-printer-map",
        action="store_true",
        help="Actualiza PRINTER_MAP_JSON en .env usando las impresoras instaladas",
    )
    parser.add_argument(
        "--centers",
        default="COCINA,BARRA,PARRILLAS",
        help="Centros separados por coma para generar PRINTER_MAP_JSON",
    )
    args = parser.parse_args(argv)

    if args.list_printers:
        printers = list_installed_printers()
        if not printers:
            print("No se encontraron impresoras instaladas.")
            return True
        print("Impresoras instaladas:")
        for printer_name in printers:
            print(f"- {printer_name}")
        return True

    if args.sync_printer_map:
        centers = parse_printer_centers(args.centers)
        if not centers:
            raise ValueError("Debes indicar al menos un centro en --centers")
        printer_map = generate_printer_map(centers)
        json_value = json.dumps(printer_map, ensure_ascii=False, separators=(",", ":"))
        update_env_value(ENV_PATH, "PRINTER_MAP_JSON", json_value)
        print("PRINTER_MAP_JSON actualizado en .env")
        print(json_value)
        return True

    return False


class Config:
    def __init__(self) -> None:
        self.orders_url = env_required("API_ORDERS_URL")
        self.company = os.getenv("Company", "").strip()
        self.order_status = os.getenv("ORDER_STATUS", "Registrado").strip()
        self.auth_mode = os.getenv("API_AUTH_MODE", "basic").strip().lower()
        self.username = os.getenv("API_USERNAME", "").strip()
        self.password = os.getenv("API_PASSWORD", "").strip()
        self.token = os.getenv("API_TOKEN", "").strip()
        self.token_header = os.getenv("API_TOKEN_HEADER", "Authorization").strip()
        self.token_prefix = os.getenv("API_TOKEN_PREFIX", "Bearer").strip()
        self.timeout = int(os.getenv("API_TIMEOUT_SECONDS", "30"))
        self.verify_tls = os.getenv("API_VERIFY_TLS", "true").strip().lower() not in {"0", "false", "no"}
        self.poll_seconds = int(os.getenv("POLL_SECONDS", "10"))
        self.idle_sleep_seconds = int(os.getenv("IDLE_SLEEP_SECONDS", str(self.poll_seconds)))
        self.error_sleep_seconds = int(os.getenv("ERROR_SLEEP_SECONDS", "30"))
        self.websocket_url = self._build_websocket_url()
        self.websocket_verify_tls = os.getenv("WS_VERIFY_TLS", str(self.verify_tls)).strip().lower() not in {
            "0",
            "false",
            "no",
        }
        self.websocket_ping_interval_seconds = int(os.getenv("WS_PING_INTERVAL_SECONDS", "20"))
        self.websocket_ping_timeout_seconds = int(os.getenv("WS_PING_TIMEOUT_SECONDS", "20"))
        self.websocket_connect_timeout_seconds = int(
            os.getenv("WS_CONNECT_TIMEOUT_SECONDS", str(self.timeout))
        )
        self.websocket_event_types = {
            item.strip()
            for item in os.getenv("WS_EVENT_TYPES", "order.created,order.updated").split(",")
            if item.strip()
        }
        self.query_params = parse_json_env("API_QUERY_PARAMS_JSON", {})
        self.extra_headers = parse_json_env("API_HEADERS_JSON", {})
        self.printer_map = {
            normalize_center(key): value
            for key, value in parse_json_env("PRINTER_MAP_JSON", {}).items()
            if str(value).strip()
        }
        self.print_copies = int(os.getenv("PRINT_COPIES", "1"))
        self.print_encoding = os.getenv("PRINT_ENCODING", "cp850")
        self.print_codepage_command_hex = os.getenv("PRINT_CODEPAGE_COMMAND_HEX", "1B7402").strip()
        self.document_title = os.getenv("PRINT_DOCUMENT_TITLE", "TicketProduccion")
        self.cut_lines = int(os.getenv("PRINT_FEED_LINES", "4"))
        self.highlight_command_hex = os.getenv("PRINT_HIGHLIGHT_COMMAND_HEX", "1D2111").strip()
        self.highlight_reset_command_hex = os.getenv("PRINT_HIGHLIGHT_RESET_COMMAND_HEX", "1D2100").strip()
        self.cut_enabled = os.getenv("PRINT_CUT_ENABLED", "true").strip().lower() not in {"0", "false", "no"}
        self.cut_command_hex = os.getenv("PRINT_CUT_COMMAND_HEX", "1D5641")
        self.pre_cut_feed_lines = int(os.getenv("PRINT_PRE_CUT_FEED_LINES", "3"))
        self.after_job_sleep_ms = int(os.getenv("PRINT_AFTER_JOB_SLEEP_MS", "300"))
        self.combine_ticket_and_cut = os.getenv("PRINT_COMBINE_TICKET_AND_CUT", "true").strip().lower() not in {
            "0",
            "false",
            "no",
        }
        self.beep_enabled = os.getenv("PRINT_BEEP_ENABLED", "false").strip().lower() not in {"0", "false", "no"}
        self.beep_command_hex = os.getenv("PRINT_BEEP_COMMAND_HEX", "").strip()
        self.beep_position = os.getenv("PRINT_BEEP_POSITION", "before_cut").strip().lower()
        self.reprint_when_not_confirmed = os.getenv("REPRINT_WHEN_NOT_CONFIRMED", "false").strip().lower() not in {
            "0",
            "false",
            "no",
        }

        self.printed_url_template = env_required("API_PRINTED_URL_TEMPLATE")
        self.printed_method = os.getenv("API_PRINTED_METHOD", "PATCH").strip().upper()

        if self.auth_mode == "basic" and (not self.username or not self.password):
            raise ValueError("API_USERNAME y API_PASSWORD son obligatorias para API_AUTH_MODE=basic")
        if self.auth_mode in {"bearer", "token"} and not self.token:
            raise ValueError("API_TOKEN es obligatoria para API_AUTH_MODE=bearer/token")
        if not self.printer_map:
            raise ValueError("PRINTER_MAP_JSON es obligatorio y debe mapear centros a impresoras")
        if not self.websocket_url:
            raise ValueError("WS_ORDERS_URL es obligatoria o debe poder derivarse desde API_ORDERS_URL")

    def build_orders_query_params(self) -> Dict[str, Any]:
        params = dict(self.query_params)
        if self.company:
            params["Company"] = self.company
        if self.order_status:
            params["Status"] = self.order_status
        return params

    def _build_websocket_url(self) -> str:
        explicit_url = os.getenv("WS_ORDERS_URL", "").strip()
        if explicit_url:
            return explicit_url

        restaurant_id = os.getenv("WS_RESTAURANT_ID", self.company).strip()
        if not restaurant_id:
            return ""

        parsed_url = urlparse(self.orders_url)
        if not parsed_url.scheme or not parsed_url.netloc:
            return ""

        websocket_scheme = "wss" if parsed_url.scheme == "https" else "ws"
        return f"{websocket_scheme}://{parsed_url.netloc}/ws/restaurants/{restaurant_id}/tables/"


class OrderApiClient:
    def __init__(self, config: Config) -> None:
        self.config = config
        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})
        if config.extra_headers:
            self.session.headers.update(config.extra_headers)
        if config.auth_mode == "basic":
            self.session.auth = HTTPBasicAuth(config.username, config.password)
        elif config.auth_mode in {"bearer", "token"}:
            token_value = config.token
            if config.token_prefix:
                token_value = f"{config.token_prefix} {token_value}"
            self.session.headers[config.token_header] = token_value

    def get_pending_orders(self) -> List[Dict[str, Any]]:
        response = self.session.get(
            self.config.orders_url,
            params=self.config.build_orders_query_params(),
            timeout=self.config.timeout,
            verify=self.config.verify_tls,
        )
        response.raise_for_status()
        payload = response.json()
        orders = self._normalize_orders_payload(payload)
        return [order for order in orders if self._has_pending_details(order)]

    def mark_order_details_as_printed(self, order: Dict[str, Any], details: List[Dict[str, Any]]) -> None:
        if not details:
            return

        substitutions = self._build_substitutions(order, details[0])
        url = self.config.printed_url_template.format(**substitutions)
        payload = self._build_printed_payload(order, details)
        response = self.session.request(
            self.config.printed_method,
            url,
            json=payload,
            timeout=self.config.timeout,
            verify=self.config.verify_tls,
        )
        response.raise_for_status()

    def build_websocket_headers(self) -> List[str]:
        headers: List[str] = []
        for key, value in self.session.headers.items():
            headers.append(f"{key}: {value}")

        if self.config.auth_mode == "basic":
            credentials = f"{self.config.username}:{self.config.password}".encode("utf-8")
            encoded_credentials = base64.b64encode(credentials).decode("ascii")
            headers.append(f"Authorization: Basic {encoded_credentials}")

        return headers

    @staticmethod
    def _normalize_orders_payload(payload: Any) -> List[Dict[str, Any]]:
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            for key in ("results", "data", "items", "orders"):
                value = payload.get(key)
                if isinstance(value, list):
                    return value
            if "id" in payload and "Details" in payload:
                return [payload]
        raise ValueError("La respuesta de la API no tiene un formato de ordenes soportado")

    @staticmethod
    def _has_pending_details(order: Dict[str, Any]) -> bool:
        details = order.get("Details") or []
        return any(not bool(detail.get("Printed")) for detail in details)

    @staticmethod
    def _build_printed_payload(order: Dict[str, Any], details_to_mark: List[Dict[str, Any]]) -> Dict[str, Any]:
        details_to_mark_ids = {detail.get("id") for detail in details_to_mark}
        payload_details: List[Dict[str, Any]] = []
        for detail in order.get("Details") or []:
            detail_payload = dict(detail)
            if detail.get("id") in details_to_mark_ids:
                detail_payload["Printed"] = True
            payload_details.append(detail_payload)
        return {"Details": payload_details}

    @staticmethod
    def _build_substitutions(order: Dict[str, Any], detail: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "order_id": order.get("id"),
            "detail_id": detail.get("id"),
            "item": detail.get("Item"),
            "product_id": detail.get("ProductId"),
            "production_center": normalize_center(detail.get("ProductionCenter")),
        }


class TicketPrinter:
    def __init__(self, config: Config) -> None:
        self.config = config

    def print_order_group(self, order: Dict[str, Any], center: str, details: Iterable[Dict[str, Any]]) -> None:
        printer_name = self.config.printer_map.get(center)
        if not printer_name:
            raise ValueError(f"No hay impresora configurada para el centro '{center}'")

        raw_bytes = self._build_document_bytes(order, center, list(details))
        beep_bytes = self._build_beep_bytes()
        cut_bytes = self._build_cut_bytes()
        payload_bytes = self._build_payload_bytes(raw_bytes, beep_bytes, cut_bytes)

        logging.info("Enviando pedido %s al centro %s en impresora '%s'", order.get("id"), center, printer_name)
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            for copy_number in range(self.config.print_copies):
                job_id = win32print.StartDocPrinter(printer_handle, 1, (self.config.document_title, None, "RAW"))
                try:
                    win32print.StartPagePrinter(printer_handle)
                    win32print.WritePrinter(printer_handle, payload_bytes)
                    if cut_bytes and not self.config.combine_ticket_and_cut:
                        win32print.WritePrinter(printer_handle, cut_bytes)
                    win32print.EndPagePrinter(printer_handle)
                    logging.info(
                        "Trabajo de impresion generado. Pedido=%s Centro=%s Copia=%s JobId=%s",
                        order.get("id"),
                        center,
                        copy_number + 1,
                        job_id,
                    )
                finally:
                    win32print.EndDocPrinter(printer_handle)
                    if self.config.after_job_sleep_ms > 0:
                        time.sleep(self.config.after_job_sleep_ms / 1000)
        finally:
            win32print.ClosePrinter(printer_handle)

    def _build_payload_bytes(self, raw_bytes: bytes, beep_bytes: bytes, cut_bytes: bytes) -> bytes:
        if not self.config.combine_ticket_and_cut:
            return raw_bytes

        if self.config.beep_position == "before_ticket":
            return beep_bytes + raw_bytes + cut_bytes
        if self.config.beep_position == "after_ticket":
            return raw_bytes + beep_bytes + cut_bytes
        if self.config.beep_position == "after_cut":
            return raw_bytes + cut_bytes + beep_bytes
        return raw_bytes + beep_bytes + cut_bytes

    def _build_beep_bytes(self) -> bytes:
        if not self.config.beep_enabled:
            return b""
        if not self.config.beep_command_hex:
            raise ValueError("PRINT_BEEP_COMMAND_HEX es obligatorio cuando PRINT_BEEP_ENABLED=true")
        return parse_hex_commands(self.config.beep_command_hex, "PRINT_BEEP_COMMAND_HEX")

    def _build_cut_bytes(self) -> bytes:
        if not self.config.cut_enabled:
            return b""
        pre_cut_feed = b"\n" * max(self.config.pre_cut_feed_lines, 0)
        return pre_cut_feed + parse_hex_commands(self.config.cut_command_hex, "PRINT_CUT_COMMAND_HEX")

    def _build_document_bytes(self, order: Dict[str, Any], center: str, details: List[Dict[str, Any]]) -> bytes:
        order_id = order.get("id", "")
        order_number = f"{order.get('OrderSerie', '')}-{order.get('OrderNumber', '')}".strip("-")
        customer = order.get("ExternalPerson", "")
        cashier = order.get("Cashier") or order.get("InternalPerson") or ""
        workspace = order.get("WorkSpace") or ""
        space = order.get("Space") or ""
        observations = order.get("Observations") or ""
        created_at = order.get("Hour") or order.get("Created") or ""
        header_lines = [
            "=" * TICKET_WIDTH,
            f"CENTRO DE PRODUCCIÓN: {center}".center(TICKET_WIDTH),
            "=" * TICKET_WIDTH,
            f"NUMERO: {order_number}",
            f"FECHA Y HORA: {self._format_datetime(created_at)}",
        ]
        if cashier:
            header_lines.append(f"ASESOR: {cashier}")
        item_lines = []
        item_lines.append("-" * TICKET_WIDTH)
        item_lines.append(self._format_table_header("CANTIDAD", "DESCRIPCIÓN"))
        item_lines.append("-" * TICKET_WIDTH)
        for detail in details:
            quantity = self._format_quantity(detail.get("Quantity"))
            product = (detail.get("Product") or "").strip()
            item_lines.extend(self._format_detail_rows(quantity, product))
            detail_obs = (detail.get("Observations") or "").strip()
            if detail_obs:
                item_lines.extend(self._format_detail_rows("", f"Obs: {detail_obs}"))

        footer_lines = []
        if observations:
            footer_lines.append(f"OBSERVACIONES: {observations}"[:120])
        footer_lines.append(f"Sayri Versión {SCRIPT_VERSION}".center(TICKET_WIDTH))

        lines: List[Any] = list(header_lines)
        lines.append("-" * TICKET_WIDTH)
        if workspace:
            lines.append(self._highlight_line(f"SALÓN: {workspace}"))
        if space:
            lines.append(self._highlight_line(f"ESPACIO: {space}"))
        lines.extend(item_lines)
        lines.append("-" * TICKET_WIDTH)
        lines.extend(footer_lines)
        lines.extend([""] * self.config.cut_lines)
        return self._build_codepage_bytes() + self._encode_mixed_lines(lines)

    def _build_codepage_bytes(self) -> bytes:
        if not self.config.print_codepage_command_hex:
            return b""
        return parse_hex_commands(self.config.print_codepage_command_hex, "PRINT_CODEPAGE_COMMAND_HEX")

    def _highlight_line(self, text: str) -> bytes:
        highlight_on = parse_hex_commands(self.config.highlight_command_hex, "PRINT_HIGHLIGHT_COMMAND_HEX")
        bold_on = b"\x1b\x45\x01"
        bold_off = b"\x1b\x45\x00"
        highlight_off = parse_hex_commands(
            self.config.highlight_reset_command_hex, "PRINT_HIGHLIGHT_RESET_COMMAND_HEX"
        )
        encoded_text = self._encode_text(text)
        return highlight_on + bold_on + encoded_text + bold_off + highlight_off

    def _encode_mixed_lines(self, lines: List[Any]) -> bytes:
        encoded_lines: List[bytes] = []
        for line in lines:
            if isinstance(line, bytes):
                encoded_lines.append(line)
            else:
                encoded_lines.append(self._encode_text(str(line)))
        return b"\n".join(encoded_lines)

    def _encode_text(self, text: str) -> bytes:
        normalized_text = unicodedata.normalize("NFC", text)
        return normalized_text.encode(self.config.print_encoding, errors="replace")

    def _format_detail_rows(self, quantity: str, description: str) -> List[str]:
        wrapped_description = textwrap.wrap(description, width=DESC_COL_WIDTH) or [""]
        rows: List[str] = []
        for index, chunk in enumerate(wrapped_description):
            qty_value = quantity if index == 0 else ""
            rows.append(self._format_table_row(qty_value, chunk))
        return rows

    @staticmethod
    def _format_table_row(quantity: str, description: str) -> str:
        qty_text = str(quantity or "")[:QTY_COL_WIDTH].center(QTY_COL_WIDTH)
        desc_text = str(description or "")[:DESC_COL_WIDTH].ljust(DESC_COL_WIDTH)
        return f"{qty_text}{desc_text}"

    @staticmethod
    def _format_table_header(quantity: str, description: str) -> str:
        qty_text = str(quantity or "")[:QTY_COL_WIDTH].center(QTY_COL_WIDTH)
        desc_text = str(description or "")[:DESC_COL_WIDTH].center(DESC_COL_WIDTH)
        return f"{qty_text}{desc_text}"

    @staticmethod
    def _format_datetime(value: str) -> str:
        if not value:
            return ""
        try:
            return datetime.fromisoformat(value).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            return value

    @staticmethod
    def _format_quantity(value: Any) -> str:
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return str(value or "")
        if numeric.is_integer():
            return str(int(numeric))
        return f"{numeric:.2f}".rstrip("0").rstrip(".")


def group_pending_details(order: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
    grouped: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for detail in order.get("Details") or []:
        if bool(detail.get("Printed")):
            continue
        center = normalize_center(detail.get("ProductionCenter"))
        if not has_printable_center(center):
            continue
        grouped[center].append(detail)
    return grouped


def split_order_details(
    order: Dict[str, Any], print_cache: Dict[str, float], config: Config
) -> Tuple[Dict[str, List[Dict[str, Any]]], List[Dict[str, Any]]]:
    grouped_to_print: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    pending_confirmation: List[Dict[str, Any]] = []

    for detail in order.get("Details") or []:
        if bool(detail.get("Printed")):
            continue

        center = normalize_center(detail.get("ProductionCenter"))
        if not has_printable_center(center):
            logging.info(
                "Omitiendo detalle sin centro de produccion valido. Pedido=%s Detalle=%s Producto='%s' Centro='%s'",
                order.get("id"),
                detail.get("id"),
                detail.get("Product"),
                center,
            )
            pending_confirmation.append(detail)
            continue

        if not config.reprint_when_not_confirmed and is_detail_cached(print_cache, detail.get("id")):
            pending_confirmation.append(detail)
            continue

        grouped_to_print[center].append(detail)

    return grouped_to_print, pending_confirmation


def log_order_detail_centers(order: Dict[str, Any], config: Config) -> None:
    details = order.get("Details") or []
    if not details:
        logging.info("Pedido %s no contiene detalles.", order.get("id"))
        return

    summary: Dict[str, int] = defaultdict(int)
    for detail in details:
        center = normalize_center(detail.get("ProductionCenter"))
        summary[center] += 1
        printer_name = config.printer_map.get(center)
        logging.info(
            "Pedido=%s Detalle=%s Producto='%s' Centro='%s' Impresora='%s' PrintedApi=%s",
            order.get("id"),
            detail.get("id"),
            detail.get("Product"),
            center,
            printer_name or "SIN_MAPEO",
            detail.get("Printed"),
        )

    logging.info("Resumen de centros del pedido %s: %s", order.get("id"), dict(summary))


def try_mark_order_details(
    client: OrderApiClient, order: Dict[str, Any], details: List[Dict[str, Any]], print_cache: Dict[str, float]
) -> int:
    if not details:
        return 0

    try:
        client.mark_order_details_as_printed(order, details)
        for detail in details:
            uncache_detail_printed(print_cache, detail.get("id"))
            detail["Printed"] = True
            logging.info(
                "Detalle marcado como impreso. Pedido=%s Detalle=%s Centro=%s",
                order.get("id"),
                detail.get("id"),
                normalize_center(detail.get("ProductionCenter")),
            )
        return len(details)
    except Exception as exc:
        detail_ids = [detail.get("id") for detail in details]
        logging.exception(
            "No se pudo actualizar el estado de impresion del pedido %s para detalles %s: %s",
            order.get("id"),
            detail_ids,
            exc,
        )
        return 0


def wait_for_next_cycle(seconds: int) -> bool:
    if SERVICE_STOP_EVENT:
        wait_result = win32event.WaitForSingleObject(SERVICE_STOP_EVENT, max(seconds, 1) * 1000)
        return wait_result == win32event.WAIT_OBJECT_0
    time.sleep(max(seconds, 1))
    return False


def process_order(
    client: OrderApiClient, printer: TicketPrinter, print_cache: Dict[str, float], config: Config, order: Dict[str, Any]
) -> Tuple[int, int]:
    printed_count = 0
    confirmed_count = 0
    order_id = order.get("id")
    log_order_detail_centers(order, config)
    grouped, pending_confirmation = split_order_details(order, print_cache, config)
    details_to_confirm: List[Dict[str, Any]] = list(pending_confirmation)
    if grouped:
        logging.info("Procesando pedido %s en %s centro(s).", order_id, len(grouped))
    for center, details in grouped.items():
        printer.print_order_group(order, center, details)
        for detail in details:
            cache_detail_printed(print_cache, detail.get("id"))
            printed_count += 1
        details_to_confirm.extend(details)

    confirmed_count += try_mark_order_details(client, order, details_to_confirm, print_cache)

    save_print_cache(print_cache)
    return confirmed_count, printed_count


def should_process_order_event(order: Dict[str, Any], config: Config) -> bool:
    if not order:
        return False
    if config.company and str(order.get("Company", "")) != str(config.company):
        return False
    if config.order_status and str(order.get("Status", "")).strip() != config.order_status:
        return False
    details = order.get("Details") or []
    return any(not bool(detail.get("Printed")) for detail in details)


def extract_order_from_event(payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    order_payload = payload.get("order")
    if isinstance(order_payload, dict):
        return order_payload
    if "id" in payload and "Details" in payload:
        return payload
    return None


def listen_for_order_events(
    client: OrderApiClient, printer: TicketPrinter, print_cache: Dict[str, float], config: Config
) -> None:
    logging.info("Escuchando eventos en %s", config.websocket_url)

    while True:
        websocket = None
        try:
            ssl_options = {"cert_reqs": ssl.CERT_REQUIRED if config.websocket_verify_tls else ssl.CERT_NONE}
            websocket = create_connection(
                config.websocket_url,
                timeout=config.websocket_connect_timeout_seconds,
                header=client.build_websocket_headers(),
                sslopt=ssl_options,
            )
            websocket.settimeout(1)
            last_ping_at = time.monotonic()
            logging.info("Conexion WebSocket establecida correctamente.")

            while True:
                try:
                    raw_message = websocket.recv()
                except WebSocketTimeoutException:
                    now = time.monotonic()
                    if config.websocket_ping_interval_seconds > 0 and (
                        now - last_ping_at
                    ) >= config.websocket_ping_interval_seconds:
                        websocket.ping()
                        last_ping_at = now
                    if SERVICE_STOP_EVENT and wait_for_next_cycle(1):
                        logging.info("Se recibio senal de parada del servicio.")
                        return
                    continue

                if raw_message is None:
                    raise WebSocketConnectionClosedException("La conexion WebSocket fue cerrada")

                payload = json.loads(raw_message)
                event_type = str(payload.get("type", "")).strip()
                if event_type == "connection.accepted":
                    logging.info("Canal suscrito: %s", payload.get("group"))
                    continue
                if event_type not in config.websocket_event_types:
                    continue

                order = extract_order_from_event(payload)
                if not order:
                    logging.warning("Evento %s recibido sin payload de pedido util.", event_type)
                    continue
                if not should_process_order_event(order, config):
                    logging.info(
                        "Evento %s omitido para pedido %s por estado, compania o sin detalles pendientes.",
                        event_type,
                        order.get("id"),
                    )
                    continue

                confirmed_count, printed_count = process_order(client, printer, print_cache, config, order)
                logging.info(
                    "Evento procesado. Tipo=%s Pedido=%s Detalles impresos=%s Detalles confirmados=%s Pendientes locales=%s",
                    event_type,
                    order.get("id"),
                    printed_count,
                    confirmed_count,
                    len(print_cache),
                )
        except (OSError, socket.error, TimeoutError, WebSocketConnectionClosedException, ConnectionError) as exc:
            logging.error("Conexion WebSocket interrumpida: %s", exc, exc_info=True)
        except json.JSONDecodeError as exc:
            logging.error("Se recibio un mensaje WebSocket invalido: %s", exc, exc_info=True)
        except Exception as exc:
            logging.error("Error procesando eventos WebSocket: %s", exc, exc_info=True)
        finally:
            if websocket is not None:
                try:
                    websocket.close()
                except Exception:
                    pass

        logging.info("Reintentando conexion WebSocket en %s segundo(s).", config.error_sleep_seconds)
        if wait_for_next_cycle(config.error_sleep_seconds):
            logging.info("Se recibio senal de parada del servicio.")
            break


def main() -> None:
    load_environment()
    setup_logging()
    config = Config()
    client = OrderApiClient(config)
    printer = TicketPrinter(config)
    print_cache = load_print_cache()

    logging.info("ATIC Peru")
    logging.info("Iniciando print_server.py version %s", SCRIPT_VERSION)
    logging.info("Modo de escucha configurado por WebSocket.")
    listen_for_order_events(client, printer, print_cache, config)


if __name__ == "__main__":
    if handle_cli(sys.argv[1:]):
        sys.exit(0)
    main()
