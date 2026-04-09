import json
import logging
import os
import time
from collections import defaultdict
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from string import Template
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
import win32event
import win32print
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth


SCRIPT_VERSION = "1.0.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
LOG_FILE_PATH = os.path.join(BASE_DIR, "print_server.log")
SERVICE_STOP_EVENT = None


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


def normalize_center(value: Optional[str]) -> str:
    return (value or "SIN_CENTRO").strip().upper()


class Config:
    def __init__(self) -> None:
        self.orders_url = env_required("API_ORDERS_URL")
        self.auth_mode = os.getenv("API_AUTH_MODE", "basic").strip().lower()
        self.username = os.getenv("API_USERNAME", "").strip()
        self.password = os.getenv("API_PASSWORD", "").strip()
        self.token = os.getenv("API_TOKEN", "").strip()
        self.token_header = os.getenv("API_TOKEN_HEADER", "Authorization").strip()
        self.token_prefix = os.getenv("API_TOKEN_PREFIX", "Bearer").strip()
        self.timeout = int(os.getenv("API_TIMEOUT_SECONDS", "30"))
        self.poll_seconds = int(os.getenv("POLL_SECONDS", "10"))
        self.idle_sleep_seconds = int(os.getenv("IDLE_SLEEP_SECONDS", str(self.poll_seconds)))
        self.error_sleep_seconds = int(os.getenv("ERROR_SLEEP_SECONDS", "30"))
        self.query_params = parse_json_env("API_QUERY_PARAMS_JSON", {})
        self.extra_headers = parse_json_env("API_HEADERS_JSON", {})
        self.printer_map = {
            normalize_center(key): value
            for key, value in parse_json_env("PRINTER_MAP_JSON", {}).items()
            if str(value).strip()
        }
        self.print_copies = int(os.getenv("PRINT_COPIES", "1"))
        self.print_encoding = os.getenv("PRINT_ENCODING", "cp850")
        self.document_title = os.getenv("PRINT_DOCUMENT_TITLE", "TicketProduccion")
        self.cut_lines = int(os.getenv("PRINT_FEED_LINES", "4"))
        self.verify_tls = os.getenv("API_VERIFY_TLS", "true").strip().lower() not in {"0", "false", "no"}

        self.printed_url_template = env_required("API_PRINTED_URL_TEMPLATE")
        self.printed_method = os.getenv("API_PRINTED_METHOD", "PATCH").strip().upper()
        self.printed_payload_template = os.getenv("API_PRINTED_PAYLOAD_TEMPLATE", '{"Printed": true}')

        if self.auth_mode == "basic" and (not self.username or not self.password):
            raise ValueError("API_USERNAME y API_PASSWORD son obligatorias para API_AUTH_MODE=basic")
        if self.auth_mode in {"bearer", "token"} and not self.token:
            raise ValueError("API_TOKEN es obligatoria para API_AUTH_MODE=bearer/token")
        if not self.printer_map:
            raise ValueError("PRINTER_MAP_JSON es obligatorio y debe mapear centros a impresoras")


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
            params=self.config.query_params,
            timeout=self.config.timeout,
            verify=self.config.verify_tls,
        )
        response.raise_for_status()
        payload = response.json()
        orders = self._normalize_orders_payload(payload)
        return [order for order in orders if self._has_pending_details(order)]

    def mark_detail_as_printed(self, order: Dict[str, Any], detail: Dict[str, Any]) -> None:
        substitutions = self._build_substitutions(order, detail)
        url = self.config.printed_url_template.format(**substitutions)
        payload = self._render_payload_template(substitutions)
        response = self.session.request(
            self.config.printed_method,
            url,
            json=payload,
            timeout=self.config.timeout,
            verify=self.config.verify_tls,
        )
        response.raise_for_status()

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

    def _render_payload_template(self, substitutions: Dict[str, Any]) -> Dict[str, Any]:
        rendered = Template(self.config.printed_payload_template).safe_substitute(
            {key: str(value) for key, value in substitutions.items()}
        )
        payload = json.loads(rendered)
        if not isinstance(payload, dict):
            raise ValueError("API_PRINTED_PAYLOAD_TEMPLATE debe renderizar un objeto JSON")
        return payload

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

        document = self._build_document(order, center, list(details))
        raw_bytes = document.encode(self.config.print_encoding, errors="replace")

        logging.info("Enviando pedido %s al centro %s en impresora '%s'", order.get("id"), center, printer_name)
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            for copy_number in range(self.config.print_copies):
                job_id = win32print.StartDocPrinter(printer_handle, 1, (self.config.document_title, None, "RAW"))
                try:
                    win32print.StartPagePrinter(printer_handle)
                    win32print.WritePrinter(printer_handle, raw_bytes)
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
        finally:
            win32print.ClosePrinter(printer_handle)

    def _build_document(self, order: Dict[str, Any], center: str, details: List[Dict[str, Any]]) -> str:
        order_id = order.get("id", "")
        order_number = f"{order.get('OrderSerie', '')}-{order.get('OrderNumber', '')}".strip("-")
        customer = order.get("ExternalPerson", "")
        cashier = order.get("Cashier") or order.get("InternalPerson") or ""
        observations = order.get("Observations") or ""
        created_at = order.get("Hour") or order.get("Created") or ""
        header_lines = [
            "PEDIDO DE PRODUCCION",
            f"CENTRO: {center}",
            f"PEDIDO: {order_id}",
            f"NUMERO: {order_number}",
            f"FECHA: {self._format_datetime(created_at)}",
        ]
        if cashier:
            header_lines.append(f"CAJERO: {cashier}")
        if customer:
            header_lines.append(f"CLIENTE: {customer}")

        item_lines = []
        for detail in details:
            quantity = self._format_quantity(detail.get("Quantity"))
            product = (detail.get("Product") or "").strip()
            item_lines.append(f"{quantity} x {product}"[:120])
            detail_obs = (detail.get("Observations") or "").strip()
            if detail_obs:
                item_lines.append(f"  Obs: {detail_obs}"[:120])

        footer_lines = []
        if observations:
            footer_lines.append(f"OBS GENERAL: {observations}"[:120])
        footer_lines.append(f"ITEMS: {len(details)}")

        lines = header_lines + ["-" * 40] + item_lines + ["-" * 40] + footer_lines
        lines.extend([""] * self.config.cut_lines)
        return "\n".join(lines)

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
        grouped[center].append(detail)
    return grouped


def wait_for_next_cycle(seconds: int) -> bool:
    if SERVICE_STOP_EVENT:
        wait_result = win32event.WaitForSingleObject(SERVICE_STOP_EVENT, max(seconds, 1) * 1000)
        return wait_result == win32event.WAIT_OBJECT_0
    time.sleep(max(seconds, 1))
    return False


def run_cycle(client: OrderApiClient, printer: TicketPrinter) -> Tuple[int, int]:
    printed_count = 0
    orders = client.get_pending_orders()
    if not orders:
        logging.info("No hay pedidos pendientes por imprimir.")
        return 0, 0

    logging.info("Se encontraron %s pedido(s) con detalles pendientes.", len(orders))
    for order in orders:
        order_id = order.get("id")
        grouped = group_pending_details(order)
        if not grouped:
            continue

        logging.info("Procesando pedido %s en %s centro(s).", order_id, len(grouped))
        for center, details in grouped.items():
            printer.print_order_group(order, center, details)
            for detail in details:
                client.mark_detail_as_printed(order, detail)
                printed_count += 1
                logging.info(
                    "Detalle marcado como impreso. Pedido=%s Detalle=%s Centro=%s",
                    order_id,
                    detail.get("id"),
                    center,
                )
    return len(orders), printed_count


def main() -> None:
    load_environment()
    setup_logging()
    config = Config()
    client = OrderApiClient(config)
    printer = TicketPrinter(config)

    logging.info("ATIC Peru")
    logging.info("Iniciando print_server.py version %s", SCRIPT_VERSION)
    logging.info("Consulta configurada cada %s segundo(s).", config.poll_seconds)

    while True:
        try:
            orders_count, printed_count = run_cycle(client, printer)
            sleep_seconds = config.idle_sleep_seconds if printed_count == 0 else config.poll_seconds
            logging.info(
                "Ciclo finalizado. Pedidos=%s Detalles impresos=%s Proxima consulta en %s segundo(s).",
                orders_count,
                printed_count,
                sleep_seconds,
            )
        except Exception as exc:
            logging.error("Error en el ciclo principal: %s", exc, exc_info=True)
            sleep_seconds = config.error_sleep_seconds

        if wait_for_next_cycle(sleep_seconds):
            logging.info("Se recibio senal de parada del servicio.")
            break


if __name__ == "__main__":
    main()
