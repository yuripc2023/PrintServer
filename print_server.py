import argparse
import json
import logging
import os
import sys
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
PRINT_CACHE_PATH = os.path.join(BASE_DIR, "printed_cache.json")
SERVICE_STOP_EVENT = None
TICKET_WIDTH = 40


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
        self.reprint_when_not_confirmed = os.getenv("REPRINT_WHEN_NOT_CONFIRMED", "true").strip().lower() not in {
            "0",
            "false",
            "no",
        }
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

    def build_orders_query_params(self) -> Dict[str, Any]:
        params = dict(self.query_params)
        if self.company:
            params["Company"] = self.company
        if self.order_status:
            params["Status"] = self.order_status
        return params


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

    def _build_document(self, order: Dict[str, Any], center: str, details: List[Dict[str, Any]]) -> str:
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
            f"CENTRO DE PRODUCCION: {center}".center(TICKET_WIDTH),
            "=" * TICKET_WIDTH,
            f"NUMERO: {order_number}",
            f"FECHA: {self._format_datetime(created_at)}",
        ]
        if cashier:
            header_lines.append(f"ASESOR: {cashier}")
        if workspace or space:
            header_lines.append("-" * TICKET_WIDTH)
        if workspace:
            header_lines.append(f"SALON: {workspace}")
        if space:
            header_lines.append(f"ESPACIO: {space}")

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

        lines = header_lines + ["-" * TICKET_WIDTH] + item_lines + ["-" * TICKET_WIDTH] + footer_lines
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


def split_order_details(
    order: Dict[str, Any], print_cache: Dict[str, float], config: Config
) -> Tuple[Dict[str, List[Dict[str, Any]]], List[Dict[str, Any]]]:
    grouped_to_print: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    pending_confirmation: List[Dict[str, Any]] = []

    for detail in order.get("Details") or []:
        if bool(detail.get("Printed")):
            continue

        if not config.reprint_when_not_confirmed and is_detail_cached(print_cache, detail.get("id")):
            pending_confirmation.append(detail)
            continue

        center = normalize_center(detail.get("ProductionCenter"))
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


def try_mark_detail(
    client: OrderApiClient, order: Dict[str, Any], detail: Dict[str, Any], print_cache: Dict[str, float]
) -> bool:
    try:
        client.mark_detail_as_printed(order, detail)
        uncache_detail_printed(print_cache, detail.get("id"))
        logging.info(
            "Detalle marcado como impreso. Pedido=%s Detalle=%s Centro=%s",
            order.get("id"),
            detail.get("id"),
            normalize_center(detail.get("ProductionCenter")),
        )
        return True
    except Exception as exc:
        return False


def wait_for_next_cycle(seconds: int) -> bool:
    if SERVICE_STOP_EVENT:
        wait_result = win32event.WaitForSingleObject(SERVICE_STOP_EVENT, max(seconds, 1) * 1000)
        return wait_result == win32event.WAIT_OBJECT_0
    time.sleep(max(seconds, 1))
    return False


def run_cycle(
    client: OrderApiClient, printer: TicketPrinter, print_cache: Dict[str, float], config: Config
) -> Tuple[int, int]:
    printed_count = 0
    confirmed_count = 0
    orders = client.get_pending_orders()
    if not orders:
        logging.info("No hay pedidos pendientes por imprimir.")
        return 0, 0

    logging.info("Se encontraron %s pedido(s) con detalles pendientes.", len(orders))
    for order in orders:
        order_id = order.get("id")
        log_order_detail_centers(order, config)
        grouped, pending_confirmation = split_order_details(order, print_cache, config)
        if grouped:
            logging.info("Procesando pedido %s en %s centro(s).", order_id, len(grouped))
        for center, details in grouped.items():
            printer.print_order_group(order, center, details)
            for detail in details:
                cache_detail_printed(print_cache, detail.get("id"))
                printed_count += 1
                if try_mark_detail(client, order, detail, print_cache):
                    confirmed_count += 1

        for detail in pending_confirmation:
            if try_mark_detail(client, order, detail, print_cache):
                confirmed_count += 1

    save_print_cache(print_cache)
    logging.info(
        "Estado del ciclo. Detalles impresos=%s Detalles confirmados en API=%s Pendientes locales=%s",
        printed_count,
        confirmed_count,
        len(print_cache),
    )
    return len(orders), printed_count


def main() -> None:
    load_environment()
    setup_logging()
    config = Config()
    client = OrderApiClient(config)
    printer = TicketPrinter(config)
    print_cache = load_print_cache()

    logging.info("ATIC Peru")
    logging.info("Iniciando print_server.py version %s", SCRIPT_VERSION)
    logging.info("Consulta configurada cada %s segundo(s).", config.poll_seconds)

    while True:
        try:
            orders_count, printed_count = run_cycle(client, printer, print_cache, config)
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
    if handle_cli(sys.argv[1:]):
        sys.exit(0)
    main()
