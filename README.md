# PrintServer

Servicio Windows en Python para consultar pedidos pendientes desde una API REST, separarlos por `ProductionCenter`, imprimir cada grupo en su impresora y luego actualizar el pedido con `PATCH` enviando `Details.Printed = true`.

## Archivos principales

- `print_server.py`: logica principal de consulta, agrupacion, impresion y actualizacion.
- `print_server_service.py`: wrapper para instalar y ejecutar como servicio Windows.
- `.env.example`: plantilla de configuracion.
- `print_server.log`: log rotativo diario con retencion configurable.

## Configuracion

1. Instalar Python para Windows y desactivar los alias `python.exe` y `python3.exe` de Microsoft Store si aplica.
2. Instalar dependencias:

```powershell
pip install -r requirements.txt
```

3. Crear `.env` a partir de `.env.example` y completar:

- `API_ORDERS_URL`: endpoint de consulta de pedidos.
- `Company`: id de empresa a consultar.
- `ORDER_STATUS`: estado a consultar. Por defecto `Registrado`.
- `API_AUTH_MODE`: `basic`, `bearer`, `token` o `none`.
- `API_USERNAME` y `API_PASSWORD`, o `API_TOKEN` segun el caso.
- `API_PRINTED_URL_TEMPLATE`: endpoint real para marcar el pedido como impreso enviando `Details`.
- `PRINTER_MAP_JSON`: mapa de centros hacia el nombre exacto de la impresora instalada en Windows.

Con eso la consulta queda asi:

```text
https://apisayri.atic.pe/api/orders/ordersales/?Company=10813&Status=Registrado
```

## Ejecucion manual

```powershell
python .\print_server.py
```

Listar impresoras instaladas en Windows:

```powershell
python .\print_server.py --list-printers
```

Generar `PRINTER_MAP_JSON` en `.env` usando la primera impresora instalada para cada centro:

```powershell
python .\print_server.py --sync-printer-map --centers COCINA,BARRA,PARRILLAS
```

Luego puedes editar el valor generado si quieres asignar una impresora distinta a cada centro.

Si la impresora no corta al final del ticket, configura el comando ESC/POS en `.env`:

```env
PRINT_CUT_ENABLED=true
PRINT_CUT_COMMAND_HEX=1D5641
```

Algunos modelos usan `1D5600` en lugar de `1D5641`.

Si ves mal las tildes o la `Ñ`, revisa la combinacion de:

```env
PRINT_ENCODING=cp850
PRINT_CODEPAGE_COMMAND_HEX=1B7402
```

Ese comando selecciona la tabla de caracteres en la impresora antes de enviar el texto. Si tu modelo usa otra tabla ESC/POS, solo cambia `PRINT_CODEPAGE_COMMAND_HEX`.

## Instalacion como servicio Windows

Instalar:

```powershell
python .\print_server_service.py install
```

Iniciar:

```powershell
python .\print_server_service.py start
```

Dejar inicio automatico al reiniciar Windows:

```powershell
cmd /c sc config ATICPrintServer start= auto
```

Ver estado:

```powershell
Get-Service -Name "ATICPrintServer"
```

Detener:

```powershell
python .\print_server_service.py stop
```

Eliminar:

```powershell
python .\print_server_service.py remove
```

## Observaciones tecnicas

- El log rota diariamente y conserva `LOG_BACKUP_COUNT` archivos historicos.
- El valor por defecto de `POLL_SECONDS` es `10`. Es mejor punto de partida que `5` para no cargar innecesariamente la API. Si la API soporta filtros por pendientes, usalos en `API_QUERY_PARAMS_JSON`.
- El servicio imprime un ticket por pedido y por centro de produccion.
- Si un detalle ya viene con `Printed=true`, no se vuelve a imprimir.
- La confirmacion se hace con `PATCH` al pedido completo usando `API_PRINTED_URL_TEMPLATE`, por ejemplo `https://apisayri.atic.pe/api/orders/ordersales/{order_id}/`.
- Si la confirmacion a la API falla, el cache local evita reimpresiones repetidas mientras el servicio sigue encendido cuando `REPRINT_WHEN_NOT_CONFIRMED=false`.
