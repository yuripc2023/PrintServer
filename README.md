# PrintServer

Servicio Windows en Python para escuchar pedidos en tiempo real por WebSocket, separarlos por `ProductionCenter`, imprimir cada grupo en su impresora y luego actualizar el pedido con `PATCH` enviando `Details.Printed = true`.

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

- `API_ORDERS_URL`: endpoint REST usado para marcar pedidos como impresos.
- `Company`: id de empresa a consultar.
- `ORDER_STATUS`: estado a consultar. Por defecto `Registrado`.
- `API_AUTH_MODE`: `basic`, `bearer`, `token` o `none`.
- `API_USERNAME` y `API_PASSWORD`, o `API_TOKEN` segun el caso.
- `WS_ORDERS_URL`: canal WebSocket de pedidos. Si se deja vacio, se construye como `wss://<host>/ws/restaurants/<WS_RESTAURANT_ID>/tables/`.
- `WS_RESTAURANT_ID`: id del restaurante usado para construir el WebSocket cuando `WS_ORDERS_URL` no esta definido.
- `WS_RECONNECT_DELAY_SECONDS`: segundos de espera antes de reconectar el WebSocket si la conexion se corta.
- `SYNC_PENDING_ON_CONNECT`: cuando vale `true`, al conectar o reconectar hace una consulta REST de respaldo para recuperar pedidos pendientes perdidos durante una suspension o corte de red.
- `API_PRINTED_URL_TEMPLATE`: endpoint real para marcar el pedido como impreso enviando `Details`.
- `PRINTER_MAP_JSON`: mapa de centros hacia el nombre exacto de la impresora instalada en Windows.
- `PRECUENTA_PRINTER_NAME`: impresora para la precuenta. Si se deja vacio, se usa la primera impresora disponible del mapa de centros del pedido.
- `PRECUENTA_COPIES`: cantidad de copias a imprimir para la precuenta.

Ejemplo de canal:

```text
wss://api.atic.pe/ws/restaurants/1/tables/
```

## Ejecucion manual

```powershell
python .\print_server.py
```

El proceso queda escuchando eventos `order.created` y `order.updated` para imprimir solo cuando llegue un pedido nuevo o actualizado.

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

Instalar o actualizar automaticamente el servicio existente:

```powershell
powershell -ExecutionPolicy Bypass -File .\install_or_update_service.ps1
```

Si usas otro ejecutable de Python:

```powershell
powershell -ExecutionPolicy Bypass -File .\install_or_update_service.ps1 -PythonExe "C:\Ruta\python.exe"
```

## Observaciones tecnicas

- El log rota diariamente y conserva `LOG_BACKUP_COUNT` archivos historicos.
- La impresion ya no depende de un `GET` ciclico; ahora se activa por eventos WebSocket.
- Si Windows entra en suspension o la red cae, el servicio reintenta el WebSocket y puede hacer una sincronizacion REST de respaldo al reconectar.
- Si `WS_ORDERS_URL` no esta definido, el servicio construye la ruta usando el host de `API_ORDERS_URL` y `WS_RESTAURANT_ID`.
- El servicio imprime un ticket por pedido y por centro de produccion.
- Si `StatusInvoice` llega como `precuenta`, el servicio imprime una precuenta con el mismo JSON del evento y luego actualiza `StatusInvoice` a `-` para no repetirla.
- Si un detalle ya viene con `Printed=true`, no se vuelve a imprimir.
- La confirmacion se hace con `PATCH` al pedido completo usando `API_PRINTED_URL_TEMPLATE`, por ejemplo `https://apisayri.atic.pe/api/orders/ordersales/{order_id}/`.
- Si la confirmacion a la API falla, el cache local evita reimpresiones repetidas mientras el servicio sigue encendido cuando `REPRINT_WHEN_NOT_CONFIRMED=false`.
