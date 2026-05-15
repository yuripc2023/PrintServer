# PrintServer

Servicio Windows en Python para escuchar pedidos en tiempo real por WebSocket, separarlos por `ProductionCenter`, imprimir cada grupo en su impresora y luego actualizar el pedido con `PATCH` enviando `Details.Printed = true`.

## GIT

https://github.com/yuripc2023/PrintServer

## Archivos principales

- `print_server.py`: logica principal de consulta, agrupacion, impresion y actualizacion.
- `print_server_service.py`: wrapper para instalar y ejecutar como servicio Windows.
- `settings.toml`: configuracion operativa del servicio, API, WebSocket e impresion.
- `.env`: credenciales de acceso a la API.
- `print_server.log`: log rotativo diario con retencion configurable.

## Configuracion

1. Instalar Python para Windows y desactivar los alias `python.exe` y `python3.exe` de Microsoft Store si aplica.
2. Instalar dependencias:

```powershell
pip install -r requirements.txt
```

3. Completar API y credenciales en `.env`:

- `API_ORDERS_URL`: endpoint REST de pedidos.
- `API_USERNAME` y `API_PASSWORD` para autenticacion basic.
- `API_TOKEN` para autenticacion bearer/token.

4. Ajustar configuracion en `settings.toml`:

- `[api].company`: id de empresa a consultar.
- `[api].order_status`: estado a consultar. Por defecto `Registrado`.
- `[api].auth_mode`: `basic`, `bearer`, `token` o `none`.
- `[websocket].url`: canal WebSocket de pedidos. Si se deja vacio, se construye como `wss://<host>/ws/restaurants/<company>/tables/`.
- `[websocket].reconnect_delay_seconds`: segundos de espera antes de reconectar el WebSocket si la conexion se corta.
- `[websocket].sync_pending_on_connect`: cuando vale `true`, al conectar o reconectar hace una consulta REST de respaldo para recuperar pedidos pendientes.
- `[api].printed_url_template`: endpoint real para marcar el pedido como impreso enviando `Details`. Si se deja vacio, se genera desde `[api].orders_url` agregando `{order_id}/`.
- `[printers].map`: mapa de centros hacia el nombre exacto de la impresora instalada en Windows.
- `[print].precuenta_printer_name`: impresora para la precuenta. Si se deja vacio, se usa la primera impresora disponible del mapa de centros del pedido.
- `[print].precuenta_copies`: cantidad de copias a imprimir para la precuenta.

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

Generar `[printers].map` en `settings.toml` usando la primera impresora instalada para cada centro:

```powershell
python .\print_server.py --sync-printer-map --centers COCINA,BARRA,PARRILLAS
```

Luego puedes editar el valor generado si quieres asignar una impresora distinta a cada centro.

Si la impresora no corta al final del ticket, configura el comando ESC/POS en `settings.toml`:

```toml
[print]
cut_enabled = true
cut_command_hex = "1D5641"
```

Algunos modelos usan `1D5600` en lugar de `1D5641`.

Para agrandar el detalle de produccion sin cambiar el ancho de las columnas:

```toml
[print]
detail_size_command_hex = "1D2110"
detail_size_reset_command_hex = "1D2100"
```

Si ves mal las tildes o la `Ñ`, revisa la combinacion de:

```toml
[print]
encoding = "cp850"
codepage_command_hex = "1B7402"
```

Ese comando selecciona la tabla de caracteres en la impresora antes de enviar el texto. Si tu modelo usa otra tabla ESC/POS, solo cambia `[print].codepage_command_hex`.

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

- El log rota diariamente y conserva `[service].log_backup_count` archivos historicos.
- La impresion ya no depende de un `GET` ciclico; ahora se activa por eventos WebSocket.
- Si Windows entra en suspension o la red cae, el servicio reintenta el WebSocket y puede hacer una sincronizacion REST de respaldo al reconectar.
- Si `[websocket].url` no esta definido, el servicio construye la ruta usando el host de `[api].orders_url` y `[api].company`.
- El servicio imprime un ticket por pedido y por centro de produccion.
- Si `StatusInvoice` llega como `precuenta`, el servicio imprime una precuenta con el mismo JSON del evento y luego actualiza `StatusInvoice` a `-` para no repetirla.
- Si un detalle ya viene con `Printed=true`, no se vuelve a imprimir.
- La confirmacion se hace con `PATCH` al pedido completo. Si `[api].printed_url_template` esta vacio, se usa `[api].orders_url` y se agrega `{order_id}/`, por ejemplo `https://apisayri.atic.pe/api/orders/ordersales/{order_id}/`.
- Si la confirmacion a la API falla, el cache local evita reimpresiones repetidas mientras el servicio sigue encendido cuando `[print].reprint_when_not_confirmed=false`.
