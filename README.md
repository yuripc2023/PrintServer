# PrintServer

Servicio Windows en Python para consultar pedidos pendientes desde una API REST, separarlos por `ProductionCenter`, imprimir cada grupo en su impresora y luego marcar `Details.Printed = true`.

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
- `API_AUTH_MODE`: `basic`, `bearer`, `token` o `none`.
- `API_USERNAME` y `API_PASSWORD`, o `API_TOKEN` segun el caso.
- `API_PRINTED_URL_TEMPLATE`: endpoint real para marcar un detalle como impreso.
- `PRINTER_MAP_JSON`: mapa de centros hacia el nombre exacto de la impresora instalada en Windows.

## Ejecucion manual

```powershell
python .\print_server.py
```

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
- La ruta exacta para actualizar `Printed` puede variar segun tu backend. Por eso se dejo configurable con `API_PRINTED_URL_TEMPLATE` y `API_PRINTED_METHOD`.
