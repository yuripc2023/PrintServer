[CmdletBinding()]
param(
    [string]$PythonExe = "python",
    [string]$ServiceName = "ATICPrintServer"
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "==> $Message"
}

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    throw "Este script debe ejecutarse como Administrador."
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Step "Validando Python"
& $PythonExe --version | Out-Null

$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

if ($null -ne $existingService) {
    Write-Step "Servicio existente detectado: $ServiceName"

    if ($existingService.Status -ne "Stopped") {
        Write-Step "Deteniendo servicio"
        & $PythonExe .\print_server_service.py stop
        Start-Sleep -Seconds 2
    }

    Write-Step "Eliminando instalacion anterior"
    & $PythonExe .\print_server_service.py remove
    Start-Sleep -Seconds 2
}

Write-Step "Instalando servicio"
& $PythonExe .\print_server_service.py install

Write-Step "Configurando inicio automatico"
& sc.exe config $ServiceName start= auto | Out-Null

Write-Step "Iniciando servicio"
& $PythonExe .\print_server_service.py start

Write-Step "Estado final del servicio"
Get-Service -Name $ServiceName | Format-Table -AutoSize

Write-Step "Proceso completado"
