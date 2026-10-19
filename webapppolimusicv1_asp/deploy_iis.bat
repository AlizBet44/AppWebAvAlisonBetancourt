@echo off
chcp 65001 >nul
title PoliMusic - Despliegue Automatizado en IIS (Plug & Play)

echo ==============================================================================
echo        PoliMusic - Spotify Blue Edition (ASP Clásico + SQL Server)
echo                  Script de Despliegue Automatizado en IIS
echo ==============================================================================
echo.

set TARGET_DIR=C:\inetpub\wwwroot\PoliMusic
set SOURCE_DIR=%~dp0

:: 1. Comprobar privilegios de Administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ADVERTENCIA] Por favor ejecuta este script como Administrador.
    echo (Haz clic derecho sobre "deploy_iis.bat" y elige "Ejecutar como Administrador").
    echo.
    pause
    exit /b 1
)

:: 2. Crear directorios de destino si no existen
echo [1/4] Creando directorios en IIS: %TARGET_DIR% ...
if not exist "C:\inetpub" mkdir "C:\inetpub"
if not exist "C:\inetpub\wwwroot" mkdir "C:\inetpub\wwwroot"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: 3. Copiar todos los archivos de la aplicación PoliMusic
echo [2/4] Copiando archivos de la aplicacion PoliMusic a inetpub...
robocopy "%SOURCE_DIR%" "%TARGET_DIR%" /E /XD .git .vs .system_generated .gemini scratch /XF *.log *.ps1 /NFL /NDL /NJH /NJS /nc /ns /np

:: 4. Configurar Permisos de Seguridad NTFS requeridos por IIS
echo [3/4] Asignando permisos NTFS a cuentas de servicio IIS...
icacls "%TARGET_DIR%" /grant "IIS_IUSRS":(OI)(CI)RX /T >nul 2>&1
icacls "%TARGET_DIR%" /grant "IUSR":(OI)(CI)RX /T >nul 2>&1
icacls "%TARGET_DIR%" /grant "Usuarios":(OI)(CI)RX /T >nul 2>&1
icacls "%TARGET_DIR%" /grant "Todos":(OI)(CI)RX /T >nul 2>&1

:: 5. Comprobar si el módulo ASP Clásico está habilitado en Windows
echo [4/4] Verificando modulo de ASP Clasico en Windows IIS...
dism /online /get-featureinfo /featurename:IIS-ASP | findstr /i "Enabled Habilitada" >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo [AVISO] Si IIS no tiene activado el modulo ASP Clasico,
    echo puedes habilitarlo ejecutando en PowerShell como Administrador:
    echo Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP,IIS-WebServerRole -All
)

echo.
echo ==============================================================================
echo                      DESPLIEGUE COMPLETADO EXITOSAMENTE
echo ==============================================================================
echo  Ruta Fisica:   %TARGET_DIR%
echo  URL IIS:       http://localhost/PoliMusic/login.asp
echo.
echo  Credenciales de Prueba (en login.asp):
echo    - Administrador: userAdmin  / 12345678
echo    - Usuario EPN:   userNormal / 12345678
echo ==============================================================================
echo.
pause
