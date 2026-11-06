# =====================================================================
#  arreglar_knime.ps1
#  Repara KNIME para el proyecto NOVA vs SAP:
#    1. Encuentra knime.exe en la PC.
#    2. Crea una carpeta de trabajo LOCAL (workspace) donde KNIME si puede escribir.
#    3. Copia el workflow NOVA_SAP_Trigger.knwf (de Descargas) dentro de esa carpeta.
#    4. Crea un acceso directo nuevo en el Escritorio que abre KNIME con esa carpeta.
#    5. Abre KNIME.
#
#  Como correrlo (no necesita ser administrador):
#    Clic derecho sobre el archivo -> "Ejecutar con PowerShell"
#  o en una ventana de PowerShell:
#    powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\Downloads\arreglar_knime.ps1"
# =====================================================================

$ErrorActionPreference = 'Stop'

function Paso($texto) { Write-Host "`n==> $texto" -ForegroundColor Cyan }
function Ok($texto)   { Write-Host "    OK: $texto" -ForegroundColor Green }
function Aviso($texto){ Write-Host "    AVISO: $texto" -ForegroundColor Yellow }
function Falla($texto){ Write-Host "    ERROR: $texto" -ForegroundColor Red }

function Probar-Escritura($carpeta) {
    try {
        New-Item -ItemType Directory -Force -Path $carpeta | Out-Null
        $prueba = Join-Path $carpeta ('.prueba_' + [guid]::NewGuid().ToString('N'))
        Set-Content -Path $prueba -Value 'ok' -ErrorAction Stop
        Remove-Item $prueba -Force
        return $true
    } catch { return $false }
}

function Esta-En-OneDrive($ruta) {
    $od = @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer) | Where-Object { $_ }
    foreach ($o in $od) { if ($ruta -like "$o*") { return $true } }
    return ($ruta -match '\\OneDrive')
}

try {
    # -----------------------------------------------------------------
    Paso 'Buscando KNIME en esta computadora'
    $candidatos = New-Object System.Collections.Generic.List[string]

    # a) Accesos directos existentes (Escritorio y menu Inicio)
    $shell = New-Object -ComObject WScript.Shell
    $carpetasLnk = @(
        [Environment]::GetFolderPath('Desktop'),
        [Environment]::GetFolderPath('CommonDesktopDirectory'),
        [Environment]::GetFolderPath('StartMenu'),
        [Environment]::GetFolderPath('CommonStartMenu')
    ) | Where-Object { $_ -and (Test-Path $_) }
    foreach ($c in $carpetasLnk) {
        Get-ChildItem -Path $c -Filter '*KNIME*.lnk' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $destino = $shell.CreateShortcut($_.FullName).TargetPath
            if ($destino) { $candidatos.Add($destino) }
        }
    }
    # b) Rutas de instalacion tipicas
    $raices = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA,
                (Join-Path $env:LOCALAPPDATA 'Programs'), $env:USERPROFILE, 'C:\', 'D:\') |
              Where-Object { $_ -and (Test-Path $_) }
    foreach ($r in $raices) {
        Get-ChildItem -Path $r -Directory -Filter '*knime*' -ErrorAction SilentlyContinue | ForEach-Object {
            $candidatos.Add((Join-Path $_.FullName 'knime.exe'))
            Get-ChildItem -Path $_.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $candidatos.Add((Join-Path $_.FullName 'knime.exe'))
            }
        }
    }
    # c) Proceso de KNIME abierto ahora mismo
    Get-Process -Name knime -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Path) { $candidatos.Add($_.Path) }
    }

    $knimeExe = $candidatos | Where-Object { $_ -and (Test-Path $_) -and ($_ -like '*knime.exe') } |
                Select-Object -Unique | Select-Object -First 1
    if (-not $knimeExe) {
        Falla 'No encontre knime.exe. Instala KNIME desde https://www.knime.com/downloads (o pidelo a TI) y vuelve a correr este script.'
        throw 'KNIME no encontrado'
    }
    Ok "KNIME encontrado: $knimeExe"

    # -----------------------------------------------------------------
    Paso 'Creando la carpeta de trabajo local (workspace)'
    $opciones = @('C:\KNIME_Robot\workspace', (Join-Path $env:USERPROFILE 'KNIME_workspace'), (Join-Path $env:LOCALAPPDATA 'KNIME_workspace'))
    $workspace = $null
    foreach ($o in $opciones) {
        if (Esta-En-OneDrive $o) { continue }
        if (Probar-Escritura $o) { $workspace = $o; break }
        Aviso "No se puede escribir en $o, pruebo otra ubicacion"
    }
    if (-not $workspace) { throw 'No encontre una carpeta local donde se pueda escribir.' }
    Ok "Workspace: $workspace"

    # -----------------------------------------------------------------
    Paso 'Buscando el archivo NOVA_SAP_Trigger.knwf'
    $descargas = $null
    try { $descargas = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path } catch {}
    $dondeBuscar = @($descargas, (Join-Path $env:USERPROFILE 'Downloads'), [Environment]::GetFolderPath('Desktop'),
                     [Environment]::GetFolderPath('MyDocuments'), $PSScriptRoot) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
    $knwf = $null
    foreach ($d in $dondeBuscar) {
        $knwf = Get-ChildItem -Path $d -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^NOVA_SAP_Trigger.*\.knwf(\.zip)?$' -or $_.Name -match '^NOVA_SAP_Trigger.*\.zip$' } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($knwf) { break }
    }

    $destinoWf = Join-Path $workspace 'NOVA_SAP_Trigger'
    if ($knwf) {
        Ok "Archivo encontrado: $($knwf.FullName)"
        if (Test-Path $destinoWf) {
            $respaldo = "$destinoWf.respaldo_$(Get-Date -Format yyyyMMdd_HHmmss)"
            Move-Item $destinoWf $respaldo
            Aviso "Ya existia un NOVA_SAP_Trigger; lo movi a $respaldo"
        }
        # Un .knwf es un .zip con la carpeta del workflow adentro
        $tmpZip = Join-Path $env:TEMP 'NOVA_SAP_Trigger_import.zip'
        Copy-Item $knwf.FullName $tmpZip -Force
        $tmpDir = Join-Path $env:TEMP ('knwf_' + [guid]::NewGuid().ToString('N'))
        Expand-Archive -Path $tmpZip -DestinationPath $tmpDir -Force
        $wfKnime = Get-ChildItem -Path $tmpDir -Filter 'workflow.knime' -Recurse | Sort-Object { $_.FullName.Length } | Select-Object -First 1
        if (-not $wfKnime) { throw 'El archivo descargado no parece un workflow de KNIME (no trae workflow.knime). Vuelve a descargarlo.' }
        Copy-Item -Path $wfKnime.Directory.FullName -Destination $destinoWf -Recurse -Force
        Remove-Item $tmpZip, $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
        Ok "Workflow listo en: $destinoWf"
    } else {
        Aviso 'No encontre NOVA_SAP_Trigger.knwf en Descargas, Escritorio ni Documentos.'
        Aviso 'Descargalo y vuelve a correr el script, o importalo desde KNIME despues.'
    }

    # -----------------------------------------------------------------
    Paso 'Creando acceso directo nuevo en el Escritorio'
    $escritorio = [Environment]::GetFolderPath('Desktop')
    $viejo = Join-Path $escritorio 'KNIME Analytics Platform.lnk'
    if (Test-Path $viejo) {
        $destViejo = $shell.CreateShortcut($viejo).TargetPath
        if (-not ($destViejo -and (Test-Path $destViejo))) {
            Remove-Item $viejo -Force
            Aviso 'Borre el acceso directo viejo que apuntaba a una ubicacion que ya no existe.'
        }
    }
    $lnkPath = Join-Path $escritorio 'KNIME (NOVA SAP).lnk'
    $lnk = $shell.CreateShortcut($lnkPath)
    $lnk.TargetPath = $knimeExe
    $lnk.Arguments = "-data `"$workspace`""
    $lnk.WorkingDirectory = Split-Path $knimeExe
    $lnk.IconLocation = "$knimeExe,0"
    $lnk.Description = 'KNIME con workspace local para NOVA vs SAP'
    $lnk.Save()
    Ok "Acceso directo: $lnkPath"

    # -----------------------------------------------------------------
    Paso 'Abriendo KNIME con la carpeta nueva'
    $abierto = Get-Process -Name knime -ErrorAction SilentlyContinue
    if ($abierto) {
        $r = Read-Host '    KNIME esta abierto. Cierralo (guarda lo que necesites) y presiona ENTER para continuar, o escribe N para no abrirlo'
        if ($r -match '^[nN]') { $knimeExe = $null }
        elseif (Get-Process -Name knime -ErrorAction SilentlyContinue) {
            Aviso 'KNIME sigue abierto; lo cierro.'
            Get-Process -Name knime | Stop-Process -Force
            Start-Sleep -Seconds 3
        }
    }
    if ($knimeExe) {
        Start-Process -FilePath $knimeExe -ArgumentList '-data', "`"$workspace`"" -WorkingDirectory (Split-Path $knimeExe)
        Ok 'KNIME se esta abriendo (puede tardar 1-2 minutos).'
    }

    Write-Host "`n================ LISTO ================" -ForegroundColor Green
    Write-Host " En KNIME: Home -> Local space -> abre 'NOVA_SAP_Trigger' (si no aparece, dale Reload)."
    Write-Host " Luego: Execute all. Para tus datos: doble clic en cada Excel Reader y elige tu archivo."
    Write-Host " De ahora en adelante abre KNIME con el acceso directo 'KNIME (NOVA SAP)' del Escritorio."
    Write-Host " NO le des 'Sign in' a ningun Hub; no hace falta."
}
catch {
    Falla $_.Exception.Message
    Write-Host "`nMandale una captura de esta ventana a quien te ayuda con esto." -ForegroundColor Yellow
}
finally {
    Read-Host "`nPresiona ENTER para cerrar"
}
