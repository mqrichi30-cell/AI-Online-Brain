<#
.SYNOPSIS
    Arma la tabla dinamica de correos por hora a partir de un CSV ya exportado,
    sin volver a consultar Microsoft Graph.

.DESCRIPTION
    Pensado para cuando la exportacion ya termino (el CSV pesa y tardo en bajarse)
    pero el paso de Excel fallo. Repite solo la parte de Excel, con diagnostico
    paso a paso y verificacion real del archivo de salida.

    Diferencias frente al intento embebido en Export-And-PivotByHour.ps1:

      * Excel queda VISIBLE, asi cualquier dialogo que aparezca se ve.
      * DisplayAlerts se deja en $true durante SaveAs. Con DisplayAlerts en
        $false, Excel responde que NO al aviso de sobrescritura y la grabacion
        se salta EN SILENCIO, sin lanzar error: ese es el modo de fallo que
        hace que el script diga "listo" y el .xlsx no exista.
      * Se borra el destino antes de grabar y se comprueba con Test-Path
        despues. Nunca se declara exito sin que el archivo exista.
      * Si la carpeta de destino no acepta escritura (OneDrive, Acceso
        controlado a carpetas de Windows Defender), reintenta en el Escritorio
        y luego en TEMP antes de rendirse.

.EXAMPLE
    .\Build-PivotFromCsv.ps1

.EXAMPLE
    .\Build-PivotFromCsv.ps1 -CsvPath "C:\Users\marin.c\Downloads\CorreosUltimos6Meses.csv"
#>

[CmdletBinding()]
param(
    [string] $CsvPath  = "$env:USERPROFILE\Downloads\CorreosUltimos6Meses.csv",
    [string] $XlsxPath = "",
    [string] $Titulo   = "Correos por hora del dia (24h)",
    [switch] $Invisible
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "No existe el CSV '$CsvPath'."
}
$CsvPath = (Resolve-Path -LiteralPath $CsvPath).Path

if ([string]::IsNullOrWhiteSpace($XlsxPath)) {
    $XlsxPath = [IO.Path]::ChangeExtension($CsvPath, '.xlsx')
}

Write-Host "CSV origen : $CsvPath"
Write-Host "Destino    : $XlsxPath"
Write-Host ""

# Candidatos de destino, por si la carpeta preferida rechaza la escritura.
# Escritorio y TEMP pueden no estar disponibles: se descartan si vienen vacios.
$nombre    = [IO.Path]::GetFileName($XlsxPath)
$escritorio = [Environment]::GetFolderPath('Desktop')
$destinos = @($XlsxPath)
if (-not [string]::IsNullOrWhiteSpace($escritorio)) { $destinos += (Join-Path $escritorio $nombre) }
if (-not [string]::IsNullOrWhiteSpace($env:TEMP))   { $destinos += (Join-Path $env:TEMP   $nombre) }
$destinos = @($destinos | Select-Object -Unique)

$xl = $null
$guardado = $null

try {
    Write-Host "[1/6] Abriendo Excel..." -ForegroundColor Cyan
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = (-not $Invisible)
    $xl.DisplayAlerts = $false     # solo mientras se construye

    Write-Host "[2/6] Abriendo el CSV..." -ForegroundColor Cyan
    $wb = $xl.Workbooks.Open($CsvPath)
    $ws = $wb.Worksheets.Item(1)

    $filas    = $ws.UsedRange.Rows.Count
    $columnas = $ws.UsedRange.Columns.Count
    Write-Host "      $filas filas x $columnas columnas"

    $encabezados = @()
    for ($c = 1; $c -le $columnas; $c++) { $encabezados += [string]$ws.Cells(1, $c).Value2 }
    Write-Host "      Encabezados: $($encabezados -join ', ')"

    foreach ($requerido in @('Rango horario', 'Received')) {
        if ($encabezados -notcontains $requerido) {
            throw "El CSV no tiene la columna '$requerido'. Encabezados encontrados: $($encabezados -join ', ')"
        }
    }

    Write-Host "[3/6] Creando la hoja y la cache de la dinamica..." -ForegroundColor Cyan
    $rng     = $ws.Range($ws.Cells(1, 1), $ws.Cells($filas, $columnas))
    $srcAddr = "'" + $ws.Name + "'!" + $rng.Address($true, $true, 1)

    $wsP = $wb.Worksheets.Add()
    $wsP.Name = "Correos por hora"

    $pc = $wb.PivotCaches().Create(1, $srcAddr)          # 1 = xlDatabase
    $pt = $pc.CreatePivotTable($wsP.Range("A3"), "ptCorreosPorHora")

    Write-Host "[4/6] Colocando los campos..." -ForegroundColor Cyan
    $pt.PivotFields("Rango horario").Orientation = 1     # xlRowField
    if ($encabezados -contains 'Folder') {
        $pt.PivotFields("Folder").Orientation = 2        # xlColumnField
    }
    [void]$pt.AddDataField($pt.PivotFields("Received"), "Correos", -4112)  # xlCount
    $pt.RowGrand    = $true
    $pt.ColumnGrand = $true

    $wsP.Range("A1").Value2   = $Titulo
    $wsP.Range("A1").Font.Bold = $true
    $wsP.Range("A1").Font.Size = 13

    Write-Host "[5/6] Agregando el grafico..." -ForegroundColor Cyan
    $co = $wsP.ChartObjects().Add(320, 20, 560, 320)
    $co.Chart.SetSourceData($pt.TableRange1)
    $co.Chart.ChartType = 51                             # xlColumnClustered
    $co.Chart.HasTitle  = $true
    $co.Chart.ChartTitle.Text = "Correos por hora"

    $wsP.Columns("A:Z").AutoFit() | Out-Null
    $wsP.Activate()

    Write-Host "[6/6] Guardando..." -ForegroundColor Cyan
    # A partir de aca los avisos SI se muestran: un SaveAs silenciado es
    # justamente lo que enmascara el fallo.
    $xl.DisplayAlerts = $true

    foreach ($destino in $destinos) {
        try {
            if (Test-Path -LiteralPath $destino) { Remove-Item -LiteralPath $destino -Force }
            $wb.SaveAs($destino, 51)                     # 51 = xlOpenXMLWorkbook
            if (Test-Path -LiteralPath $destino) {
                $guardado = $destino
                break
            }
            Write-Warning "      Excel no dio error pero '$destino' no existe. Probando otra carpeta..."
        }
        catch {
            Write-Warning "      No se pudo guardar en '$destino': $($_.Exception.Message)"
        }
    }

    if (-not $guardado) {
        throw "No se pudo guardar el .xlsx en ninguna de estas rutas:`n  " + ($destinos -join "`n  ")
    }

    $wb.Close($false)
    $xl.Quit()

    Write-Host ""
    Write-Host "LISTO -> $guardado" -ForegroundColor Green
    Invoke-Item $guardado
}
catch {
    Write-Host ""
    Write-Warning "Fallo: $($_.Exception.Message)"
    if ($xl) {
        Write-Warning "Excel queda abierto para que puedas ver el estado y guardar a mano."
        try { $xl.Visible = $true } catch { }
    }
    throw
}
finally {
    if ($guardado -and $xl) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
