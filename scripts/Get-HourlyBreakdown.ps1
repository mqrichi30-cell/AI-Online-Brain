<#
.SYNOPSIS
    Genera el desglose de correos por hora del dia a partir del CSV exportado,
    sin abrir Excel ni depender de automatizacion COM.

.DESCRIPTION
    Alternativa a la tabla dinamica cuando Excel bloquea el guardado (por
    ejemplo, con etiquetas de confidencialidad obligatorias de Purview/AIP,
    que abren un dialogo modal y dejan SaveAs esperando indefinidamente).

    Produce lo mismo que la dinamica -- conteo por franja horaria, abierto por
    carpeta -- en tres formatos:

      * Tabla en consola con histograma.
      * CSV cruzado (una fila por franja, una columna por carpeta) listo para
        abrir en Excel: ya es el resultado, no hace falta dinamica.
      * Opcionalmente un desglose por dia de la semana.

.PARAMETER CsvPath
    CSV generado por Export-And-PivotByHour.ps1 (necesita las columnas
    'Rango horario' o 'Hora', y 'Folder').

.PARAMETER OutputPath
    CSV cruzado de salida. Por defecto, el de entrada con sufijo -PorHora.

.PARAMETER PorDiaSemana
    Ademas del desglose por hora, genera el cruce hora x dia de la semana.

.EXAMPLE
    .\Get-HourlyBreakdown.ps1

.EXAMPLE
    .\Get-HourlyBreakdown.ps1 -CsvPath "C:\Users\marin.c\Downloads\CorreosUltimos6Meses.csv" -PorDiaSemana
#>

[CmdletBinding()]
param(
    [string] $CsvPath    = "$env:USERPROFILE\Downloads\CorreosUltimos6Meses.csv",
    [string] $OutputPath = "",
    [switch] $PorDiaSemana
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $CsvPath)) { throw "No existe el CSV '$CsvPath'." }
$CsvPath = (Resolve-Path -LiteralPath $CsvPath).Path

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Split-Path -Path $CsvPath -Parent) `
                            ([IO.Path]::GetFileNameWithoutExtension($CsvPath) + "-PorHora.csv")
}

Write-Host "Leyendo $CsvPath ..." -ForegroundColor Cyan
$datos = @(Import-Csv -LiteralPath $CsvPath)
if ($datos.Count -eq 0) { throw "El CSV no tiene filas." }

$columnas = $datos[0].PSObject.Properties.Name

# La hora puede venir ya calculada; si no, se deriva de Received.
function Get-Hora {
    param($fila)

    if ($columnas -contains 'Hora' -and -not [string]::IsNullOrWhiteSpace($fila.Hora)) {
        $n = 0
        if ([int]::TryParse([string]$fila.Hora, [ref]$n) -and $n -ge 0 -and $n -le 23) { return $n }
    }
    if ($columnas -contains 'Received') {
        $d = [datetime]::MinValue
        if ([datetime]::TryParse([string]$fila.Received, [ref]$d)) { return $d.Hour }
    }
    return $null
}

$conHora  = [System.Collections.Generic.List[object]]::new()
$sinHora  = 0
foreach ($fila in $datos) {
    $h = Get-Hora -fila $fila
    if ($null -eq $h) { $sinHora++; continue }
    $conHora.Add([pscustomobject]@{
        Hora   = [int]$h
        Rango  = ('{0:00}:00 - {0:00}:59' -f [int]$h)
        Folder = $(if ($columnas -contains 'Folder') { $fila.Folder } else { '(todas)' })
        Dia    = $(if ($columnas -contains 'DiaSemana') { $fila.DiaSemana } else { '' })
    })
}

if ($conHora.Count -eq 0) { throw "No se pudo determinar la hora de ninguna fila." }
if ($sinHora -gt 0) { Write-Warning "$sinHora fila(s) sin hora reconocible quedaron fuera del desglose." }

$carpetas = @($conHora.Folder | Sort-Object -Unique)
$total    = $conHora.Count

# --- Cruce hora x carpeta -------------------------------------------------

$filasSalida = [System.Collections.Generic.List[object]]::new()
$porHora     = $conHora | Group-Object Hora | Sort-Object { [int]$_.Name }
$maximo      = ($porHora | Measure-Object Count -Maximum).Maximum

Write-Host ""
Write-Host "Correos por hora del dia (24h) - $total correos" -ForegroundColor Cyan
Write-Host ("{0,-16} {1,7} {2,7}  {3}" -f 'Franja', 'Correos', '%', 'Distribucion')
Write-Host ("-" * 78)

foreach ($grupo in $porHora) {
    $hora = [int]$grupo.Name
    $pct  = 100 * $grupo.Count / $total
    $fila = [ordered]@{
        'Rango horario' = ('{0:00}:00 - {0:00}:59' -f $hora)
        'Hora'          = $hora
    }
    foreach ($carpeta in $carpetas) {
        $fila[$carpeta] = @($grupo.Group | Where-Object { $_.Folder -eq $carpeta }).Count
    }
    $fila['Total']      = $grupo.Count
    $fila['Porcentaje'] = [math]::Round($pct, 2)
    $filasSalida.Add([pscustomobject]$fila)

    $barra = '#' * [math]::Max(1, [math]::Round(40 * $grupo.Count / $maximo))
    Write-Host ("{0,-16} {1,7} {2,6:N1}%  {3}" -f $fila['Rango horario'], $grupo.Count, $pct, $barra)
}

# Fila de totales, para que el CSV se lea solo sin recalcular nada.
$totales = [ordered]@{ 'Rango horario' = 'TOTAL'; 'Hora' = '' }
foreach ($carpeta in $carpetas) {
    $totales[$carpeta] = @($conHora | Where-Object { $_.Folder -eq $carpeta }).Count
}
$totales['Total']      = $total
$totales['Porcentaje'] = 100
$filasSalida.Add([pscustomobject]$totales)

$encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$filasSalida | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding $encoding -Force

# --- Franja de mayor volumen ---------------------------------------------

$pico     = $porHora | Sort-Object Count -Descending | Select-Object -First 1
$horaPico = [int]$pico.Name
$jornada  = @($conHora | Where-Object { $_.Hora -ge 6 -and $_.Hora -lt 18 }).Count

Write-Host ""
Write-Host ("Pico          : {0:00}:00 con {1} correos ({2:N1}%)" -f $horaPico, $pico.Count, (100 * $pico.Count / $total))
Write-Host ("Entre 06 y 18 : {0} correos ({1:N1}%)" -f $jornada, (100 * $jornada / $total))
Write-Host ("Fuera de esa franja: {0} correos ({1:N1}%)" -f ($total - $jornada), (100 * ($total - $jornada) / $total))

# --- Opcional: hora x dia de la semana ------------------------------------

if ($PorDiaSemana) {
    $conDia = @($conHora | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Dia) })
    if ($conDia.Count -eq 0) {
        Write-Warning "El CSV no trae la columna 'DiaSemana'; se omite ese desglose."
    }
    else {
        $dias    = @($conDia.Dia | Sort-Object -Unique)
        $salida2 = [System.Collections.Generic.List[object]]::new()
        foreach ($grupo in ($conDia | Group-Object Hora | Sort-Object { [int]$_.Name })) {
            $fila = [ordered]@{ 'Rango horario' = ('{0:00}:00 - {0:00}:59' -f [int]$grupo.Name) }
            foreach ($dia in $dias) {
                $fila[$dia] = @($grupo.Group | Where-Object { $_.Dia -eq $dia }).Count
            }
            $fila['Total'] = $grupo.Count
            $salida2.Add([pscustomobject]$fila)
        }
        $rutaDias = Join-Path (Split-Path -Path $OutputPath -Parent) `
                              ([IO.Path]::GetFileNameWithoutExtension($OutputPath) + "-PorDia.csv")
        $salida2 | Export-Csv -Path $rutaDias -NoTypeInformation -Encoding $encoding -Force
        Write-Host ""
        Write-Host "Cruce hora x dia: $rutaDias"
    }
}

Write-Host ""
Write-Host "CSV con el desglose: $OutputPath" -ForegroundColor Green
