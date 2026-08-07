<#
.SYNOPSIS
    Exporta los correos de los ultimos N meses desde carpetas concretas de un buzon
    compartido de Microsoft 365 y arma, en el mismo paso, un Excel con la tabla
    dinamica de correos por hora del dia.

.DESCRIPTION
    Une los dos pasos que antes eran manuales:

      1. Lectura via Microsoft Graph con filtro de fecha en el SERVIDOR
         ($filter=receivedDateTime ge ...), $select acotado y paginacion a
         $top=999 siguiendo @odata.nextLink, con reintentos en 429 / 5xx.
      2. Automatizacion de Excel (COM) para crear la tabla dinamica y el grafico
         de barras, y guardar el .xlsx.

    La hora se calcula en PowerShell a partir del DateTime real que devuelve
    Graph, no parseando texto en Excel: no depende del formato regional ni de
    como Excel interprete la columna.

    Si Excel no esta disponible o la automatizacion falla, el CSV igual queda
    escrito y el resumen por hora se imprime en consola.

.PARAMETER Months
    Meses hacia atras. Por defecto 6.

.EXAMPLE
    .\Export-And-PivotByHour.ps1

.EXAMPLE
    .\Export-And-PivotByHour.ps1 -Months 12

.NOTES
    Permisos: Mail.Read.Shared (delegado) sobre el buzon compartido.
#>

[CmdletBinding()]
param(
    [string] $Mailbox    = "pgcustservw2.im@pg.com",
    [int]    $Months     = 6,
    [string] $OutputDir  = "$env:USERPROFILE\Downloads",
    [string] $BaseName   = "CorreosUltimos6Meses",
    [switch] $SkipExcel
)

$ErrorActionPreference = 'Stop'
$graphBase = 'https://graph.microsoft.com/v1.0'
$csvPath   = Join-Path $OutputDir "$BaseName.csv"
$xlsxPath  = Join-Path $OutputDir "$BaseName.xlsx"

# ---------------------------------------------------------------------------
# Graph: peticion con reintentos
# ---------------------------------------------------------------------------

function Invoke-GraphGet {
    param([Parameter(Mandatory)][string] $Uri, [int] $MaxAttempts = 5)

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $status = 0
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            if ($status -eq 0 -and $_.Exception.Message -match '\b(429|500|502|503|504)\b') {
                $status = [int]$Matches[1]
            }
            if (-not (@(429, 500, 502, 503, 504) -contains $status) -or $attempt -ge $MaxAttempts) { throw }

            $wait = 0
            try {
                $ra = $_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1
                if ($ra) { $wait = [int]$ra }
            } catch { }
            if ($wait -le 0) { $wait = [math]::Min(60, [math]::Pow(2, $attempt)) }

            Write-Warning "Graph respondio $status. Reintento $attempt/$MaxAttempts en $wait s..."
            Start-Sleep -Seconds $wait
        }
    }
}

function Get-ChildFolder {
    param([Parameter(Mandatory)][string] $ParentId)

    $all = [System.Collections.Generic.List[object]]::new()
    $uri = "$graphBase/users/$Mailbox/mailFolders/$ParentId/childFolders?`$top=100&`$select=id,displayName,totalItemCount"
    while ($uri) {
        $page = Invoke-GraphGet -Uri $uri
        foreach ($f in $page.value) { $all.Add($f) }
        $uri = $page.'@odata.nextLink'
    }
    return $all
}

# ---------------------------------------------------------------------------
# 1. Exportacion
# ---------------------------------------------------------------------------

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

if (-not (Get-MgContext)) {
    Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "Mail.Read.Shared"
}

$cutoffUtc     = (Get-Date).AddMonths(-$Months).ToUniversalTime()
$cutoffLiteral = $cutoffUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host ""
Write-Host "Buzon : $Mailbox"
Write-Host "Desde : $($cutoffUtc.ToString('yyyy-MM-dd HH:mm:ss')) UTC ($Months meses)"
Write-Host ""

$inbox = Invoke-GraphGet -Uri "$graphBase/users/$Mailbox/mailFolders/inbox?`$select=id"

$parent = @(Get-ChildFolder -ParentId $inbox.id) |
          Where-Object { $_.displayName -like "*AWG*Wakefern*" } |
          Select-Object -First 1

if (-not $parent) {
    Write-Host "Subcarpetas encontradas bajo Inbox:" -ForegroundColor Yellow
    (Get-ChildFolder -ParentId $inbox.id).displayName | ForEach-Object { Write-Host "  - $_" }
    throw "No encontre la carpeta que contiene AWG + Wakefern. Arriba estan los nombres reales."
}

$targets = @(Get-ChildFolder -ParentId $parent.id) |
           Where-Object { $_.displayName -eq "AWG Complete" -or $_.displayName -eq "Wakefern Complete" }

if (-not $targets) {
    throw "No encontre 'AWG Complete' ni 'Wakefern Complete' dentro de '$($parent.displayName)'."
}

$select = [uri]::EscapeDataString("receivedDateTime,subject,from,sender,toRecipients,ccRecipients,hasAttachments,isRead")
$filter = [uri]::EscapeDataString("receivedDateTime ge $cutoffLiteral")

$rows = [System.Collections.Generic.List[object]]::new()

foreach ($folder in $targets) {
    Write-Host "Leyendo: $($folder.displayName)" -ForegroundColor Cyan
    $uri  = "$graphBase/users/$Mailbox/mailFolders/$($folder.id)/messages?`$filter=$filter&`$select=$select&`$top=999"
    $page = 0
    $n    = 0

    while ($uri) {
        $page++
        Write-Progress -Activity "Leyendo '$($folder.displayName)'" -Status "Pagina $page - $n correos" -PercentComplete -1
        $response = Invoke-GraphGet -Uri $uri

        foreach ($m in $response.value) {
            $utc   = [datetime]::Parse($m.receivedDateTime, [cultureinfo]::InvariantCulture,
                                       [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor
                                       [System.Globalization.DateTimeStyles]::AssumeUniversal)
            $local = $utc.ToLocalTime()
            # La hora sale del DateTime real, no de parsear texto en Excel.
            $hora  = $local.Hour
            $orig  = if ($m.from) { $m.from } elseif ($m.sender) { $m.sender } else { $null }

            $rows.Add([pscustomobject]@{
                SortKey        = $utc
                Received       = $local.ToString('yyyy-MM-dd HH:mm:ss')
                Fecha          = $local.ToString('yyyy-MM-dd')
                Hora           = $hora
                'Rango horario'= ('{0:00}:00 - {0:00}:59' -f $hora)
                DiaSemana      = $local.ToString('dddd')
                Subject        = $m.subject
                From           = $(if ($orig) { $orig.emailAddress.address })
                To             = (($m.toRecipients | ForEach-Object { $_.emailAddress.address }) -join '; ')
                Cc             = (($m.ccRecipients | ForEach-Object { $_.emailAddress.address }) -join '; ')
                HasAttachments = $m.hasAttachments
                IsRead         = $m.isRead
                Folder         = $folder.displayName
            })
            $n++
        }

        $uri = $response.'@odata.nextLink'
    }

    Write-Progress -Activity "Leyendo '$($folder.displayName)'" -Completed
    Write-Host "  Encontrados: $n" -ForegroundColor DarkGray
}

if ($rows.Count -eq 0) {
    Write-Warning "No se encontraron correos en el rango solicitado. No se genero nada."
    return
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$export   = @($rows | Sort-Object SortKey | Select-Object -Property * -ExcludeProperty SortKey)

try {
    $export | Export-Csv -Path $csvPath -NoTypeInformation -Encoding $encoding -Force
}
catch {
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath  = Join-Path $OutputDir "${BaseName}_$stamp.csv"
    $xlsxPath = Join-Path $OutputDir "${BaseName}_$stamp.xlsx"
    Write-Warning "El CSV estaba en uso. Guardando en '$csvPath'."
    $export | Export-Csv -Path $csvPath -NoTypeInformation -Encoding $encoding -Force
}

# ---------------------------------------------------------------------------
# 2. Resumen por hora en consola (sirve aunque Excel falle)
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Correos por hora (formato 24h)" -ForegroundColor Cyan
Write-Host "------------------------------"
$porHora = $rows | Group-Object Hora | Sort-Object { [int]$_.Name }
$maximo  = ($porHora | Measure-Object Count -Maximum).Maximum
foreach ($g in $porHora) {
    $barra = '#' * [math]::Max(1, [math]::Round(40 * $g.Count / $maximo))
    Write-Host ("{0:00}:00  {1,6}  {2}" -f [int]$g.Name, $g.Count, $barra)
}

# ---------------------------------------------------------------------------
# 3. Excel: tabla dinamica + grafico
# ---------------------------------------------------------------------------

$excelOk = $false

if (-not $SkipExcel) {
    $xl = $null
    try {
        Write-Host ""
        Write-Host "Armando la tabla dinamica en Excel..." -ForegroundColor Cyan

        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false

        $wb = $xl.Workbooks.Open($csvPath)
        $ws = $wb.Worksheets.Item(1)

        $lastRow = $ws.UsedRange.Rows.Count
        $lastCol = $ws.UsedRange.Columns.Count
        $rng     = $ws.Range($ws.Cells(1, 1), $ws.Cells($lastRow, $lastCol))
        $srcAddr = "'" + $ws.Name + "'!" + $rng.Address($true, $true, 1)

        $wsP = $wb.Worksheets.Add()
        $wsP.Name = "Correos por hora"

        # 1 = xlDatabase
        $pc = $wb.PivotCaches().Create(1, $srcAddr)
        $pt = $pc.CreatePivotTable($wsP.Range("A3"), "ptCorreosPorHora")

        # 1 = xlRowField, 2 = xlColumnField, -4112 = xlCount
        $pt.PivotFields("Rango horario").Orientation = 1
        $pt.PivotFields("Rango horario").Position    = 1
        $pt.PivotFields("Folder").Orientation        = 2
        $pt.PivotFields("Folder").Position           = 1
        [void]$pt.AddDataField($pt.PivotFields("Received"), "Correos", -4112)
        $pt.RowGrand    = $true
        $pt.ColumnGrand = $true

        $wsP.Range("A1").Value2 = "Correos recibidos por hora del dia (ultimos $Months meses, 24h)"
        $wsP.Range("A1").Font.Bold = $true
        $wsP.Range("A1").Font.Size = 13

        # 51 = xlColumnClustered
        $co = $wsP.ChartObjects().Add(320, 20, 560, 320)
        $co.Chart.SetSourceData($pt.TableRange1)
        $co.Chart.ChartType = 51
        $co.Chart.HasTitle = $true
        $co.Chart.ChartTitle.Text = "Correos por hora"

        $wsP.Columns("A:Z").AutoFit() | Out-Null
        $wsP.Activate()

        # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.SaveAs($xlsxPath, 51)
        $wb.Close($false)
        $xl.Quit()
        $excelOk = $true
    }
    catch {
        Write-Warning "No se pudo automatizar Excel: $($_.Exception.Message)"
        Write-Warning "El CSV quedo generado igual; podes armar la dinamica a mano o con la macro VBA."
        if ($xl) { try { $xl.Quit() } catch { } }
    }
    finally {
        if ($xl) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

$stopwatch.Stop()

Write-Host ""
Write-Host "=================" -ForegroundColor Green
Write-Host "FINALIZADO"       -ForegroundColor Green
Write-Host "Correos  : $($rows.Count)"
Write-Host "Carpetas : $($targets.Count)"
Write-Host "Tiempo   : $([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) s"
Write-Host "CSV      : $csvPath"
if ($excelOk) { Write-Host "Excel    : $xlsxPath" }
Write-Host "=================" -ForegroundColor Green

if ($excelOk) { Invoke-Item $xlsxPath }
