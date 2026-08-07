$mailbox = "pgcustservw2.im@pg.com"
$meses = 6
$dir = "$env:USERPROFILE\Downloads"
$csv = "$dir\CorreosUltimos6Meses.csv"
$xlsx = "$dir\CorreosUltimos6Meses.xlsx"
$g = "https://graph.microsoft.com/v1.0"
function Get-G($u) {
  $i = 0
  while ($true) {
    $i++
    try { return Invoke-MgGraphRequest -Method GET -Uri $u -OutputType PSObject -ErrorAction Stop }
    catch {
      $s = 0; try { $s = [int]$_.Exception.Response.StatusCode } catch {}
      if ($s -eq 0 -and $_.Exception.Message -match '\b(429|500|502|503|504)\b') { $s = [int]$Matches[1] }
      if (-not (@(429,500,502,503,504) -contains $s) -or $i -ge 5) { throw }
      $w = 0; try { $w = [int](($_.Exception.Response.Headers.GetValues('Retry-After'))[0]) } catch {}
      if ($w -le 0) { $w = [math]::Min(60, [math]::Pow(2, $i)) }
      Write-Warning "Graph $s. Reintento $i/5 en $w s..."; Start-Sleep -Seconds $w
    }
  }
}
if (-not (Get-MgContext)) { Connect-MgGraph -Scopes "Mail.Read.Shared" }
$cut = (Get-Date).AddMonths(-$meses).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
Write-Host "Desde: $cut UTC ($meses meses)`n"
$inbox = Get-G "$g/users/$mailbox/mailFolders/inbox"
$kids = (Get-G "$g/users/$mailbox/mailFolders/$($inbox.id)/childFolders?`$top=100").value
$parent = $kids | Where-Object { $_.displayName -like "*AWG*Wakefern*" } | Select-Object -First 1
if (-not $parent) { $kids.displayName; throw "No encontre la carpeta padre. Arriba estan los nombres reales." }
$targets = (Get-G "$g/users/$mailbox/mailFolders/$($parent.id)/childFolders?`$top=100").value | Where-Object { $_.displayName -eq "AWG Complete" -or $_.displayName -eq "Wakefern Complete" }
if (-not $targets) { throw "No encontre AWG Complete ni Wakefern Complete dentro de '$($parent.displayName)'." }
$sel = "receivedDateTime,subject,from,sender,toRecipients,ccRecipients,hasAttachments,isRead"
$rows = New-Object System.Collections.Generic.List[object]
foreach ($f in $targets) {
  Write-Host "Leyendo: $($f.displayName)" -ForegroundColor Cyan
  $uri = "$g/users/$mailbox/mailFolders/$($f.id)/messages?`$filter=receivedDateTime%20ge%20$cut&`$select=$sel&`$top=999"
  $n = 0; $p = 0
  while ($uri) {
    $p++
    Write-Progress -Activity $f.displayName -Status "Pagina $p - $n correos" -PercentComplete -1
    $r = Get-G $uri
    foreach ($m in $r.value) {
      $t = ([datetime]$m.receivedDateTime).ToLocalTime()
      $o = if ($m.from) { $m.from } else { $m.sender }
      $rows.Add([pscustomobject]@{
        Received = $t.ToString('yyyy-MM-dd HH:mm:ss')
        Fecha = $t.ToString('yyyy-MM-dd')
        Hora = $t.Hour
        'Rango horario' = ('{0:00}:00 - {0:00}:59' -f $t.Hour)
        DiaSemana = $t.ToString('dddd')
        Subject = $m.subject
        From = $(if ($o) { $o.emailAddress.address })
        To = (($m.toRecipients | ForEach-Object { $_.emailAddress.address }) -join '; ')
        Cc = (($m.ccRecipients | ForEach-Object { $_.emailAddress.address }) -join '; ')
        HasAttachments = $m.hasAttachments
        Folder = $f.displayName })
      $n++
    }
    $uri = $r.'@odata.nextLink'
  }
  Write-Progress -Activity $f.displayName -Completed
  Write-Host "  Encontrados: $n"
}
if ($rows.Count -eq 0) { Write-Warning "Sin correos en el rango."; return }
try { $rows | Sort-Object Received | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 }
catch { $s = Get-Date -Format 'HHmmss'; $csv = "$dir\CorreosUltimos6Meses_$s.csv"; $xlsx = "$dir\CorreosUltimos6Meses_$s.xlsx"; Write-Warning "Archivo en uso, guardo en $csv"; $rows | Sort-Object Received | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 }
Write-Host "`nCorreos por hora (24h)" -ForegroundColor Cyan
$ph = $rows | Group-Object Hora | Sort-Object { [int]$_.Name }
$mx = ($ph | Measure-Object Count -Maximum).Maximum
foreach ($x in $ph) { Write-Host ("{0:00}:00 {1,6}  {2}" -f [int]$x.Name, $x.Count, ('#' * [math]::Max(1, [math]::Round(40 * $x.Count / $mx)))) }
$ok = $false
try {
  Write-Host "`nArmando la tabla dinamica en Excel..." -ForegroundColor Cyan
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open($csv)
  $ws = $wb.Worksheets.Item(1)
  $rg = $ws.Range($ws.Cells(1,1), $ws.Cells($ws.UsedRange.Rows.Count, $ws.UsedRange.Columns.Count))
  $wp = $wb.Worksheets.Add(); $wp.Name = "Correos por hora"
  $pc = $wb.PivotCaches().Create(1, "'" + $ws.Name + "'!" + $rg.Address($true,$true,1))
  $pt = $pc.CreatePivotTable($wp.Range("A3"), "ptHoras")
  $pt.PivotFields("Rango horario").Orientation = 1
  $pt.PivotFields("Folder").Orientation = 2
  [void]$pt.AddDataField($pt.PivotFields("Received"), "Correos", -4112)
  $wp.Range("A1").Value2 = "Correos por hora del dia - ultimos $meses meses (24h)"
  $wp.Range("A1").Font.Bold = $true
  $co = $wp.ChartObjects().Add(320, 20, 560, 320)
  $co.Chart.SetSourceData($pt.TableRange1)
  $co.Chart.ChartType = 51
  $wp.Columns("A:Z").AutoFit() | Out-Null
  $wb.SaveAs($xlsx, 51); $wb.Close($false); $xl.Quit(); $ok = $true
} catch { Write-Warning "Excel fallo: $($_.Exception.Message). El CSV quedo igual en $csv" }
Write-Host "`nLISTO: $($rows.Count) correos" -ForegroundColor Green
Write-Host "CSV  : $csv"
if ($ok) { Write-Host "Excel: $xlsx"; Invoke-Item $xlsx }
