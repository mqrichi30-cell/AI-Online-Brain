<#
.SYNOPSIS
    Exporta a CSV los correos recibidos en los ultimos N dias desde carpetas
    concretas de un buzon compartido de Microsoft 365, leyendo directamente de
    Exchange Online via Microsoft Graph (sin Outlook COM ni cache OST).

.DESCRIPTION
    Diferencias clave frente al enfoque "traer todo y filtrar en el cliente":

      * El filtro de fecha se aplica en el SERVIDOR ($filter=receivedDateTime ge ...),
        asi Graph solo devuelve los correos del rango pedido en lugar del historico
        completo de la carpeta.
      * Se piden solo las propiedades necesarias ($select), reduciendo el tamano de
        cada respuesta de forma drastica.
      * Paginacion manual con $top=999 siguiendo @odata.nextLink, con reintentos por
        pagina (429 / 5xx respetando el header Retry-After).
      * Acumulacion en List[object] en vez de "$array += ...", que en PowerShell es
        O(n^2) porque recrea el arreglo completo en cada iteracion.
      * Si el CSV de destino esta abierto en Excel, escribe en un archivo alternativo
        con marca de tiempo en vez de perder todo el trabajo.

.PARAMETER Mailbox
    UPN del buzon compartido a leer.

.PARAMETER Days
    Ventana de dias hacia atras. Por defecto 15.

.PARAMETER FromMidnight
    Toma como corte la medianoche (hora local) de hace $Days dias en vez de la hora
    exacta actual menos $Days dias.

.PARAMETER ParentFolderPath
    Ruta de carpetas, desde Inbox, hasta la carpeta que contiene las carpetas destino.

.PARAMETER TargetFolderName
    Nombres de las carpetas a exportar dentro de ParentFolderPath.

.PARAMETER IncludeSubfolders
    Incluye tambien, de forma recursiva, las subcarpetas de cada carpeta destino.

.PARAMETER OutputPath
    Ruta del CSV de salida.

.PARAMETER ListFoldersOnly
    No exporta nada: solo imprime el arbol de carpetas del buzon. Util cuando un
    nombre de carpeta no coincide (acentos, espacios dobles, guiones distintos).

.EXAMPLE
    .\Export-SharedMailboxEmails.ps1

.EXAMPLE
    .\Export-SharedMailboxEmails.ps1 -Days 30 -OutputPath "C:\Temp\Ultimos30.csv"

.EXAMPLE
    .\Export-SharedMailboxEmails.ps1 -ListFoldersOnly

.NOTES
    Permisos necesarios: Mail.Read.Shared (delegado) sobre el buzon compartido.
        Connect-MgGraph -Scopes "Mail.Read.Shared"
#>

[CmdletBinding()]
param(
    [string]   $Mailbox          = "pgcustservw2.im@pg.com",
    [int]      $Days             = 15,
    [switch]   $FromMidnight,
    [string[]] $ParentFolderPath = @("Marín, Cristhofer - AWG + Wakefern"),
    [string[]] $TargetFolderName = @("AWG Complete", "Wakefern Complete"),
    [switch]   $IncludeSubfolders,
    [string]   $OutputPath       = "$env:USERPROFILE\Downloads\Last15DaysEmails.csv",
    [switch]   $ListFoldersOnly
)

$ErrorActionPreference = 'Stop'
$graphBase = 'https://graph.microsoft.com/v1.0'

# ---------------------------------------------------------------------------
# Conexion
# ---------------------------------------------------------------------------

function Assert-GraphConnection {
    $context = $null
    try { $context = Get-MgContext } catch { }

    if (-not $context) {
        Write-Host "No hay sesion de Graph activa. Conectando..." -ForegroundColor Yellow
        # -NoWelcome solo existe en el SDK v2.
        if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('NoWelcome')) {
            Connect-MgGraph -Scopes "Mail.Read.Shared" -NoWelcome
        } else {
            Connect-MgGraph -Scopes "Mail.Read.Shared"
        }
        $context = Get-MgContext
    }

    if (-not $context) {
        throw "No fue posible establecer la conexion con Microsoft Graph."
    }

    $hasScope = $context.Scopes | Where-Object { $_ -in @('Mail.Read.Shared', 'Mail.ReadWrite.Shared', 'Mail.Read', 'Mail.ReadWrite') }
    if (-not $hasScope) {
        Write-Warning "La sesion actual no incluye un scope de lectura de correo. Si falla, reconecta con: Connect-MgGraph -Scopes 'Mail.Read.Shared'"
    }

    Write-Host "Conectado como: $($context.Account)" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Llamada HTTP con reintentos (throttling 429 y errores transitorios 5xx)
# ---------------------------------------------------------------------------

function Invoke-GraphGet {
    param(
        [Parameter(Mandatory)][string] $Uri,
        [int] $MaxAttempts = 5
    )

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

            $isRetryable = @(429, 500, 502, 503, 504) -contains $status
            if (-not $isRetryable -or $attempt -ge $MaxAttempts) { throw }

            # Graph indica cuanto esperar cuando aplica throttling; respetarlo.
            $wait = 0
            try {
                $retryAfter = $_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1
                if ($retryAfter) { $wait = [int]$retryAfter }
            } catch { }
            if ($wait -le 0) { $wait = [math]::Min(60, [math]::Pow(2, $attempt)) }

            Write-Warning "Graph respondio $status. Reintento $attempt/$MaxAttempts en $wait s..."
            Start-Sleep -Seconds $wait
        }
    }
}

# ---------------------------------------------------------------------------
# Carpetas
# ---------------------------------------------------------------------------

function Get-ChildFolder {
    param(
        [Parameter(Mandatory)][string] $ParentId
    )

    $folders = [System.Collections.Generic.List[object]]::new()
    # $top=100 es el maximo practico para mailFolders; el numero de subcarpetas es
    # pequeno, asi que se filtra por nombre en el cliente y se evitan problemas de
    # codificacion con acentos y comas en $filter.
    $uri = "$graphBase/users/$Mailbox/mailFolders/$ParentId/childFolders?`$top=100&`$select=id,displayName,totalItemCount,childFolderCount"

    while ($uri) {
        $page = Invoke-GraphGet -Uri $uri
        foreach ($f in $page.value) { $folders.Add($f) }
        $uri = $page.'@odata.nextLink'
    }

    return $folders
}

function Resolve-FolderPath {
    param(
        [Parameter(Mandatory)][string]   $RootId,
        [Parameter(Mandatory)][string[]] $Path
    )

    $currentId   = $RootId
    $currentName = 'Inbox'

    foreach ($segment in $Path) {
        $children = @(Get-ChildFolder -ParentId $currentId)
        # Comparacion tolerante: ignora mayusculas y espacios repetidos/extremos.
        $normalized = ($segment -replace '\s+', ' ').Trim()
        $match = $children | Where-Object {
            (($_.displayName -replace '\s+', ' ').Trim()) -ieq $normalized
        } | Select-Object -First 1

        if (-not $match) {
            $available = ($children.displayName | ForEach-Object { "  - $_" }) -join [Environment]::NewLine
            throw "No se encontro la carpeta '$segment' dentro de '$currentName'.`nCarpetas disponibles:`n$available"
        }

        $currentId   = $match.id
        $currentName = $match.displayName
    }

    return [pscustomobject]@{ Id = $currentId; DisplayName = $currentName }
}

function Get-FolderTree {
    param(
        [Parameter(Mandatory)][string] $ParentId,
        [int] $Depth = 0
    )

    foreach ($f in Get-ChildFolder -ParentId $ParentId) {
        Write-Host ("{0}{1}  [{2} items]" -f ('  ' * $Depth), $f.displayName, $f.totalItemCount)
        if ($f.childFolderCount -gt 0) {
            Get-FolderTree -ParentId $f.id -Depth ($Depth + 1)
        }
    }
}

function Expand-FolderRecursive {
    param([Parameter(Mandatory)][object] $Folder)

    $list = [System.Collections.Generic.List[object]]::new()
    $list.Add($Folder)

    if ($Folder.childFolderCount -gt 0) {
        foreach ($child in Get-ChildFolder -ParentId $Folder.id) {
            foreach ($descendant in (Expand-FolderRecursive -Folder $child)) {
                $list.Add($descendant)
            }
        }
    }

    return $list
}

# ---------------------------------------------------------------------------
# Mensajes
# ---------------------------------------------------------------------------

function Get-FolderMessagesSince {
    param(
        [Parameter(Mandatory)][string]   $FolderId,
        [Parameter(Mandatory)][string]   $FolderName,
        [Parameter(Mandatory)][datetime] $CutoffUtc
    )

    # Formato exigido por Graph: ISO 8601 en UTC, sin comillas alrededor.
    $cutoffLiteral = $CutoffUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $filter  = [uri]::EscapeDataString("receivedDateTime ge $cutoffLiteral")
    $select  = [uri]::EscapeDataString("receivedDateTime,subject,from,sender,toRecipients,ccRecipients,hasAttachments,isRead,importance,conversationId,internetMessageId,webLink")
    $orderBy = [uri]::EscapeDataString("receivedDateTime desc")

    # $top=999 es el maximo por pagina para /messages: minimiza el numero de
    # round-trips frente al valor por defecto de 10.
    $base = "$graphBase/users/$Mailbox/mailFolders/$FolderId/messages?`$filter=$filter&`$select=$select&`$top=999"
    $uri  = "$base&`$orderby=$orderBy"

    $messages = [System.Collections.Generic.List[object]]::new()
    $page     = 0

    while ($uri) {
        $page++
        Write-Progress -Activity "Leyendo '$FolderName'" -Status "Pagina $page - $($messages.Count) correos" -PercentComplete -1

        try {
            $response = Invoke-GraphGet -Uri $uri
        }
        catch {
            # Algunos buzones rechazan $filter combinado con $orderby
            # (InefficientFilter). El orden final se hace en el cliente igualmente.
            if ($page -eq 1 -and $uri -ne $base) {
                Write-Warning "Graph rechazo el ordenamiento en servidor para '$FolderName'. Reintentando sin `$orderby."
                $uri = $base
                $page = 0
                continue
            }
            throw
        }

        foreach ($m in $response.value) { $messages.Add($m) }
        $uri = $response.'@odata.nextLink'
    }

    Write-Progress -Activity "Leyendo '$FolderName'" -Completed
    return $messages
}

function Get-AddressList {
    param($Recipients)

    if (-not $Recipients) { return '' }
    return (($Recipients | ForEach-Object { $_.emailAddress.address } | Where-Object { $_ }) -join '; ')
}

# ---------------------------------------------------------------------------
# Salida
# ---------------------------------------------------------------------------

function Export-ResultsToCsv {
    param(
        [Parameter(Mandatory)][object[]] $Rows,
        [Parameter(Mandatory)][string]   $Path
    )

    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    # PowerShell 7 escribe UTF-8 sin BOM y Excel rompe los acentos; forzar BOM.
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }

    try {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding $encoding -Force
        return $Path
    }
    catch {
        # Caso tipico: el CSV esta abierto en Excel y el archivo queda bloqueado.
        $isLocked = $false
        $ex = $_.Exception
        while ($ex) {
            if ($ex -is [System.IO.IOException] -or $ex -is [System.UnauthorizedAccessException]) { $isLocked = $true; break }
            $ex = $ex.InnerException
        }
        if (-not $isLocked) { throw }

        $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
        $fallback = Join-Path (Split-Path -Path $Path -Parent) `
                              ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($Path), $stamp, [IO.Path]::GetExtension($Path))
        Write-Warning "El archivo '$Path' esta en uso (probablemente abierto en Excel). Guardando en '$fallback'."
        $Rows | Export-Csv -Path $fallback -NoTypeInformation -Encoding $encoding -Force
        return $fallback
    }
}

# ---------------------------------------------------------------------------
# Ejecucion
# ---------------------------------------------------------------------------

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Assert-GraphConnection

$inbox = Invoke-GraphGet -Uri "$graphBase/users/$Mailbox/mailFolders/inbox?`$select=id,displayName"

if ($ListFoldersOnly) {
    Write-Host ""
    Write-Host "Arbol de carpetas bajo Inbox de $Mailbox" -ForegroundColor Cyan
    Write-Host "----------------------------------------"
    Get-FolderTree -ParentId $inbox.id
    return
}

$cutoffUtc = if ($FromMidnight) {
    (Get-Date).Date.AddDays(-$Days).ToUniversalTime()
} else {
    (Get-Date).AddDays(-$Days).ToUniversalTime()
}

Write-Host ""
Write-Host "Buzon : $Mailbox"
Write-Host "Desde : $($cutoffUtc.ToString('yyyy-MM-dd HH:mm:ss')) UTC ($Days dias)"
Write-Host ""

$parent = Resolve-FolderPath -RootId $inbox.id -Path $ParentFolderPath

$children = @(Get-ChildFolder -ParentId $parent.Id)
$folders  = [System.Collections.Generic.List[object]]::new()

foreach ($name in $TargetFolderName) {
    $normalized = ($name -replace '\s+', ' ').Trim()
    $match = $children | Where-Object {
        (($_.displayName -replace '\s+', ' ').Trim()) -ieq $normalized
    } | Select-Object -First 1

    if (-not $match) {
        Write-Warning "No se encontro la carpeta '$name' dentro de '$($parent.DisplayName)'. Se omite."
        continue
    }

    if ($IncludeSubfolders) {
        foreach ($f in (Expand-FolderRecursive -Folder $match)) { $folders.Add($f) }
    } else {
        $folders.Add($match)
    }
}

if ($folders.Count -eq 0) {
    throw "No se resolvio ninguna carpeta destino. Ejecuta el script con -ListFoldersOnly para ver los nombres exactos."
}

$results = [System.Collections.Generic.List[object]]::new()

foreach ($folder in $folders) {
    Write-Host "Leyendo: $($folder.displayName)" -ForegroundColor Cyan

    $messages = @(Get-FolderMessagesSince -FolderId $folder.id -FolderName $folder.displayName -CutoffUtc $cutoffUtc)

    foreach ($m in $messages) {
        $receivedUtc = [datetime]::Parse($m.receivedDateTime, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)

        # from puede venir vacio en borradores o correos de sistema; sender es el respaldo.
        $originator = if ($m.from) { $m.from } elseif ($m.sender) { $m.sender } else { $null }
        $fromAddress = if ($originator) { $originator.emailAddress.address } else { '' }
        $fromName    = if ($originator) { $originator.emailAddress.name }    else { '' }

        $results.Add([pscustomobject]@{
            SortKey           = $receivedUtc
            Received          = $receivedUtc.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
            ReceivedUtc       = $receivedUtc.ToString('yyyy-MM-dd HH:mm:ss')
            Subject           = $m.subject
            From              = $fromAddress
            FromName          = $fromName
            To                = Get-AddressList -Recipients $m.toRecipients
            Cc                = Get-AddressList -Recipients $m.ccRecipients
            HasAttachments    = $m.hasAttachments
            IsRead            = $m.isRead
            Importance        = $m.importance
            Folder            = $folder.displayName
            ConversationId    = $m.conversationId
            InternetMessageId = $m.internetMessageId
            WebLink           = $m.webLink
        })
    }

    Write-Host "  Encontrados: $($messages.Count)" -ForegroundColor DarkGray
}

$stopwatch.Stop()

if ($results.Count -eq 0) {
    Write-Warning "No se encontraron correos en el rango solicitado. No se genero el CSV."
    return
}

$rows = @($results | Sort-Object SortKey | Select-Object -Property * -ExcludeProperty SortKey)
$file = Export-ResultsToCsv -Rows $rows -Path $OutputPath

Write-Host ""
Write-Host "=================" -ForegroundColor Green
Write-Host "FINALIZADO"       -ForegroundColor Green
Write-Host "Total    : $($results.Count)"
Write-Host "Carpetas : $($folders.Count)"
Write-Host "Tiempo   : $([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) s"
Write-Host "Archivo  : $file"
Write-Host "=================" -ForegroundColor Green
