<#
.SYNOPSIS
    Depura copias de seguridad de usuarios en un repositorio de Veeam Backup for Microsoft 365.

.DESCRIPTION
    Lee un CSV con cuentas en formato nombre_usuario@uaesp.gov.co, valida cada cuenta contra el
    contenido del repositorio REPO_CACHE_OLD y elimina datos asociados al usuario: Exchange mailbox,
    archive mailbox, OneDrive y sitios personales/asociados (-Sites).

.NOTAS IMPORTANTES
    - Ejecutar desde Veeam Backup for Microsoft 365 PowerShell Toolkit o una consola PowerShell
      en el servidor de VBO365 con privilegios suficientes.
    - La eliminación es destructiva. Use -WhatIfMode $true para prueba controlada.
    - No funciona contra repositorios de object storage con inmutabilidad habilitada.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepositoryName = "REPO_CACHE_OLD",

    # Ajuste esta ruta a la ubicación real del CSV en el servidor de VBO365.
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = "C:\Scripts\Depurar_Repo_Cache_Old.csv",

    # Si el CSV tiene encabezado, indique aquí el nombre de la columna. Si se deja vacío,
    # el script intentará detectar una columna común o tomará la primera columna.
    [Parameter(Mandatory = $false)]
    [string]$CsvColumn = "",

    # Modo de prueba: muestra qué se eliminaría sin borrar datos.
    [Parameter(Mandatory = $false)]
    [bool]$WhatIfMode = $false,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message, [string]$Color = "Gray")
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message" -ForegroundColor $Color
}

function Import-VBOModule {
    $modulePath = "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"
    if (Get-Module -Name Veeam.Archiver.PowerShell -ErrorAction SilentlyContinue) { return }

    if (Test-Path -LiteralPath $modulePath) {
        Import-Module $modulePath -ErrorAction Stop
    }
    else {
        Import-Module Veeam.Archiver.PowerShell -ErrorAction Stop
    }
}

function Get-AccountFromCsvRow {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [string]$PreferredColumn
    )

    $props = @($Row.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' })
    if ($props.Count -eq 0) { return $null }

    if (-not [string]::IsNullOrWhiteSpace($PreferredColumn)) {
        $selected = $props | Where-Object { $_.Name -eq $PreferredColumn } | Select-Object -First 1
        if ($null -ne $selected) { return ([string]$selected.Value).Trim() }
        throw "La columna '$PreferredColumn' no existe en el CSV. Columnas encontradas: $($props.Name -join ', ')"
    }

    $commonNames = @('UserPrincipalName','UPN','Cuenta','CuentaUsuario','Usuario','Email','Correo','Mail','SamAccountName')
    foreach ($name in $commonNames) {
        $selected = $props | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($null -ne $selected -and -not [string]::IsNullOrWhiteSpace([string]$selected.Value)) {
            return ([string]$selected.Value).Trim()
        }
    }

    return ([string]$props[0].Value).Trim()
}

function Get-EntityIdentityTokens {
    param([Parameter(Mandatory = $true)]$Entity)

    $tokens = New-Object System.Collections.Generic.List[string]
    $candidateProperties = @(
        'UserName','UserPrincipalName','UPN','Email','Mail','PrimarySmtpAddress',
        'DisplayName','Name','Title','Login','Account','ExternalId'
    )

    foreach ($propName in $candidateProperties) {
        $prop = $Entity.PSObject.Properties[$propName]
        if ($null -ne $prop -and $null -ne $prop.Value) {
            $value = ([string]$prop.Value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) { $tokens.Add($value.ToLowerInvariant()) }
        }
    }

    # Algunas versiones muestran toda la identidad al convertir el objeto a string.
    try {
        $asString = ([string]$Entity).Trim()
        if (-not [string]::IsNullOrWhiteSpace($asString)) { $tokens.Add($asString.ToLowerInvariant()) }
    } catch {}

    return @($tokens | Select-Object -Unique)
}

function Find-VBOUserEntityByUpn {
    param(
        [Parameter(Mandatory = $true)][array]$AllUsers,
        [Parameter(Mandatory = $true)][string]$Upn
    )

    $normalizedUpn = $Upn.Trim().ToLowerInvariant()

    # Coincidencia estricta contra propiedades conocidas y representación textual del objeto.
    $matches = @($AllUsers | Where-Object {
        $tokens = Get-EntityIdentityTokens -Entity $_
        $tokens -contains $normalizedUpn
    })

    if ($matches.Count -gt 0) { return $matches }

    # Respaldo: si Veeam expone datos anidados, buscar propiedad donde el valor sea el UPN exacto.
    $matches = @($AllUsers | Where-Object {
        $found = $false
        foreach ($p in $_.PSObject.Properties) {
            if ($null -ne $p.Value -and ([string]$p.Value).Trim().ToLowerInvariant() -eq $normalizedUpn) {
                $found = $true
                break
            }
        }
        $found
    })

    return $matches
}

# =========================
# Inicio
# =========================
Write-Info "Iniciando depuración del repositorio '$RepositoryName'." Cyan
if ($WhatIfMode) { Write-Info "WhatIfMode activo: no se borrarán datos." Yellow }

Import-VBOModule

if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "No se encontró el archivo CSV: $CsvPath"
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $baseDir = Split-Path -Path $CsvPath -Parent
    if ([string]::IsNullOrWhiteSpace($baseDir)) { $baseDir = (Get-Location).Path }
    $LogPath = Join-Path $baseDir ("Depurar_Repo_Cache_Old_resultado_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

$repo = Get-VBORepository -Name $RepositoryName
if ($null -eq $repo) {
    throw "No se encontró el repositorio '$RepositoryName'."
}

Write-Info "Cargando entidades tipo User desde el repositorio. Esto puede tardar según el tamaño del repositorio..." Cyan
$allRepoUsers = @(Get-VBOEntityData -Type User -Repository $repo)
Write-Info "Usuarios encontrados en repositorio: $($allRepoUsers.Count)." Gray

$rows = @(Import-Csv -LiteralPath $CsvPath)
if ($rows.Count -eq 0) {
    throw "El CSV no contiene registros: $CsvPath"
}

$validDomainPattern = '^[A-Za-z0-9._%+\-]+@uaesp\.gov\.co$'
$successCount = 0
$notFoundCount = 0
$invalidCount = 0
$errorCount = 0
$processedCount = 0
$results = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $processedCount++
    $account = Get-AccountFromCsvRow -Row $row -PreferredColumn $CsvColumn

    if ([string]::IsNullOrWhiteSpace($account) -or ($account -notmatch $validDomainPattern)) {
        $invalidCount++
        $message = "Formato inválido. Se esperaba nombre_usuario@uaesp.gov.co"
        Write-Info "$account - $message" Yellow
        $results.Add([pscustomobject]@{
            Cuenta = $account
            Estado = "Formato inválido"
            Mensaje = $message
            Fecha = Get-Date
        })
        continue
    }

    $account = $account.ToLowerInvariant()
    Write-Info "Procesando cuenta: $account" Cyan

    $matches = @(Find-VBOUserEntityByUpn -AllUsers $allRepoUsers -Upn $account)

    if ($matches.Count -eq 0) {
        $notFoundCount++
        Write-Info "$account - Cuenta no encontrada" Yellow
        $results.Add([pscustomobject]@{
            Cuenta = $account
            Estado = "Cuenta no encontrada"
            Mensaje = "Cuenta no encontrada"
            Fecha = Get-Date
        })
        continue
    }

    if ($matches.Count -gt 1) {
        Write-Info "$account - Se encontraron $($matches.Count) coincidencias exactas; se intentará borrar cada entidad coincidente." Yellow
    }

    $removedForAccount = 0
    foreach ($userEntity in $matches) {
        try {
            if ($WhatIfMode) {
                Remove-VBOEntityData -Repository $repo -User $userEntity -Mailbox -ArchiveMailbox -OneDrive -Sites -Confirm:$false -WhatIf
            }
            else {
                Remove-VBOEntityData -Repository $repo -User $userEntity -Mailbox -ArchiveMailbox -OneDrive -Sites -Confirm:$false
            }

            $removedForAccount++
            $successCount++
            Write-Info "$account - Borrado exitoso: Exchange, Archive Mailbox, OneDrive y Sites." Green
            $results.Add([pscustomobject]@{
                Cuenta = $account
                Estado = $(if ($WhatIfMode) { "Simulado" } else { "Borrado exitoso" })
                Mensaje = "Proceso finalizado exitosamente para la entidad encontrada."
                Fecha = Get-Date
            })
        }
        catch {
            $errorCount++
            $err = $_.Exception.Message
            Write-Info "$account - Error durante el borrado: $err" Red
            $results.Add([pscustomobject]@{
                Cuenta = $account
                Estado = "Error"
                Mensaje = $err
                Fecha = Get-Date
            })
        }
    }
}

$results | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "================ RESUMEN FINAL ================" -ForegroundColor Cyan
Write-Host "Repositorio: $RepositoryName"
Write-Host "CSV procesado: $CsvPath"
Write-Host "Registros recorridos: $processedCount"
Write-Host "Cuentas/copias borradas exitosamente: $successCount" -ForegroundColor Green
Write-Host "Cuentas no encontradas: $notFoundCount" -ForegroundColor Yellow
Write-Host "Registros con formato inválido: $invalidCount" -ForegroundColor Yellow
Write-Host "Errores: $errorCount" -ForegroundColor Red
Write-Host "Log: $LogPath"
Write-Host "===============================================" -ForegroundColor Cyan
