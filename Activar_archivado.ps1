<#
.SYNOPSIS
    Aplica la política de retención de Exchange Online y fuerza el movimiento
    de correos al Archive del buzón especificado.

.DESCRIPTION
    El script realiza las siguientes acciones:

    1. Conecta a Exchange Online.
    2. Valida que el buzón exista.
    3. Consulta el estado actual del buzón.
    4. Consulta el estado del Archive.
    5. Aplica la política de retención.
    6. Ejecuta Start-ManagedFolderAssistant.
    7. Verifica si el Archive empezó a recibir elementos.
    8. Muestra resultados en pantalla.

.PARAMETER Mailbox
    Cuenta de correo a procesar.

.EXAMPLE
    .\Invoke-ArchiveProcessing.ps1 -Mailbox sandra.ruiz@uaesp.gov.co

.NOTES
    Requisitos:
     - Exchange Online Management Module
     - Permisos de Exchange Administrator
     - Archive habilitado en el buzón

    Instalación del módulo:
        Install-Module ExchangeOnlineManagement

#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Mailbox
)

# Política de retención
$RetentionPolicy = "Archivo 6 meses"

# Tiempo de espera para verificar cambios
$WaitSeconds = 120

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " INICIO DEL PROCESO DE ARCHIVADO" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

try {

    #---------------------------------------------------------
    # Conexión Exchange Online
    #---------------------------------------------------------
    Write-Host "`n[1/8] Conectando a Exchange Online..." -ForegroundColor Yellow

    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

    #---------------------------------------------------------
    # Validar existencia del buzón
    #---------------------------------------------------------
    Write-Host "[2/8] Validando buzón..." -ForegroundColor Yellow

    $MailboxObject = Get-Mailbox -Identity $Mailbox -ErrorAction Stop

    Write-Host "Buzón encontrado: $($MailboxObject.DisplayName)" -ForegroundColor Green

    #---------------------------------------------------------
    # Estado actual del Archive
    #---------------------------------------------------------
    Write-Host "[3/8] Consultando Archive..." -ForegroundColor Yellow

    $ArchiveInfo = Get-Mailbox $Mailbox |
        Select-Object DisplayName, ArchiveStatus, ArchiveName

    $ArchiveInfo | Format-Table -AutoSize

    if ($ArchiveInfo.ArchiveStatus -ne "Active") {

        throw "El Archive no se encuentra habilitado para este buzón."
    }

    #---------------------------------------------------------
    # Estadísticas iniciales del buzón
    #---------------------------------------------------------
    Write-Host "[4/8] Consultando estadísticas del buzón..." -ForegroundColor Yellow

    $MailboxStatsBefore = Get-MailboxStatistics $Mailbox -ErrorAction Stop

    Write-Host ""
    Write-Host "Estado inicial:" -ForegroundColor Cyan

    $MailboxStatsBefore |
        Select-Object DisplayName,
                      TotalItemSize,
                      ItemCount,
                      LastLogonTime |
        Format-Table -AutoSize

    #---------------------------------------------------------
    # Estadísticas iniciales del Archive
    #---------------------------------------------------------
    try {

        $ArchiveStatsBefore = Get-MailboxStatistics $Mailbox -Archive -ErrorAction Stop

        Write-Host ""
        Write-Host "Archive inicial:" -ForegroundColor Cyan

        $ArchiveStatsBefore |
            Select-Object TotalItemSize, ItemCount |
            Format-Table -AutoSize
    }
    catch {

        throw "No fue posible consultar las estadísticas del Archive."
    }

    #---------------------------------------------------------
    # Aplicar política de retención
    #---------------------------------------------------------
    Write-Host "`n[5/8] Aplicando política '$RetentionPolicy' ..." -ForegroundColor Yellow

    Set-Mailbox `
        -Identity $Mailbox `
        -RetentionPolicy $RetentionPolicy `
        -ErrorAction Stop

    Start-Sleep -Seconds 5

    $RetentionCheck = Get-Mailbox -Identity $Mailbox |
        Select-Object Name, RetentionPolicy

    Write-Host "Política aplicada correctamente:" -ForegroundColor Green

    $RetentionCheck | Format-Table -AutoSize

    #---------------------------------------------------------
    # Forzar Managed Folder Assistant
    #---------------------------------------------------------
    Write-Host "`n[6/8] Ejecutando Managed Folder Assistant..." -ForegroundColor Yellow

    Start-ManagedFolderAssistant `
        -Identity $Mailbox `
        -ErrorAction Stop

    Write-Host "Procesamiento iniciado correctamente." -ForegroundColor Green

    #---------------------------------------------------------
    # Esperar procesamiento
    #---------------------------------------------------------
    Write-Host "`nEsperando $WaitSeconds segundos para verificar resultados..." -ForegroundColor Yellow

    Start-Sleep -Seconds $WaitSeconds

    #---------------------------------------------------------
    # Verificación de movimiento al Archive
    #---------------------------------------------------------
    Write-Host "`n[7/8] Verificando crecimiento del Archive..." -ForegroundColor Yellow

    $ArchiveStatsAfter = Get-MailboxStatistics $Mailbox -Archive -ErrorAction Stop

    Write-Host ""
    Write-Host "Estado Archive después del procesamiento:" -ForegroundColor Cyan

    $ArchiveStatsAfter |
        Select-Object TotalItemSize, ItemCount |
        Format-Table -AutoSize

    $InitialItems = [int]$ArchiveStatsBefore.ItemCount
    $FinalItems   = [int]$ArchiveStatsAfter.ItemCount

    Write-Host ""

    if ($FinalItems -gt $InitialItems)
    {
        Write-Host "✓ Se detectó movimiento de elementos al Archive." -ForegroundColor Green
        Write-Host "Items iniciales : $InitialItems"
        Write-Host "Items finales   : $FinalItems"
        Write-Host "Diferencia      : $($FinalItems - $InitialItems)"
    }
    else
    {
        Write-Warning "No se detectó incremento inmediato en el Archive."

        Write-Host ""
        Write-Host "NOTA:" -ForegroundColor Yellow
        Write-Host "Exchange Online puede tardar varios minutos u horas" -ForegroundColor Yellow
        Write-Host "en completar el movimiento de elementos según la" -ForegroundColor Yellow
        Write-Host "carga del servicio y volumen del buzón." -ForegroundColor Yellow
    }

    Write-Host "`n[8/8] Proceso completado." -ForegroundColor Green

}
catch {

    Write-Error ""
    Write-Error "ERROR EN EL PROCESO"
    Write-Error $_.Exception.Message
}
finally {

    Write-Host "`nDesconectando sesión..." -ForegroundColor Yellow

    Disconnect-ExchangeOnline `
        -Confirm:$false `
        -ErrorAction SilentlyContinue

    Write-Host "Sesión finalizada." -ForegroundColor Cyan
}