function Update-VBODesvinculacionJob {

<#
===============================================================================
UPDATE-VBODESVINCULACIONJOB
===============================================================================

OBJETIVO
--------
Actualizar el Job "Desvinculacion" en Veeam Backup for Microsoft 365
a partir de una lista de usuarios contenida en un archivo CSV.

PRELIMINARES
------------
Se debe crear el archivo CSV con el formato que se indica a continuacion.
El archivo contiene UserPrincipalName de los usuarios que se encuentran 
bloqueados en M365 y que tienen licencia asignada. El archivo CSV se debe
cargar en la carpeta C:\Scripts\ del servidor de VBO365.

FORMATO DEL CSV
---------------

UserPrincipalName
usuario1@uaesp.gov.co
usuario2@uaesp.gov.co
usuario3@uaesp.gov.co

CARGAR LA FUNCION EN MEMORIA
----------------------------

Si modifica este archivo .ps1, recargue la función:

Remove-Item Function:\Update-VBODesvinculacionJob `
    -ErrorAction SilentlyContinue

. C:\Scripts\Update-VBODesvinculacionJob.ps1

Verificar carga:

Get-Command Update-VBODesvinculacionJob

MODO SIMULACION
---------------

Update-VBODesvinculacionJob `
    -CsvPath "C:\Scripts\UsuariosDesvinculacion.csv" `
    -WhatIf

MODO PRODUCTIVO
---------------

Update-VBODesvinculacionJob `
    -CsvPath "C:\Scripts\UsuariosDesvinculacion.csv" `
    -Confirm:$false

NOTA IMPORTANTE
---------------

Veeam no permite que un Job quede sin objetos protegidos.

Por esta razón el script:

1. Agrega usuarios nuevos.
2. Verifica que fueron agregados.
3. Elimina usuarios antiguos.

===============================================================================
#>

    [CmdletBinding(
        SupportsShouldProcess = $true,
        ConfirmImpact = 'High'
    )]

    param(

        [Parameter(Mandatory = $true)]
        [ValidateScript({
            Test-Path $_
        })]
        [string]$CsvPath,
        [string]$OrganizationName = "uaespdc.onmicrosoft.com",
        [string]$JobName = "Desvinculacion",
        [string]$RepositoryName = "REPO_CACHE_AZURE"
    )

    #==========================================================================
    # FUNCION DE LOG
    #==========================================================================

    function Write-Log {

        param(
            [string]$Message,
            [string]$Level = "INFO"
        )

        $Color = switch ($Level) {

            "INFO"    { "White"  }
            "OK"      { "Green"  }
            "WARNING" { "Yellow" }
            "ERROR"   { "Red"    }

            default   { "White"  }
        }

        Write-Host "[$Level] $Message" -ForegroundColor $Color
    }

    #==========================================================================
    # INICIO TRANSCRIPT
    #==========================================================================

    $LogFolder = "C:\Logs"

    if (-not (Test-Path $LogFolder)) {

        New-Item `
            -Path $LogFolder `
            -ItemType Directory `
            -Force | Out-Null
    }

    $TranscriptFile = Join-Path `
        $LogFolder `
        ("Desvinculacion_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    Start-Transcript -Path $TranscriptFile

    try {

        Write-Log "==========================================="
        Write-Log "INICIO DEL PROCESO"
        Write-Log "==========================================="

        if ($WhatIfPreference) {

            Write-Log "MODO WHATIF ACTIVADO" "WARNING"
        }
        else {

            Write-Log "MODO PRODUCTIVO ACTIVADO" "OK"
        }

        #======================================================================
        # CARGAR CSV
        #======================================================================

        Write-Log "Cargando CSV..."

        $CsvUsers = Import-Csv -Path $CsvPath

        if (-not $CsvUsers) {
            throw "El archivo CSV no contiene registros."
        }

        if (-not ($CsvUsers[0].PSObject.Properties.Name -contains "UserPrincipalName")) {
            throw "No existe la columna UserPrincipalName."
        }

        Write-Log "Usuarios encontrados: $($CsvUsers.Count)" "OK"

        #======================================================================
        # CONEXION VBO365
        #======================================================================

        Write-Log "Conectando a VBO365..."
        Connect-VBOServer -Server localhost
        Write-Log "Conexion exitosa." "OK"

        #======================================================================
        # VALIDAR VERSION
        #======================================================================

        Write-Log "Validando version..."
        $DllPath = "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell.dll"
        if (-not (Test-Path $DllPath)) {
            throw "No se encontró: $DllPath"
        }

        $InstalledVersion = (
            Get-Item $DllPath
        ).VersionInfo.FileVersion

        Write-Log "Version detectada: $InstalledVersion"

        if ($InstalledVersion -notlike "13.4.*") {
            throw "Version incompatible: $InstalledVersion"
        }

        Write-Log "Version validada correctamente." "OK"

        #======================================================================
        # ORGANIZACION
        #======================================================================

        Write-Log "Buscando organizacion..."

        $Organization = Get-VBOOrganization |
            Where-Object {
                $_.Name -eq $OrganizationName
            }

        if (-not $Organization) {
            throw "No se encontró la organización [$OrganizationName]"
        }

        Write-Log "Organizacion encontrada." "OK"

        #======================================================================
        # JOB
        #======================================================================

        Write-Log "Buscando Job..."

        $Job = Get-VBOJob |
            Where-Object {
                $_.Name -eq $JobName
            }

        if (-not $Job) {
            throw "No se encontró el Job [$JobName]"
        }

        Write-Log "Job encontrado." "OK"

        #======================================================================
        # REPOSITORIO
        #======================================================================

        Write-Log "Validando repositorio..."

        $Repository = Get-VBORepository |
            Where-Object {
                $_.Name -eq $RepositoryName
            }

        if (-not $Repository) {
            throw "Repositorio no encontrado [$RepositoryName]"
        }

        Write-Log "Repositorio validado." "OK"

        #======================================================================
        # OBJETOS ACTUALES
        #======================================================================

        $OriginalObjects = Get-VBOBackupItem -Job $Job

        Write-Log "Objetos actuales encontrados: $($OriginalObjects.Count)"

        #======================================================================
        # CARGAR USUARIOS ORGANIZACION
        #======================================================================

        Write-Log "Cargando usuarios de la organización..."

        $OrganizationUsers = Get-VBOOrganizationUser `
            -Organization $Organization

        Write-Log "Usuarios cargados: $($OrganizationUsers.Count)"

        #======================================================================
        # AGREGAR USUARIOS CSV
        #======================================================================

        Write-Log "Agregando usuarios del CSV..."

        $UsersAdded = 0

        foreach ($User in $CsvUsers) {

            try {
                $UPN = $User.UserPrincipalName.Trim()
                Write-Log "Procesando $UPN"
                $OrgUser = $OrganizationUsers |
                    Where-Object {
                        $_.UserName -ieq $UPN
                    } |
                    Select-Object -First 1

                if (-not $OrgUser) {
                    Write-Log "Usuario no encontrado: $UPN" "WARNING"
                    continue
                }

                $AlreadyExists = Get-VBOBackupItem -Job $Job |
                    Where-Object {
                        $_.User -ieq $UPN
                    }

                if ($AlreadyExists) {
                    Write-Log "Usuario ya existe en el Job: $UPN" "WARNING"
                    continue
                }

                if ($PSCmdlet.ShouldProcess(
                    $UPN,
                    "Agregar usuario al Job"
                ))
                {
                    # Crear objeto BackupItem compatible con VBO365

                    $BackupItem = New-VBOBackupItem `
                        -User $OrgUser `
                        -Mailbox `
                        -ArchiveMailbox `
                        -OneDrive `
                        -Sites

                    # Agregar al Job

                    Add-VBOBackupItem `
                        -Job $Job `
                        -BackupItem $BackupItem

                    $UsersAdded++

                    Write-Log "Usuario agregado: $UPN" "OK"
                }
            }
            catch {

                Write-Log "Error agregando usuario: $UPN" "ERROR"
                Write-Log $_.Exception.Message "ERROR"
            }
        }

        #======================================================================
        # VALIDACION DE SEGURIDAD
        #======================================================================

        $CurrentCount = (
            Get-VBOBackupItem -Job $Job
        ).Count

        Write-Log "Objetos despues del agregado: $CurrentCount"

        if ($UsersAdded -eq 0) {
            throw @"

No se agregó ningún usuario al Job.
Para proteger la configuración existente
NO se eliminarán los objetos actuales.
"@
        }

        #======================================================================
        # ELIMINAR OBJETOS ORIGINALES
        #======================================================================

        Write-Log "Eliminando objetos antiguos..."

        foreach ($CurrentObject in $OriginalObjects) {

            try {

                if ($PSCmdlet.ShouldProcess(
                    $CurrentObject.User,
                    "Eliminar objeto del Job"
                )) {

                    Remove-VBOBackupItem `
                        -Job $Job `
                        -BackupItem $CurrentObject

                    Write-Log "Eliminado: $($CurrentObject.User)" "OK"
                }
            }
            catch {

                Write-Log "Error eliminando: $($CurrentObject.User)" "ERROR"
                Write-Log $_.Exception.Message "ERROR"
            }
        }

        #======================================================================
        # LIMPIAR EXCLUSIONES
        #======================================================================

        Write-Log "Limpiando exclusiones..."

        $ExcludedItems = Get-VBOExcludedBackupItem `
            -Job $Job `
            -ErrorAction SilentlyContinue

        foreach ($ExcludedItem in $ExcludedItems) {

            try {

                if ($PSCmdlet.ShouldProcess(
                    "Exclusion",
                    "Eliminar exclusion"
                )) {

                    Remove-VBOExcludedBackupItem `
                        -Job $Job `
                        -ExcludedItem $ExcludedItem

                    Write-Log "Exclusion eliminada." "OK"
                }
            }
            catch {

                Write-Log $_.Exception.Message "ERROR"
            }
        }

        #======================================================================
        # RESUMEN FINAL
        #======================================================================

        $FinalCount = (
            Get-VBOBackupItem -Job $Job
        ).Count

        Write-Log "===========================================" "OK"
        Write-Log "RESUMEN FINAL" "OK"
        Write-Log "===========================================" "OK"

        Write-Log "Usuarios CSV      : $($CsvUsers.Count)" "OK"
        Write-Log "Usuarios agregados: $UsersAdded" "OK"
        Write-Log "Objetos finales   : $FinalCount" "OK"

        Write-Log "Proceso finalizado correctamente." "OK"
    }
    catch {

        Write-Log "ERROR GENERAL DEL PROCESO" "ERROR"
        Write-Log $_.Exception.Message "ERROR"

        throw
    }
    finally {

        Disconnect-VBOServer -ErrorAction SilentlyContinue
        Write-Log "Sesion desconectada."
        Stop-Transcript
    }
}

        