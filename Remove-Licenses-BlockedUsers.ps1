<#
.SYNOPSIS
Remueve todas las licencias de usuarios bloqueados en Microsoft 365.

.DESCRIPTION
Lee un archivo CSV con una lista de usuarios y realiza las siguientes acciones:

1. Busca el usuario en Microsoft 365.
2. Consulta el estado de inicio de sesión (AccountEnabled).
3. Si está activo (Allowed), no realiza ninguna acción.
4. Si está bloqueado (Blocked), consulta las licencias asignadas.
5. Si tiene licencias, las elimina todas.
6. Muestra información detallada de cada paso.
7. Soporta modo simulación mediante -WhatIf.

.REQUIREMENTS
- Microsoft Graph PowerShell SDK
- Permisos:
    User.Read.All
    Directory.Read.All
    User.ReadWrite.All

.EXAMPLE
Simulación:

.\Remove-Licenses-BlockedUsers.ps1 `
    -CsvPath "C:\Temp\Usuarios.csv" `
    -WhatIf

.EXAMPLE
Ejecución real:

.\Remove-Licenses-BlockedUsers.ps1 `
    -CsvPath "C:\Temp\Usuarios.csv"

.NOTES
Autor: Microsoft 365 / PowerShell Automation
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath
)

#----------------------------------------------------------------------------------
# INICIO
#----------------------------------------------------------------------------------

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " PROCESO DE VALIDACIÓN Y REMOCIÓN DE LICENCIAS M365 " -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

if ($WhatIfPreference)
{
    Write-Host ""
    Write-Host "MODO SIMULACIÓN ACTIVADO (-WhatIf)" -ForegroundColor Yellow
    Write-Host "NO se realizarán cambios reales en Microsoft 365." -ForegroundColor Yellow
    Write-Host ""
}
else
{
    Write-Host ""
    Write-Host "MODO EJECUCIÓN REAL" -ForegroundColor Green
    Write-Host ""
}

#----------------------------------------------------------------------------------
# VALIDAR CSV
#----------------------------------------------------------------------------------

try
{
    if (-not (Test-Path $CsvPath))
    {
        throw "El archivo CSV no existe: $CsvPath"
    }

    $Users = Import-Csv $CsvPath

    if ($Users.Count -eq 0)
    {
        throw "El archivo CSV no contiene registros."
    }

    Write-Host "Archivo CSV cargado correctamente." -ForegroundColor Green
    Write-Host "Cantidad de usuarios encontrados: $($Users.Count)" -ForegroundColor Green
}
catch
{
    Write-Host ""
    Write-Host "ERROR al cargar el archivo CSV." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit
}

#----------------------------------------------------------------------------------
# CONEXIÓN A MICROSOFT GRAPH
#----------------------------------------------------------------------------------

try
{
    Write-Host ""
    Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan

    Connect-MgGraph `
        -Scopes "User.Read.All","Directory.Read.All","User.ReadWrite.All" `
        -NoWelcome

    Write-Host "Conexión exitosa." -ForegroundColor Green
}
catch
{
    Write-Host ""
    Write-Host "ERROR al conectar a Microsoft Graph." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit
}

#----------------------------------------------------------------------------------
# VARIABLES DE RESUMEN
#----------------------------------------------------------------------------------

$TotalProcesados = 0
$TotalActivos = 0
$TotalBloqueados = 0
$TotalLicenciasRemovidas = 0
$TotalErrores = 0

#----------------------------------------------------------------------------------
# PROCESAMIENTO
#----------------------------------------------------------------------------------

foreach ($Item in $Users)
{
    $TotalProcesados++

    try
    {
        #----------------------------------------------------------------------
        # AJUSTAR EL NOMBRE DE LA COLUMNA SEGÚN EL CSV
        #----------------------------------------------------------------------

        $UPN = $Item.UserPrincipalName

        Write-Host ""
        Write-Host "=====================================================" -ForegroundColor DarkGray
        Write-Host "Procesando usuario: $UPN" -ForegroundColor Yellow
        Write-Host "=====================================================" -ForegroundColor DarkGray

        #----------------------------------------------------------------------
        # CONSULTAR USUARIO
        #----------------------------------------------------------------------

        $User = Get-MgUser `
            -UserId $UPN `
            -Property Id,DisplayName,UserPrincipalName,AccountEnabled

        if (-not $User)
        {
            Write-Host "Usuario no encontrado." -ForegroundColor Red
            continue
        }

        Write-Host "Nombre: $($User.DisplayName)"
        Write-Host "UPN: $($User.UserPrincipalName)"

        #----------------------------------------------------------------------
        # ESTADO DE INICIO DE SESIÓN
        #----------------------------------------------------------------------

        if ($User.AccountEnabled -eq $true)
        {
            $TotalActivos++

            Write-Host "Estado de inicio de sesión: ACTIVO / ALLOWED" -ForegroundColor Green
            Write-Host "Resultado: No se requiere acción." -ForegroundColor Green

            continue
        }

        $TotalBloqueados++

        Write-Host "Estado de inicio de sesión: BLOQUEADO / BLOCKED" -ForegroundColor Red

        #----------------------------------------------------------------------
        # CONSULTAR LICENCIAS
        #----------------------------------------------------------------------

        $Licenses = Get-MgUserLicenseDetail -UserId $User.Id

        $LicenseCount = $Licenses.Count

        Write-Host "Cantidad de licencias asignadas: $LicenseCount"

        if ($LicenseCount -eq 0)
        {
            Write-Host "No tiene licencias asignadas." -ForegroundColor Yellow
            continue
        }

        Write-Host ""
        Write-Host "Licencias detectadas:" -ForegroundColor Cyan

        $SkuIds = @()

        foreach ($License in $Licenses)
        {
            Write-Host "  - $($License.SkuPartNumber)"

            $SkuIds += $License.SkuId
        }

        Write-Host ""

        #----------------------------------------------------------------------
        # REMOVER LICENCIAS (CON SOPORTE WHATIF)
        #----------------------------------------------------------------------

        if ($PSCmdlet.ShouldProcess(
                $User.UserPrincipalName,
                "Remover $($SkuIds.Count) licencias"))
        {
            Set-MgUserLicense `
                -UserId $User.Id `
                -AddLicenses @() `
                -RemoveLicenses $SkuIds

            $TotalLicenciasRemovidas += $SkuIds.Count

            Write-Host "Licencias removidas correctamente." -ForegroundColor Green
        }
    }
    catch
    {
        $TotalErrores++

        Write-Host ""
        Write-Host "ERROR procesando usuario: $UPN" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}

#----------------------------------------------------------------------------------
# RESUMEN FINAL
#----------------------------------------------------------------------------------

Write-Host ""
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "RESUMEN DEL PROCESO" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

Write-Host "Usuarios procesados      : $TotalProcesados"
Write-Host "Usuarios activos         : $TotalActivos"
Write-Host "Usuarios bloqueados      : $TotalBloqueados"
Write-Host "Licencias removidas      : $TotalLicenciasRemovidas"
Write-Host "Errores                  : $TotalErrores"

#----------------------------------------------------------------------------------
# DESCONECTAR
#----------------------------------------------------------------------------------

try
{
    Disconnect-MgGraph | Out-Null
}
catch
{
}

Write-Host ""
Write-Host "Proceso finalizado." -ForegroundColor Green
Write-Host ""