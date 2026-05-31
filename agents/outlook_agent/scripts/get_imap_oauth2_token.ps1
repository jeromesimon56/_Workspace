param(
    [string]$ClientId = '',
    [string]$TenantId = 'common',
    [string]$Scope = 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access openid profile',
    [switch]$UseDeviceCode
)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

function Ensure-ModuleInstalled {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing PowerShell module: $Name"
        Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module $Name -Force
}

if (-not $ClientId) {
    if ($env:AZURE_CLIENT_ID) {
        $ClientId = $env:AZURE_CLIENT_ID
    } else {
        Write-Error "Un ClientId est requis pour récupérer un jeton OAuth2. Définissez -ClientId ou AZURE_CLIENT_ID."
        exit 1
    }
}

try {
    Ensure-ModuleInstalled -Name 'MSAL.PS'
} catch {
    Write-Error "Impossible d'installer ou de charger MSAL.PS : $($_.Exception.Message)"
    exit 1
}

Write-Host "Demande de jeton OAuth2 pour IMAP avec ClientId=$ClientId TenantId=$TenantId"

$tokenParams = @{
    ClientId = $ClientId
    TenantId = $TenantId
    Scopes = $Scope
    ErrorAction = 'Stop'
}

if ($UseDeviceCode) {
    $tokenParams['UseDeviceCode'] = $true
}

try {
    $result = Get-MsalToken @tokenParams
} catch {
    Write-Error "Impossible d'obtenir le jeton OAuth2 : $($_.Exception.Message)"
    exit 1
}

if (-not $result.AccessToken) {
    Write-Error "Aucun jeton d'accès reçu."
    exit 1
}

Write-Host "Token acquired. Expires: $($result.ExpiresOn)"
Write-Host "Use this token with -AuthMethod oauth2 -OAuth2Token '<token>'"
Write-Output $result.AccessToken
