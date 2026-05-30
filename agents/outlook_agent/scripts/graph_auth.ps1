param()

Write-Host "Authenticating to Microsoft Graph..."

if ($env:AZURE_CLIENT_ID -and $env:AZURE_TENANT_ID -and $env:AZURE_CLIENT_SECRET) {
    Write-Host "Using app-only authentication via AZURE_CLIENT_ID/AZURE_TENANT_ID/AZURE_CLIENT_SECRET."
    Connect-MgGraph -ClientId $env:AZURE_CLIENT_ID -TenantId $env:AZURE_TENANT_ID -ClientSecret $env:AZURE_CLIENT_SECRET -Scopes "https://graph.microsoft.com/.default" -NoWelcome
} else {
    $deviceAuthSupported = (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue).Parameters.Keys -contains 'UseDeviceAuthentication'
    if ($deviceAuthSupported) {
        Write-Host "Using device code authentication for Microsoft Graph."
        Connect-MgGraph -Scopes "Mail.Read","Mail.ReadWrite","User.Read" -UseDeviceAuthentication -NoWelcome
    } else {
        Write-Host "Using interactive authentication for Microsoft Graph."
        Connect-MgGraph -Scopes "Mail.Read","Mail.ReadWrite","User.Read" -NoWelcome
    }
}

if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) {
    Select-MgProfile -Name "beta"
} else {
    Write-Host "Select-MgProfile unavailable; using the default Microsoft Graph profile."
}

Write-Host "Microsoft Graph authentication complete."
