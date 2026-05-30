try {
    Write-Host "Checking for NuGet provider..."
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    }

    Write-Host "Installing Microsoft.Graph module (current user)..."
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    } else {
        Write-Host "Microsoft.Graph already installed."
    }

    Write-Host "Done. You can now run Connect-MgGraph to authenticate."
} catch {
    Write-Error "Installation failed: $_"
    exit 1
}
