$scriptDir = 'c:\Users\jerom\Documents\_Workspace\agents\outlook_agent\scripts'

function Resolve-AssemblyPath {
    param([string]$AssemblyName)
    $path = Join-Path $scriptDir "lib\$AssemblyName"
    if (Test-Path $path) { return $path }
    return $null
}

function Load-RequiredAssemblies {
    $required = @(
        'System.Runtime.dll',
        'MimeKit.dll',
        'MailKit.dll'
    )

    $missing = @()
    foreach ($dll in $required) {
        $path = Resolve-AssemblyPath -AssemblyName $dll
        if (-not $path) {
            $missing += $dll
            continue
        }

        try {
            [Reflection.Assembly]::LoadFrom($path) | Out-Null
            Write-Host "OK Charge : $dll"
        } catch {
            Write-Warning "Erreur : $dll : $($_.Exception.Message)"
            $missing += $dll
        }
    }

    if ($missing.Count -gt 0) {
        Write-Host "ERREUR: Manquants: $($missing -join ', ')"
        return $false
    }
    return $true
}

Load-RequiredAssemblies
