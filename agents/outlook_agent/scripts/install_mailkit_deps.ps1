#Requires -Version 5.0
<#
.SYNOPSIS
Telecharge et installe les dependances MailKit.
#>

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$libDir = Join-Path $scriptDir 'lib'
$tempDir = Join-Path $env:TEMP "mailkit_$([guid]::NewGuid().ToString().Substring(0,8))"

$packages = @{
    'MailKit' = '4.3.0'
    'MimeKit' = '4.3.0'
    'System.Runtime' = '4.3.0'
}

$requiredDlls = @('System.Runtime.dll','MimeKit.dll','MailKit.dll')

function Ensure-Dir {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Download-NuGet {
    param([string]$Name, [string]$Ver, [string]$Out)
    $url = "https://www.nuget.org/api/v2/package/$Name/$Ver"
    $zip = Join-Path $Out "$Name.$Ver.nupkg"
    Write-Host "Telechargement $Name/$Ver..."
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $zip -ErrorAction Stop
        Write-Host "  OK Telecharge"
        return $zip
    }
    catch {
        Write-Error "Erreur: $_"
        return $null
    }
}

function Extract-Dll {
    param([string]$Zip, [string]$Out, [string[]]$Dlls)
    $ext = Join-Path $env:TEMP "nupkg_$([guid]::NewGuid().ToString().Substring(0,8))"
    Ensure-Dir -Path $ext
    try {
        $zipFile = $Zip
        if ($Zip -like "*.nupkg") {
            $zipFile = $Zip -replace '\.nupkg$', '.zip'
            Copy-Item -Path $Zip -Destination $zipFile -Force
        }
        Expand-Archive -Path $zipFile -DestinationPath $ext -Force
        $paths = @('lib/net6.0','lib/net8.0','lib/net462','lib/netstandard2.1','lib/net5.0','lib')
        foreach ($dll in $Dlls) {
            foreach ($p in $paths) {
                $src = Join-Path (Join-Path $ext $p) $dll
                if (Test-Path $src) {
                    Copy-Item -Path $src -Destination (Join-Path $Out $dll) -Force
                    Write-Host "  OK Extrait : $dll"
                    break
                }
            }
        }
    }
    finally {
        Remove-Item -Path $ext -Recurse -Force -ErrorAction SilentlyContinue
        if ($zipFile -and $zipFile -like "*.zip" -and (Test-Path $zipFile)) {
            Remove-Item -Path $zipFile -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Host ""
Write-Host "Installation MailKit" -ForegroundColor Cyan
Write-Host ""

Ensure-Dir -Path $libDir
Ensure-Dir -Path $tempDir

$files = @{}
foreach ($pkg in $packages.GetEnumerator()) {
    $f = Download-NuGet -Name $pkg.Key -Ver $pkg.Value -Out $tempDir
    if ($f) { $files[$pkg.Key] = $f }
}

Write-Host ""

$map = @{
    'MailKit' = @('MailKit.dll')
    'MimeKit' = @('MimeKit.dll')
    'System.Runtime' = @('System.Runtime.dll')
}

foreach ($p in $map.GetEnumerator()) {
    if ($files.ContainsKey($p.Key)) {
        Write-Host "Extraction $($p.Key)..."
        Extract-Dll -Zip $files[$p.Key] -Out $libDir -Dlls $p.Value
    }
}

Write-Host ""
Write-Host "Verification :" -ForegroundColor Cyan

$ok = $true
foreach ($dll in $requiredDlls) {
    $path = Join-Path $libDir $dll
    if (Test-Path $path) {
        $kb = [math]::Round((Get-Item $path).Length / 1024, 1)
        Write-Host "  OK $dll ($kb KB)"
    }
    else {
        Write-Host "  XX $dll"
        $ok = $false
    }
}

Write-Host ""
if ($ok) {
    Write-Host "OK Succes !" -ForegroundColor Green
} else {
    Write-Host "ERR Manquant(s)" -ForegroundColor Red
}

Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
