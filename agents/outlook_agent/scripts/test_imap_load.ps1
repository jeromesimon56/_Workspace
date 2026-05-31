$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
function Resolve-AssemblyPath {
    param([string]$AssemblyName)
    $path = Join-Path $scriptDir "lib\$AssemblyName"
    if (Test-Path $path) { return $path }
    return $null
}
$required = @('System.Threading.Tasks.Extensions.dll','MimeKit.dll','MailKit.dll')
foreach ($dll in $required) {
    $path = Resolve-AssemblyPath -AssemblyName $dll
    if ($path) {
        [Reflection.Assembly]::LoadFrom($path) | Out-Null
        Write-Host "Loaded $dll"
    } else { Write-Host "Missing $dll" }
}
try {
    $client = [MailKit.Net.Imap.ImapClient]::new()
    Write-Host "Created ImapClient"
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
