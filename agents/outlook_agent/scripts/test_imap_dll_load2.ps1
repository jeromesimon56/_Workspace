$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$paths = @(
    'System.Threading.Tasks.Extensions.dll',
    'MimeKit.dll',
    'MailKit.dll'
)
foreach ($dll in $paths) {
    $full = Join-Path $scriptDir "lib\$dll"
    Write-Host "Path: $full"
    if (Test-Path $full) {
        [Reflection.Assembly]::LoadFrom($full) | Out-Null
        Write-Host "Loaded: $dll"
    } else {
        Write-Host "Missing file: $dll"
    }
}
Write-Host "Assemblies loaded:"
[AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -match 'MailKit|MimeKit|System.Threading.Tasks.Extensions' } | ForEach-Object { Write-Host $_.FullName }

Write-Host "Types containing Imap:"
$asm = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'MailKit' }
if ($asm) {
    $asm = $asm[0]
    try {
        $asm.GetTypes() | Where-Object { $_.FullName -match 'Imap' } | Sort-Object FullName | ForEach-Object { Write-Host $_.FullName }
    } catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        Write-Host "GetTypes failed: $($ex.Message)"
        if ($ex -is [Reflection.ReflectionTypeLoadException]) {
            Write-Host "LoaderExceptions:"
            foreach ($loader in $ex.LoaderExceptions) {
                Write-Host " - $($loader.Message)"
            }
        } else {
            $_ | Format-List * -Force
        }
    }
} else {
    Write-Host 'MailKit assembly not found in current domain.'
}
