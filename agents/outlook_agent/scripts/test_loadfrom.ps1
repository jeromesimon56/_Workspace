try {
    [Reflection.Assembly]::LoadFrom((Resolve-Path './lib/System.Runtime.dll'))
    [Reflection.Assembly]::LoadFrom((Resolve-Path './lib/MimeKit.dll'))
    [Reflection.Assembly]::LoadFrom((Resolve-Path './lib/MailKit.dll'))
    Write-Host "OK MimeKit et MailKit charges"
} catch {
    Write-Host "Erreur: $_"
    $_.Exception.LoaderExceptions | ForEach-Object { Write-Host "  Loader: $_" }
}
