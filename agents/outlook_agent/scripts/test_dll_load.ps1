try {
    Add-Type -Path './lib/MimeKit.dll' -ErrorAction Stop
} catch {
    Write-Host "MimeKit LoaderExceptions:"
    $_.Exception.LoaderExceptions | ForEach-Object { Write-Host "  $_" }
}

try {
    Add-Type -Path './lib/MailKit.dll' -ErrorAction Stop
} catch {
    Write-Host "MailKit LoaderExceptions:"
    $_.Exception.LoaderExceptions | ForEach-Object { Write-Host "  $_" }
}
