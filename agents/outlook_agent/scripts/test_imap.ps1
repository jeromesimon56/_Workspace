[Reflection.Assembly]::LoadFrom((Resolve-Path 'System.Runtime.dll')) | Out-Null
[Reflection.Assembly]::LoadFrom((Resolve-Path 'MimeKit.dll')) | Out-Null
[Reflection.Assembly]::LoadFrom((Resolve-Path 'MailKit.dll')) | Out-Null
try {
    $client = [MailKit.Net.Imap.ImapClient]::new()
    Write-Host "Created ImapClient"
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    if ($_.Exception.InnerException) { Write-Host "Inner: $($_.Exception.InnerException.Message)" }
}
