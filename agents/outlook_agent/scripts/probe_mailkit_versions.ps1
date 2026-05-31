[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$temp = Join-Path $env:TEMP 'mailkit_probe'
New-Item -Path $temp -ItemType Directory -Force | Out-Null
$pkg = Join-Path $temp 'MailKit.3.4.0.nupkg'
Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/MailKit/3.4.0' -OutFile $pkg -UseBasicParsing
$zip = $pkg -replace '\.nupkg$','.zip'
Copy-Item -Path $pkg -Destination $zip -Force
Expand-Archive -Path $zip -DestinationPath $temp -Force
Get-ChildItem -Path (Join-Path $temp 'lib') -Recurse | Where-Object { $_.Name -in @('MailKit.dll','MimeKit.dll') } | Select-Object FullName
Remove-Item -Path $temp -Recurse -Force
