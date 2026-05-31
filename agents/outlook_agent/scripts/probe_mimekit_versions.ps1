[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$temp = Join-Path $env:TEMP 'mimekit_probe'
New-Item -Path $temp -ItemType Directory -Force | Out-Null
$pkg = Join-Path $temp 'MimeKit.3.4.0.nupkg'
Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/MimeKit/3.4.0' -OutFile $pkg -UseBasicParsing
$zip = $pkg -replace '\.nupkg$','.zip'
Copy-Item -Path $pkg -Destination $zip -Force
Expand-Archive -Path $zip -DestinationPath $temp -Force
Get-ChildItem -Path (Join-Path $temp 'lib') -Recurse | Where-Object { $_.Extension -eq '.dll' } | Select-Object FullName
Remove-Item -Path $temp -Recurse -Force
