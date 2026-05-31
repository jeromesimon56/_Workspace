[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$temp = Join-Path $env:TEMP 'mailkit_nuspec'
New-Item -Path $temp -ItemType Directory -Force | Out-Null
$pkg = Join-Path $temp 'MailKit.3.4.0.nupkg'
Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/MailKit/3.4.0' -OutFile $pkg -UseBasicParsing
$zip = $pkg -replace '\.nupkg$','.zip'
Copy-Item -Path $pkg -Destination $zip -Force
Expand-Archive -Path $zip -DestinationPath $temp -Force
Get-ChildItem -Path $temp -Filter '*.nuspec' -Recurse | ForEach-Object { Get-Content $_.FullName }
Remove-Item -Path $temp -Recurse -Force
