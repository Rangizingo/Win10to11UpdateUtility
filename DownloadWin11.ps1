[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'
$url = "https://go.microsoft.com/fwlink/?linkid=2171764"
$out = "C:\Temp\Windows11InstallationAssistant.exe"
Write-Host "[DEBUG] Starting download from $url"
Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
Write-Host "[DEBUG] Download complete. File saved to $out"
