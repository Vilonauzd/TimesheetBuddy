# Enforce TLS 1.2 for secure downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Paths
$scriptPath = "D:\timesheetbuddy.ps1"
$outputPath = [System.IO.Path]::ChangeExtension($scriptPath, ".exe")

# Ensure ps2exe is installed
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Installing ps2exe from PSGallery..."
    Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber
}

# Import ps2exe module
Import-Module ps2exe -Force

# Compile to EXE
Write-Host "`nCompiling to EXE..."
Invoke-ps2exe -inputFile $scriptPath -outputFile $outputPath -noConsole -requireAdmin:$false

Write-Host "`nDone! Executable created at:`n$outputPath"
