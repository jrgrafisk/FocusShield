# PowerShell script to create a release package for FocusShield
# Run this from the repository root on Windows

$exePath = "FocusShield\bin\Release\net8.0-windows\FocusShield.exe"
$readmePath = "README.md"
$licensePath = "LICENSE"
$packageName = "FocusShield-v1.0.0.zip"
$tempDir = "FocusShield-v1.0.0"

# Clean up previous package
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
if (Test-Path $packageName) {
    Remove-Item $packageName -Force
}

# Create temporary directory structure
New-Item -ItemType Directory -Path $tempDir | Out-Null

# Copy files to package
Copy-Item $exePath -Destination $tempDir\FocusShield.exe
Copy-Item $readmePath -Destination $tempDir\README.md
Copy-Item $licensePath -Destination $tempDir\LICENSE.txt

# Create a simple INSTALL.txt with quick instructions
@"
FOCUSSHIELD - QUICK START
==========================

1. Extract this folder to any location (e.g., C:\Program Files or Downloads)
2. Double-click FocusShield.exe to run it
3. Right-click the FocusShield icon in your system tray for options
4. (Optional) Select "Load on Boot" to auto-start on Windows startup

REQUIREMENTS:
- Windows 10 or later
- .NET 8 Runtime (downloads automatically if missing, or get it from dotnet.microsoft.com)

For more information, see README.md
"@ | Set-Content -Path $tempDir\INSTALL.txt -Encoding UTF8

# Create the ZIP file
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $packageName)

# Clean up temporary directory
Remove-Item $tempDir -Recurse -Force

Write-Host "✓ Package created: $packageName" -ForegroundColor Green
Write-Host "Ready to upload to GitHub Releases" -ForegroundColor Green
