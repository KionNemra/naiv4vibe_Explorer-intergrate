param(
  [string]$DllPath = "$(Split-Path -Parent $PSScriptRoot)\\build\\Release\\Naiv4VibeThumbnailProvider.dll"
)

if (-not (Test-Path $DllPath)) {
  throw "DLL not found: $DllPath"
}

$regsvr32 = Join-Path $env:WINDIR "System32\\regsvr32.exe"
Write-Host "Registering $DllPath using $regsvr32"
& $regsvr32 /s $DllPath
if ($LASTEXITCODE -ne 0) {
  throw "regsvr32 failed with exit code $LASTEXITCODE"
}

Write-Host "Done. Restart Explorer or clear thumbnail cache to refresh."
