param(
  [string]$DllPath = "$(Split-Path -Parent $PSScriptRoot)\\build\\Release\\Naiv4VibeThumbnailProvider.dll"
)

$resolvedDllPath = Resolve-Path -LiteralPath $DllPath -ErrorAction SilentlyContinue
if (-not $resolvedDllPath) {
  throw "DLL not found: $DllPath"
}
$resolvedDllPath = $resolvedDllPath.Path

$regsvr32 = Join-Path $env:WINDIR "System32\\regsvr32.exe"

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$runningAsAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $runningAsAdmin) {
  Write-Warning "PowerShell is not running as Administrator. COM unregistration may fail with access denied."
}

Write-Host "Unregistering $resolvedDllPath using $regsvr32"
$proc = Start-Process -FilePath $regsvr32 -ArgumentList @('/s', '/u', $resolvedDllPath) -Wait -PassThru -NoNewWindow
$exitCode = $proc.ExitCode

if ($exitCode -ne 0) {
  $commonHint = switch ($exitCode) {
    3 { "Path not found. Confirm DLL path is correct." }
    5 { "Access denied. Re-run in Administrator PowerShell." }
    126 { "Dependent module was not found. Install required VC++ runtime / check dependencies." }
    193 { "Bad EXE format. Usually 32/64-bit mismatch." }
    default { "Check DLL architecture and dependencies, then run regsvr32 manually for full error UI." }
  }

  throw "regsvr32 /u failed with exit code $exitCode. $commonHint`nManual check: `"$regsvr32`" /u `"$resolvedDllPath`""
}

Write-Host "Done."
