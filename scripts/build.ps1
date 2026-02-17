param(
  [string]$RepoRoot = "$(Split-Path -Parent $PSScriptRoot)",
  [string]$BuildDir = "build",
  [string]$Config = "Release"
)

$repoPath = (Resolve-Path $RepoRoot).Path
$cmakeFile = Join-Path $repoPath "CMakeLists.txt"

if (-not (Test-Path $cmakeFile)) {
  throw "CMakeLists.txt not found at: $cmakeFile`nPlease set -RepoRoot to this repository root."
}

Push-Location $repoPath
try {
  Write-Host "Configuring in $repoPath ..."
  cmake -S . -B $BuildDir -A x64
  if ($LASTEXITCODE -ne 0) {
    throw "cmake configure failed with exit code $LASTEXITCODE"
  }

  Write-Host "Building $Config ..."
  cmake --build $BuildDir --config $Config
  if ($LASTEXITCODE -ne 0) {
    throw "cmake build failed with exit code $LASTEXITCODE"
  }

  $dllPath = Join-Path $repoPath "$BuildDir\$Config\Naiv4VibeThumbnailProvider.dll"
  Write-Host "Build done. DLL: $dllPath"
} finally {
  Pop-Location
}
