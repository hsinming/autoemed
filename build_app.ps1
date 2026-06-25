# Build autoemed Windows executable with Nuitka
# All Nuitka options are defined via nuitka-project: comments in main.py

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Building..."
uv run python -m nuitka main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build complete. Output: build\"
} else {
    Write-Error "Build failed (exit code $LASTEXITCODE)"
    exit $LASTEXITCODE
}
