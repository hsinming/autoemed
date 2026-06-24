# Build autoemed Windows executable with Nuitka
# All Nuitka options are defined via nuitka-project: comments in main.py

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Checking dependencies..."

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Error "Python not found. Please install Python and add it to PATH."
    exit 1
}

python -m nuitka --version 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Error "Nuitka not found. Run: pip install nuitka"
    exit 1
}

Write-Host "Building..."
python -m nuitka main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build complete. Output: build\"
} else {
    Write-Error "Build failed (exit code $LASTEXITCODE)"
    exit $LASTEXITCODE
}
