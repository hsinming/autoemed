# Build autoemed Windows executable with Nuitka
# All Nuitka options are defined via nuitka-project: comments in main.py

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Checking dependencies..."

uv run python -m nuitka --version 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    $reply = Read-Host "Nuitka not found. Install it now with 'uv add nuitka'? (y/n)"
    if ($reply -match '^[Yy]') {
        uv add nuitka
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install nuitka."
            exit 1
        }
    } else {
        Write-Error "Nuitka is required. Aborting."
        exit 1
    }
}

Write-Host "Building..."
uv run python -m nuitka main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build complete. Output: build\"
} else {
    Write-Error "Build failed (exit code $LASTEXITCODE)"
    exit $LASTEXITCODE
}
