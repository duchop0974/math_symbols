# Biên dịch bộ cài: installer\TroLySoanDe.iss -> dist\TroLySoanDeSetup.exe
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$candidates = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
)
$iscc = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
    throw "Không tìm thấy ISCC.exe. Cài Inno Setup bằng: winget install --id JRSoftware.InnoSetup"
}

& $iscc (Join-Path $PSScriptRoot "TroLySoanDe.iss")
if ($LASTEXITCODE -ne 0) { throw "Biên dịch thất bại (exit $LASTEXITCODE)" }

$output = Join-Path (Split-Path $PSScriptRoot -Parent) "dist\TroLySoanDeSetup.exe"
Write-Host "`nĐã tạo: $output"
