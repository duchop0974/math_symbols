param(
    [Parameter(Mandatory = $true)][string]$User,
    [string]$Repo = "math_symbols"
)

$manifest = Join-Path $PSScriptRoot "manifest.xml"
$content = Get-Content $manifest -Raw
$content = $content -replace 'YOUR-USERNAME\.github\.io/math-symbols', "$User.github.io/$Repo"
$content = $content -replace 'YOUR-USERNAME\.github\.io', "$User.github.io"
$content = $content -replace 'github\.com/YOUR-USERNAME/math-symbols', "github.com/$User/$Repo"
Set-Content $manifest -Value $content -NoNewline -Encoding UTF8

Write-Host "manifest.xml updated -> https://$User.github.io/$Repo/"
