<#
.SYNOPSIS
    Đăng ký / gỡ add-in "Trợ Lý Soạn Đề" vào Word cho người dùng hiện tại.

.DESCRIPTION
    Word đọc danh sách add-in sideload từ khoá registry WEF\Developer:
    tên giá trị là Id trong manifest, dữ liệu là đường dẫn tới file manifest.
    Không cần quyền admin vì chỉ ghi vào HKCU.

.EXAMPLE
    .\Register-Addin.ps1 -ManifestPath ..\manifest.xml
    .\Register-Addin.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [string]$ManifestPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "manifest.xml"),
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$DeveloperKey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer"

function Get-ManifestId {
    param([string]$Path)
    [xml]$xml = Get-Content -LiteralPath $Path -Raw
    $id = $xml.OfficeApp.Id
    if (-not $id) { throw "Không đọc được <Id> trong manifest: $Path" }
    return $id
}

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Warning "Word đang chạy. Hãy đóng hẳn Word rồi mở lại sau khi chạy script này."
}

if ($Uninstall) {
    if (-not (Test-Path $DeveloperKey)) {
        Write-Host "Không có add-in nào được đăng ký."
        return
    }
    $props = Get-ItemProperty -Path $DeveloperKey
    $removed = 0
    foreach ($name in $props.PSObject.Properties.Name) {
        if ($name -like "PS*") { continue }
        if ($props.$name -like "*math_symbols*" -or $props.$name -like "*MathSymbols*" -or
            $props.$name -like "*TroLySoanDe*") {
            Remove-ItemProperty -Path $DeveloperKey -Name $name
            Write-Host "Đã gỡ: $name"
            $removed++
        }
    }
    if ($removed -eq 0) { Write-Host "Không tìm thấy add-in Trợ Lý Soạn Đề đã đăng ký." }
    else { Write-Host "Gỡ xong. Khởi động lại Word để áp dụng." }
    return
}

$ManifestPath = (Resolve-Path -LiteralPath $ManifestPath).Path
$id = Get-ManifestId -Path $ManifestPath

if (-not (Test-Path $DeveloperKey)) {
    New-Item -Path $DeveloperKey -Force | Out-Null
}

New-ItemProperty -Path $DeveloperKey -Name $id -Value $ManifestPath -PropertyType String -Force | Out-Null

Write-Host "Đã đăng ký add-in:"
Write-Host "  Id       : $id"
Write-Host "  Manifest : $ManifestPath"
Write-Host "Mở lại Word -> nút 'Soạn đề' sẽ xuất hiện ở tab Home."
