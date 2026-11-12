# WordOllama 案例库测试版：卸载当前用户的安装。设置文件（appconfig.json）保留不动。
$ErrorActionPreference = 'Stop'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host '卸载失败：Word 正在运行。请先关闭所有 Word 窗口。' -ForegroundColor Red
    exit 1
}

$target = Join-Path $env:LOCALAPPDATA 'WordOllama\AddIn'
$manifestUri = ([Uri](Join-Path $target 'WordOllama.vsto')).AbsoluteUri

Remove-Item -Recurse -Force 'HKCU:\Software\Microsoft\Office\Word\Addins\WordOllama' -ErrorAction SilentlyContinue

$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
if (Test-Path $inclusionRoot) {
    Get-ChildItem $inclusionRoot | Where-Object { (Get-ItemProperty $_.PSPath).Url -eq $manifestUri } | Remove-Item -Recurse -Force
}

if (Test-Path $target) { Remove-Item -Recurse -Force $target }

Write-Host '已卸载 WordOllama。' -ForegroundColor Green
