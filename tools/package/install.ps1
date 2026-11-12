# WordOllama 案例库测试版：安装到当前用户，不需要管理员权限。
$ErrorActionPreference = 'Stop'

function Fail([string]$message) {
    Write-Host ''
    Write-Host "安装失败：$message" -ForegroundColor Red
    exit 1
}

$source = Join-Path $PSScriptRoot 'WordOllama'
$certPath = Join-Path $PSScriptRoot 'WordOllama.cer'
if (-not (Test-Path (Join-Path $source 'WordOllama.vsto')) -or -not (Test-Path $certPath)) {
    Fail '找不到安装文件。请先把 zip 完整解压，再运行“安装.cmd”。'
}

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Fail 'Word 正在运行。请先关闭所有 Word 窗口，再重新运行“安装.cmd”。'
}

$netRelease = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
if (-not $netRelease -or $netRelease -lt 528040) {
    Fail '需要 .NET Framework 4.8。请在微软官网下载安装 .NET Framework 4.8 Runtime 后重试。'
}

$vstoKeys = 'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
$vstoFiles = "$env:CommonProgramFiles\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe", "${env:CommonProgramFiles(x86)}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
if (-not (($vstoKeys | Where-Object { Test-Path $_ }) -or ($vstoFiles | Where-Object { $_ -and (Test-Path $_) }))) {
    Write-Host '提示：没有检测到 VSTO 运行时（Visual Studio 2010 Tools for Office Runtime）。' -ForegroundColor Yellow
    Write-Host '如果安装后 Word 里看不到 WordOllama，请在微软官网下载安装这个运行时，再重启 Word。' -ForegroundColor Yellow
}

$target = Join-Path $env:LOCALAPPDATA 'WordOllama\AddIn'
if (Test-Path $target) { Remove-Item -Recurse -Force $target }
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item -Recurse -Force (Join-Path $source '*') $target
Get-ChildItem $target -Recurse -File | Unblock-File

$manifestUri = ([Uri](Join-Path $target 'WordOllama.vsto')).AbsoluteUri

# 把本插件的签名证书加入 VSTO 信任列表，Word 加载时不再弹出“发布者无法验证”。
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $certPath
$publicKey = $cert.PublicKey.Key.ToXmlString($false)
$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
if (Test-Path $inclusionRoot) {
    Get-ChildItem $inclusionRoot | Where-Object { (Get-ItemProperty $_.PSPath).Url -eq $manifestUri } | Remove-Item -Recurse -Force
}
$entry = Join-Path $inclusionRoot ('{' + [Guid]::NewGuid().ToString() + '}')
New-Item -Force -Path $entry | Out-Null
New-ItemProperty -Path $entry -Name Url -Value $manifestUri -PropertyType String -Force | Out-Null
New-ItemProperty -Path $entry -Name PublicKey -Value $publicKey -PropertyType String -Force | Out-Null

$addinKey = 'HKCU:\Software\Microsoft\Office\Word\Addins\WordOllama'
New-Item -Force -Path $addinKey | Out-Null
New-ItemProperty -Path $addinKey -Name FriendlyName -Value 'WordOllama' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name Description -Value 'AI-assisted tools for Word documents' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name Manifest -Value "$manifestUri|vstolocal" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host ''
Write-Host "安装完成：$target" -ForegroundColor Green
Write-Host '现在打开 Word，功能区应出现 WordOllama 选项卡。'
