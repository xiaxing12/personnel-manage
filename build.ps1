$ErrorActionPreference = "Stop"

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "未找到 .venv。请先按照 README 创建虚拟环境并安装依赖。"
}

& $python -m pip install --upgrade pyinstaller
if ($LASTEXITCODE -ne 0) {
    throw "安装 PyInstaller 失败。"
}

& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name "新人员管理系统" `
    --exclude-module numpy `
    --exclude-module PIL `
    --exclude-module pandas `
    --exclude-module matplotlib `
    --exclude-module scipy `
    --exclude-module pytest `
    "$PSScriptRoot\main.py"
if ($LASTEXITCODE -ne 0) {
    throw "程序打包失败。"
}

Write-Host "打包完成：$PSScriptRoot\dist\新人员管理系统"
