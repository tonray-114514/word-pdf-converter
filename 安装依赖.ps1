# Word ⇄ PDF 互转工具 —— 环境安装脚本
#
# 用法：在 PowerShell 中执行
#     powershell -ExecutionPolicy Bypass -File .\安装依赖.ps1
#
# 作用：在程序目录下创建独立虚拟环境 env，并安装全部依赖。
#       依赖只写入本目录的 env，不会影响系统 Python 环境。

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Word - PDF 互转工具  环境初始化" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# 找到可用的 Python
$python = $null
foreach ($candidate in @('python', 'py')) {
    $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
    if ($cmd) { $python = $cmd.Source; break }
}
if (-not $python) {
    Write-Host "[错误] 未检测到 Python。" -ForegroundColor Red
    Write-Host "请先安装 Python 3.9 以上版本： https://www.python.org/downloads/"
    Write-Host "安装时务必勾选 Add Python to PATH。"
    exit 1
}
Write-Host "使用 Python: $python"
& $python -V
Write-Host ""

$venvPy = Join-Path $here 'env\Scripts\python.exe'

if (-not (Test-Path $venvPy)) {
    Write-Host "[1/3] 创建独立运行环境 env ..." -ForegroundColor Yellow
    & $python -m venv env
    if ($LASTEXITCODE -ne 0) { Write-Host "[错误] 创建虚拟环境失败。" -ForegroundColor Red; exit 1 }
} else {
    Write-Host "[1/3] 运行环境已存在，跳过创建。"
}

Write-Host ""
Write-Host "[2/3] 安装依赖 ..." -ForegroundColor Yellow
& $venvPy -m pip install --upgrade pip --disable-pip-version-check -q
& $venvPy -m pip install --no-cache-dir --disable-pip-version-check `
    python-docx PyMuPDF pywin32 tkinterdnd2
if ($LASTEXITCODE -ne 0) {
    Write-Host "[错误] 依赖安装失败，请检查网络连接后重试。" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "[3/3] 自检 ..." -ForegroundColor Yellow
& $venvPy -c "import docx, pymupdf, win32com.client, tkinterdnd2; print('  依赖自检通过')"
if ($LASTEXITCODE -ne 0) { Write-Host "[错误] 自检失败。" -ForegroundColor Red; exit 1 }

# 顺带报告转换引擎情况
& $venvPy -c @"
import sys, os
sys.path.insert(0, os.getcwd())
import converter
print('  Word -> PDF 引擎:', converter.describe_engine())
"@

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  初始化完成！双击「启动程序.bat」开始使用" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
