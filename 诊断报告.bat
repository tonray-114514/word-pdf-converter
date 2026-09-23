@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY=%~dp0env\Scripts\python.exe"

if not exist "%PY%" (
    echo.
    echo  [错误] 未找到程序自带的运行环境 env\Scripts\python.exe
    echo  请先双击运行「安装依赖.bat」。
    echo.
    pause
    exit /b 1
)

rem 若自带环境缺少关键组件，先补装，避免误报「未检测到 Word」
"%PY%" -c "import docx, pymupdf, win32com.client" >nul 2>&1
if errorlevel 1 (
    echo  正在补齐运行环境组件，请稍候...
    "%PY%" -m pip install --no-cache-dir --disable-pip-version-check python-docx PyMuPDF pywin32 tkinterdnd2
    echo.
)

echo ============================================
echo   引擎诊断报告
echo ============================================
echo.

"%PY%" "%~dp0cli.py" --diagnose dummy

echo.
echo ============================================
echo  报告已生成，可复制上面内容反馈
echo ============================================
echo.
pause
