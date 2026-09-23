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

rem 环境不完整时自动补装
"%PY%" -c "import docx, pymupdf, win32com.client" >nul 2>&1
if errorlevel 1 (
    echo  正在补齐运行环境组件，请稍候...
    "%PY%" -m pip install --no-cache-dir --disable-pip-version-check python-docx PyMuPDF pywin32 tkinterdnd2
    echo.
)

echo.
echo  正在启动手机网页服务...
echo  （首次启动需要几秒，请稍候）
echo.

"%PY%" "%~dp0server.py" %*

echo.
echo  服务已停止。
pause
