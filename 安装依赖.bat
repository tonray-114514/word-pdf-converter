@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================
echo   Word - PDF 互转工具  环境初始化
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 goto :nopython

if not exist "env\Scripts\python.exe" (
    echo [1/3] 正在创建独立运行环境 env ...
    python -m venv env
    if errorlevel 1 goto :fail
) else (
    echo [1/3] 运行环境已存在，跳过创建。
)

echo.
echo [2/3] 正在升级安装工具 pip ...
"env\Scripts\python.exe" -m pip install --upgrade pip --disable-pip-version-check -q

echo.
echo [3/3] 正在安装依赖：python-docx / PyMuPDF / pywin32 / tkinterdnd2 ...
"env\Scripts\python.exe" -m pip install --no-cache-dir --disable-pip-version-check python-docx PyMuPDF pywin32 tkinterdnd2
if errorlevel 1 goto :fail

echo.
echo 正在自检 ...
"env\Scripts\python.exe" -c "import docx, pymupdf, win32com.client, tkinterdnd2; print('  依赖自检通过')"
if errorlevel 1 goto :fail

echo.
echo ============================================
echo   初始化完成！现在可以双击「启动程序.bat」
echo ============================================
echo.
pause
exit /b 0

:nopython
echo [错误] 未检测到 Python。
echo.
echo 请先到 https://www.python.org/downloads/ 下载安装 Python 3.9 以上版本，
echo 安装时务必勾选 "Add Python to PATH"，然后重新运行本脚本。
echo.
pause
exit /b 1

:fail
echo.
echo [错误] 初始化失败，请查看上面的提示信息。
echo.
pause
exit /b 1
