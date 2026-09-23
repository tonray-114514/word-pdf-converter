@echo off
setlocal
cd /d "%~dp0"

set "PY=%~dp0env\Scripts\pythonw.exe"
set "PYC=%~dp0env\Scripts\python.exe"

if not exist "%PYC%" goto :nodeps

rem 缺少依赖时先自动补装
"%PYC%" -c "import docx, pymupdf, win32com.client" >nul 2>&1
if errorlevel 1 goto :bootstrap

:launch
start "" "%PY%" "%~dp0app.py"
exit /b 0

:bootstrap
echo.
echo  正在安装程序所需组件，首次运行需要几分钟，请勿关闭本窗口...
echo.
"%PYC%" -m pip install --no-cache-dir --disable-pip-version-check python-docx PyMuPDF pywin32 tkinterdnd2
if errorlevel 1 goto :fail
echo.
echo  组件安装完成，正在启动...
timeout /t 2 >nul
goto :launch

:nodeps
echo.
echo  [错误] 未找到程序自带的运行环境：env\Scripts\python.exe
echo.
echo  请先双击运行「安装依赖.bat」完成初始化。
echo.
pause
exit /b 1

:fail
echo.
echo  [错误] 组件安装失败，请检查网络连接后重试。
echo  也可以手动运行「安装依赖.bat」。
echo.
pause
exit /b 1
