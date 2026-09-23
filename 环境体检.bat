@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY=%~dp0env\Scripts\python.exe"

echo ============================================
echo  Word - PDF 互转工具   运行环境体检
echo ============================================
echo.

if not exist "%PY%" (
    echo  [错误] 未找到自带运行环境 env\Scripts\python.exe
    echo  请先双击运行「安装依赖.bat」。
    echo.
    pause
    exit /b 1
)

echo  正在检查各项组件，请稍候...
echo.

"%PY%" "%~dp0tools\healthcheck.py"

echo.
echo ============================================
echo  若上面全部显示「通过」，程序即可正常使用。
echo  出现「失败」的项目请连同本窗口内容反馈。
echo ============================================
echo.
pause
