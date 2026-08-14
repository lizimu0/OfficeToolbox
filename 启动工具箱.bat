@echo off
cd /d %~dp0
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw -m office_toolbox.main
    exit
)
echo 未检测到 Python,请先安装 Python(安装时勾选 Add to PATH)
echo 或直接使用 dist\Office自动化工具箱\Office自动化工具箱.exe
pause
