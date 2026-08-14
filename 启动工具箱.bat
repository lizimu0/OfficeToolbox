@echo off
cd /d %~dp0
set "EXE=%~dp0dist\Office自动化工具箱\Office自动化工具箱.exe"
if exist "%EXE%" (
    start "" "%EXE%"
    exit
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw -m office_toolbox.main
    exit
)
echo 未找到打包程序,也未检测到 Python
echo 请先安装 Python(安装时勾选 Add to PATH)并执行 pip install -r requirements.txt
pause
