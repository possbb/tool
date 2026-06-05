@echo off
setlocal
cd /d "%~dp0"
set "PY=C:\Users\weiyuchen\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PY%" set "PY=python"
start "" http://127.0.0.1:5000
"%PY%" "%~dp0app.py"
if errorlevel 1 (
  echo.
  echo 如果提示缺少依赖，请先运行：pip install -r requirements.txt
)
pause
endlocal
