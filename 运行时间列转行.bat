@echo off
setlocal
python "%~dp0time_columns_to_rows.py" %*
if errorlevel 1 (
  echo.
  echo 如果提示缺少 pandas/openpyxl，请先运行：pip install pandas openpyxl
)
endlocal
