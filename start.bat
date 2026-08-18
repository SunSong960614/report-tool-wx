@echo off
setlocal
set "PYTHON_EXE=C:\Users\PC\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON_EXE%" (
  echo Bundled Python not found. Please reopen this tool from the Codex workspace.
  pause
  exit /b 1
)
start "学校测评报告合成工具" /min "%PYTHON_EXE%" "%~dp0server.py"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:18966"
endlocal
