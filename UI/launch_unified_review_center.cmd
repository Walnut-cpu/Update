@echo off
setlocal
cd /d "%~dp0"
set "PYTHONW_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if exist "%PYTHONW_EXE%" goto runw
set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PYTHON_EXE%" goto run
set "PYTHONW_EXE=pythonw"
where "%PYTHONW_EXE%" >nul 2>nul
if %ERRORLEVEL%==0 goto runw
set "PYTHON_EXE=python"

:runw
"%PYTHONW_EXE%" "%~dp0unified_review_center.py"
goto end

:run
"%PYTHON_EXE%" "%~dp0unified_review_center.py"

:end
endlocal
