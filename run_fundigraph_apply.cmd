@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PYTHON_EXE%" goto run
set "PYTHON_EXE=python"

:run
"%PYTHON_EXE%" "%~dp0fundigraph_review_workflow.py" apply
endlocal
