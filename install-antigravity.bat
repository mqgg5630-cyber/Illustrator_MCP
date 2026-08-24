@echo off
REM ============================================
REM Illustrator MCP -> Google Antigravity installer
REM ============================================
REM Writes an absolute-path stdio entry into Antigravity's mcp_config.json
REM (global: %%USERPROFILE%%\.gemini\config\mcp_config.json,
REM  workspace: %%~dp0.agents\mcp_config.json).
REM
REM Any extra arguments are passed through, e.g.:
REM   install-antigravity.bat --scope workspace
REM   install-antigravity.bat --verify
REM   install-antigravity.bat --uninstall

setlocal

set SCRIPT_DIR=%~dp0

REM Locate a Python interpreter. Antigravity does not inherit your shell PATH,
REM so the installer records whichever interpreter runs it (absolute path).
set PYTHON_CMD=
where python >nul 2>&1 && set PYTHON_CMD=python
if not defined PYTHON_CMD (
    where py >nul 2>&1 && set PYTHON_CMD=py -3
)
if not defined PYTHON_CMD (
    echo ERROR: Python was not found on PATH.
    echo Install Python 3.10+ ^(https://www.python.org/downloads/^) or activate
    echo your conda environment first, then run this script again.
    pause
    exit /b 1
)

echo.
echo =============================================
echo  Illustrator MCP - Antigravity Installer
echo =============================================
echo  Interpreter: %PYTHON_CMD%
echo  Project:     %SCRIPT_DIR%.
echo.

REM Install the package first so `python -m illustrator_mcp.server` resolves.
%PYTHON_CMD% -m pip install -e "%SCRIPT_DIR%." >nul 2>&1
if %errorLevel% neq 0 (
    echo WARNING: "pip install -e ." did not complete cleanly.
    echo          The config will still be written; fix the install afterwards.
    echo.
)

%PYTHON_CMD% "%SCRIPT_DIR%scripts\install_antigravity.py" --project "%SCRIPT_DIR%." %*
set RESULT=%errorLevel%

echo.
if %RESULT% equ 0 (
    echo Done. Restart Antigravity, then check Agent panel - ... - MCP Servers.
) else (
    echo The installer reported problems - see the messages above.
)
echo.
pause
exit /b %RESULT%
