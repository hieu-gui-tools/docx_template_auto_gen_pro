@echo off
cd /d "%~dp0"

where uv >nul 2>&1
if errorlevel 1 (
    set "UV=C:\Users\drhie\.local\bin\uv.exe"
) else (
    set "UV=uv"
)

start "" "%UV%" run python -m word_template_pro
exit
