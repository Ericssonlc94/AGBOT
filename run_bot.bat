@echo off
chcp 65001 > nul
call venv\Scripts\activate.bat

echo Verifying Python environment...
venv\Scripts\python.exe --version
echo Python executable path:
for %%i in (venv\Scripts\python.exe) do echo %%~fi

venv\Scripts\python.exe bot.py
echo Dados coletados com sucesso!
pause