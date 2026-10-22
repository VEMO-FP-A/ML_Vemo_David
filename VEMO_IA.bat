@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PY=
python --version >nul 2>nul && set PY=python
if not defined PY py --version >nul 2>nul && set PY=py
if not defined PY (
  echo.
  echo No encuentro Python en esta computadora.
  echo Instalalo desde https://www.python.org/downloads/  ^(marca "Add python.exe to PATH"^)
  echo o desde Microsoft Store ^(busca "Python 3.12"^). Luego vuelve a dar doble clic.
  echo.
  pause
  exit /b 1
)
%PY% -c "import openpyxl, pptx" >nul 2>nul || (
  echo Instalando librerias por unica vez...
  %PY% -m pip install --user openpyxl python-pptx
)
%PY% vemo_ia.py %*
pause
