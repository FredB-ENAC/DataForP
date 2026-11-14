@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
 set "PY=py -3"
) else (
 where python >nul 2>nul
 if errorlevel 1 (
  echo Python 3.10 ou superieur est necessaire.
  pause
  exit /b 1
 )
 set "PY=python"
)
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
 echo Python 3.10 ou superieur est necessaire.
 pause
 exit /b 1
)
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
 echo openpyxl est absent. Installation Internet sans transmission des donnees sources.
 set /p "REP=Installer maintenant ? Taper O puis Entree : "
 call :installer
 if errorlevel 1 exit /b 1
)
if /I "%~1"=="--tests" (
 %PY% consolidation.py --autotest
 pause
) else (
 %PY% consolidation.py %*
 if errorlevel 1 pause
)
exit /b
:installer
if /I not "%REP%"=="O" (
 echo Installation annulee.
 pause
 exit /b 1
)
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
 echo Installation impossible.
 pause
 exit /b 1
)
exit /b 0
