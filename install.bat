@echo off
setlocal

set "TARGET=%~1"
if "%TARGET%"=="" set "TARGET=C:\tools"
if "%TARGET%"=="-Clean" goto :clean
if "%TARGET%"=="-clean" goto :clean
if "%TARGET%"=="/clean" goto :clean

if not exist "%TARGET%" mkdir "%TARGET%"
copy /Y "%~dp0pptx2jpg.py" "%TARGET%\pptx2jpg.py" >nul
echo Installed %TARGET%\pptx2jpg.py
goto :eof

:clean
set "TARGET2=%~2"
if "%TARGET2%"=="" set "TARGET2=C:\tools"
if exist "%TARGET2%\pptx2jpg.py" (
    del /F "%TARGET2%\pptx2jpg.py"
    echo Removed %TARGET2%\pptx2jpg.py
) else (
    echo %TARGET2%\pptx2jpg.py not found, nothing to remove
)
