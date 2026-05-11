@echo off
setlocal

if "%~1"=="" (
    echo Drag one or more PDF files onto this BAT file.
    echo.
    echo Or run this command in the current folder:
    echo python "%~dp0pdf_beamer_to_pptx.py" "20260516.pdf"
    echo.
    pause
    exit /b 1
)

:convert_loop
if "%~1"=="" goto done
echo.
echo Converting: %~1
python "%~dp0pdf_beamer_to_pptx.py" "%~1"
if errorlevel 1 (
    echo.
    echo The file above failed to convert.
)
shift
goto convert_loop

:done
echo.
echo All done.
pause
