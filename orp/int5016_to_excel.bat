@echo off
setlocal EnableExtensions
rem ===========================================================================
rem  INT5016 ORP file - validate and convert to Excel
rem
rem  Usage:
rem    int5016_to_excel.bat                       all files in .\input -> .\output
rem                                               (asks for a file if .\input is empty)
rem    int5016_to_excel.bat file1.txt [file2 ...] given files (drag-and-drop works)
rem    int5016_to_excel.bat C:\some\folder        all data files in that folder
rem  Any extension is accepted (.txt, .dat, none); .xlsx/.py/.bat/.xml are skipped.
rem
rem  Set INT5016_NOPAUSE=1 to skip the final pause (scheduled jobs).
rem  Exit code: 0 = PASS, 2 = validation errors, 1 = setup/processing problem
rem ===========================================================================

set "SCRIPT_DIR=%~dp0"
set "PY_SCRIPT=%SCRIPT_DIR%int5016_to_excel.py"
set "OUT_DIR=%SCRIPT_DIR%output"
set "IN_DIR=%SCRIPT_DIR%input"

if not exist "%PY_SCRIPT%" (
    echo ERROR: %PY_SCRIPT% not found.
    set "RC=1" & goto :end
)

rem --- locate a working Python 3 ---------------------------------------------
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python --version >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo ERROR: Python 3 not found. Install it from https://www.python.org/downloads/
    echo        and tick "Add python.exe to PATH".
    set "RC=1" & goto :end
)

rem --- make sure openpyxl is installed ---------------------------------------
%PY% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl ...
    %PY% -m pip install --user --quiet openpyxl
    if errorlevel 1 (
        echo ERROR: could not install openpyxl. Run:  %PY% -m pip install openpyxl
        set "RC=1" & goto :end
    )
)

rem --- run --------------------------------------------------------------------
if not "%~1"=="" (
    %PY% "%PY_SCRIPT%" %* -o "%OUT_DIR%"
    goto :result
)

if not exist "%IN_DIR%" mkdir "%IN_DIR%"
rem any file in .\input counts (no extension, .txt, .dat ...)
set "HAS_FILES="
for %%F in ("%IN_DIR%\*") do set "HAS_FILES=1"
if defined HAS_FILES (
    echo No input given - processing files in "%IN_DIR%"
    %PY% "%PY_SCRIPT%" "%IN_DIR%" -o "%OUT_DIR%"
    goto :result
)

echo The input folder "%IN_DIR%" is empty.
echo Copy the INT5016 file there and run again, or drag the file into this
echo window now and press Enter.
set "USER_IN="
set /p "USER_IN=File or folder: "
if not defined USER_IN (
    echo No file given.
    set "RC=1" & goto :end
)
set "USER_IN=%USER_IN:"=%"
%PY% "%PY_SCRIPT%" "%USER_IN%" -o "%OUT_DIR%"

:result
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" echo RESULT: PASSED - workbook(s) in "%OUT_DIR%"
if "%RC%"=="2" echo RESULT: VALIDATION ERRORS - see the Validation sheet in "%OUT_DIR%"
if "%RC%"=="1" echo RESULT: PROBLEM - see messages above

:end
if not "%INT5016_NOPAUSE%"=="1" pause
exit /b %RC%
