@echo off
setlocal EnableExtensions
rem ===========================================================================
rem  INT5016 ORP file - validate and convert to Excel
rem
rem  Usage:
rem    int5016_to_excel.bat                       every data file in this folder
rem                                               (asks for a file if none found)
rem    int5016_to_excel.bat file1 [file2 ...]     given files (drag-and-drop works)
rem    int5016_to_excel.bat C:\some\folder        every data file in that folder
rem
rem  Input files need no extension (e.g. CO050102026). Files ending in
rem  .xlsx/.py/.bat/.xml/.xsl etc. are ignored.
rem  Workbooks are written to the "output" subfolder as <name>_validated.xlsx
rem
rem  Set INT5016_NOPAUSE=1 to skip the final pause (scheduled jobs).
rem  Exit code: 0 = PASS, 2 = validation errors, 1 or 3 = setup/input problem
rem ===========================================================================

rem folder of this .bat without the trailing backslash
set "BASE=%~dp0"
set "BASE=%BASE:~0,-1%"
set "PY_SCRIPT=%BASE%\int5016_to_excel.py"
set "OUT_DIR=%BASE%\output"

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

echo Looking for INT5016 files in "%BASE%" ...
%PY% "%PY_SCRIPT%" "%BASE%" -o "%OUT_DIR%"
if not errorlevel 3 goto :result

echo.
echo Nothing to process. Copy the INT5016 file into "%BASE%" and run again,
echo or drag the file into this window now and press Enter.
set "USER_IN="
set /p "USER_IN=File or folder: "
if not defined USER_IN (
    echo No file given.
    set "RC=3" & goto :end
)
set "USER_IN=%USER_IN:"=%"
%PY% "%PY_SCRIPT%" "%USER_IN%" -o "%OUT_DIR%"

:result
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" echo RESULT: PASSED - workbook(s) in "%OUT_DIR%"
if "%RC%"=="2" echo RESULT: VALIDATION ERRORS - see the Validation sheet in "%OUT_DIR%"
if "%RC%"=="1" echo RESULT: PROBLEM - see messages above
if "%RC%"=="3" echo RESULT: NO INPUT FILE FOUND

:end
if not "%INT5016_NOPAUSE%"=="1" pause
exit /b %RC%
