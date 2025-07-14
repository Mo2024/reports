:Start
@echo off
setlocal enabledelayedexpansion

:ask_type
set /p type=Enter report type (1 for stability, 2 for transactional): 

set year=%date:~10,4%



:: Get the numeric month from the system date
for /f "tokens=2 delims=/.- " %%a in ("%date%") do set MM=%%a

:: Remove leading zero if needed
if "%MM:~0,1%"=="0" set MM=%MM:~1,1%

:: Map month number to month name
if "%MM%"=="1"  set MONTH=January
if "%MM%"=="2"  set MONTH=February
if "%MM%"=="3"  set MONTH=March
if "%MM%"=="4"  set MONTH=April
if "%MM%"=="5"  set MONTH=May
if "%MM%"=="6"  set MONTH=June
if "%MM%"=="7"  set MONTH=July
if "%MM%"=="8"  set MONTH=August
if "%MM%"=="9"  set MONTH=September
if "%MM%"=="10" set MONTH=October
if "%MM%"=="11" set MONTH=November
if "%MM%"=="12" set MONTH=December

:: Ask for source folder name (e.g., 3 July)
set /p SOURCEFOLDER=Enter the name of the folder to copy (e.g., 3 July or July 2nd): 

:: Ask for new folder name (e.g., 5 July)
set /p NEWFOLDER=Enter the new folder name (e.g., 5 July or July 9th): 

if "%type%"=="1" (
    set "REPORT_TYPE=Daily Report"
    set "NEW_FILENAME_EXCEL=%NEWFOLDER% 2025 - ila Daily GIT Stability Report Data"
    set "NEW_FILENAME_PPTX=ila Daily Stability Report - %NEWFOLDER% 2025"
    set "OLD_FILENAME_EXCEL=%SOURCEFOLDER% 2025 - ila Daily GIT Stability Report Data"
    set "OLD_FILENAME_PPTX=ila Daily Stability Report - %SOURCEFOLDER% 2025"
) else if "%type%"=="2" (
    set "REPORT_TYPE=Weekly Report"
    set "NEW_FILENAME_EXCEL=ila Bank Daily Transactional Report - %NEWFOLDER% 2025"
    set "NEW_FILENAME_PPTX=ila Bank Daily Transactional Report - %NEWFOLDER% 2025"
    set "OLD_FILENAME_EXCEL=ila Bank Daily Transactional Report - %SOURCEFOLDER% 2025"
    set "OLD_FILENAME_PPTX=ila Bank Daily Transactional Report - %SOURCEFOLDER% 2025"
) else (
    echo Invalid input. Please enter 1 or 2.
    goto ask_type
)

:: Build paths
@REM set "BASE=C:\Users\mohamed.hasan\test\ila Digital ServiceDesk\Documents and Processes\Reports\%REPORT_TYPE%\%YEAR%\%MONTH% %YEAR%"
set "BASE=\\10.150.163.62\Data\ila Digital ServiceDesk\Documents and Processes\Reports\%REPORT_TYPE%\%YEAR%\%MONTH% %YEAR%"
set "SOURCE=%BASE%\%SOURCEFOLDER%"
set "DEST=%BASE%\%NEWFOLDER%"

echo %SOURCE%
echo %DEST%

:: Copy the folder
echo Copying folder...
xcopy "!SOURCE!" "!DEST!\" /E /I /Y >nul

if %ERRORLEVEL%==0 (
    echo Folder copied successfully to !DEST!
) else (
    echo An error occurred during copy.
    pause
    goto :eof
)

:: Rename files inside the new folder
echo Renaming files in !DEST! that contain "%SOURCEFOLDER% %YEAR%"...

@REM cd /d "!DEST!"
@REM ren "%OLD_FILENAME_EXCEL%.xlsx" "%NEW_FILENAME_EXCEL%.xlsx"
@REM ren "%OLD_FILENAME_PPTX%.pptx" "%NEW_FILENAME_PPTX%.pptx"
@REM del *.pdf


ren "%BASE%\%NEWFOLDER%\%OLD_FILENAME_EXCEL%.xlsx" "%NEW_FILENAME_EXCEL%.xlsx"
ren "%BASE%\%NEWFOLDER%\%OLD_FILENAME_PPTX%.pptx" "%NEW_FILENAME_PPTX%.pptx"
del "%BASE%\%NEWFOLDER%\*.pdf"

powershell.exe -ExecutionPolicy Bypass -File "change_links.ps1" "%BASE%\%NEWFOLDER%\%NEW_FILENAME_PPTX%.pptx" "%BASE%\%NEWFOLDER%\%NEW_FILENAME_EXCEL%.xlsx"


echo Complete
pause
goto Start