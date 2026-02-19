@echo off
setlocal

REM Move working directory to the install folder root
cd /d "%~dp0\.."

REM Point to bundled R
set "R_HOME=%cd%\R-4.5.2"
set "PATH=%R_HOME%\bin\x64;%PATH%"

REM Run the Shiny app
"%R_HOME%\bin\x64\Rscript.exe" --vanilla -e "options(shiny.port=3838, shiny.host='127.0.0.1'); shiny::runApp('app', launch.browser=TRUE)"
