@echo off
if "%~1" == "" (
	echo "Usage:%~nx0 [src-dir]
	pause
	exit /b 1
)

setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

set vSrcDir=%~1
set vDestDir=%userprofile%\desktop\
set vDestFile=file_paths.csv
set vId=1

pushd %vSrcDir%
(
	echo id,title,path
	for %%i in (QA•[_*.xlsx) do (
		@echo !vId!,"%%~ni","%%~dpi"
		set /a vId+=1
	)
)> %vDestDir%%vDestFile%
popd

endlocal

pause
