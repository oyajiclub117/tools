@echo off

setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

if "%~1" == "" (
	echo "Usage:%~nx0 [src-path]"
	exit /b 1
)

if not exist %~dpnx1 (
	echo not exist src-path!! [src-path:%~1]
	exit /b 2
)

set vId=0
(
	for /r "%~1" %%i in (*.*) do (
		set /a vId += 1
		echo !vId!,"%%~ni","%%~nxi"
	)
) > %~dp1file_lists.csv
endlocal
