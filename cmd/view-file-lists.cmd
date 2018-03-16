@echo off
if "%~1" == "" (
	echo Usage:%~nx0 [src-dir]
	exit /b 1
)

dir /b %~dp1

pause
