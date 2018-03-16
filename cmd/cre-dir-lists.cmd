rem --@echo off

:proc-Begin
	setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

	set vRc=0

	if "%~1" == "" (
		echo usage:%~nx0 [src-dir]
		set vRc=1
		goto :proc-End
	)

	if not exist "%~dp1" (
		echo directory not exist [src-dir=%1]
		set vRc=2
		goto :proc-End
	)

	set vId=1
	set vSrcDir="%~1\"
:proc-Main
	for /d . %%x in ("%vSrcDir%\*.*") do (
		echo !vId!,"%%~nxx",""
		set /a vId+=1
	)

:proc-End
	if %vRc% neq 0 (
		exit /b %vRc%
	)

	endlocal
