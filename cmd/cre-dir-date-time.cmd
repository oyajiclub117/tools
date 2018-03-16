@echo off
:proc-Begin
	if "%~1" == "" (
		echo Usage:%~nx0 [dest-dir]
		exit /b 1
	)

	set vDestDir=%~dpnx1\

:proc-Main
	set vDate=%date: =0%
	set vTime=%time: =0%

	set vDate=%vDate:~0,4%%vDate:~5,2%%vDate:~8,2%
	set vTime=%vtime:~0,2%%vtime:~3,2%%vtime:~6,2%

	mkdir %vDestDir%%vDate%-%vTime%

	if %errorlevel% neq 0 (
		echo mkdir command failer!! [dir=%vDestdir% rc=%errorlevel%]
		exit /b 2
	)

:proc-End
	pause
