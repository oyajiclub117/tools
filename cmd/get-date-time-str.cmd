if "%~1" == "" (
	echo Usage:%~nx0 [var-date] [var-time]
	exit /b 1
)

if "%~2" == "" (
	echo Usage:%~nx0 [var-date] [var-time]
	exit /b 2
)

set vDate=%date: =0%
set vTime=%time: =0%
set vDate=%vDate:~0,4%%vDate:~5,2%%vDate:~8,2%
set vTime=%vTime:~0,2%%vTime:~3,2%%vTime:~6,2%%vTime:~9,3%

set %~1=%vDate%
set %~2=%vTime%
