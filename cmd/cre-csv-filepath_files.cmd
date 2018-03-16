@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

@rem // import profile 
call %~dp0profile.cmd

if "%~1" == "" (
	echo Usage:%~nx0 [src-dir]
	exit /b 1
)

if "%~2" == "" (
	echo Usage:%~nx0 [src-dir]
	exit /b 2
)

if not exist "%~1" (
	echo directory not exist!! [src-dir:%~1]
	exit /b 2
)

set vSrcDir=%~dp1

call %~dp0
set vDestDir=%userprofile%\Desktop\
set vDestFile=%vDestDir%file_path_lists_%vDate%_%vTime%.csv

set vId=1

(
	echo id,file_name,drive_name,path,file_ext,attributes_map,created,file_size
	for %%i in ("%vSrcDir%"*.*) do (
		echo !vId!,"%%~nxi","%%~di","%%~pi","%%~xi","%%~ai","%%~ti","%%~zi"
		set /a vId+=1
	)
) > %vDestFile%
endlocal

pause
