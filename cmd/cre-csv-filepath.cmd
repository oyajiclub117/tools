@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

set vDestDir=%userprofile%\Desktop\
set vDestFile=%vDestDir%%~2

if "%~1" == "" (
	echo Usage:%~nx0 [src-dir] [out-file]
	exit /b 1
)

if "%~2" == "" (
	echo Usage:%~nx0 [src-dir] [out-file]
	exit /b 2
)

if not exist "%~1" (
	echo directory not exist!! [src-dir:%~1]
	exit /b 2
)

set vSrcDir=%~1

set vId=1

(
	echo id,file_name,drive_name,path,file_ext,attributes_map,created,file_size
	for /r "%vSrcDir%" %%i in (*.*) do (
		echo !vId!,"%%~nxi","%%~di","%%~pi","%%~xi","%%~ai","%%~ti","%%~zi"
		set /a vId+=1
	)
) > %vDestFile%
endlocal

pause
