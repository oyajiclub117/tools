@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

pushd %1

set vId=0
for /f "usebackq tokens=1" %%i in (`dir /b^|findstr /v "ƒoƒbƒ`"`) do (
	@for /d %%x in ("%%i\*.*") do (
		@for /f "tokens=1,2,3 delims=() " %%a in ("%%~nxx") do (
			@set /a vId+=1
			@echo !vId!,"%%a","%%b","%%i"
		)
	)
)

popd
