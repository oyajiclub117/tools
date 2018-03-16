@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
set vId=1

for /r . %%i in (*.*) do (
	echo !vId!,"%%~nxi",""
	set /a vId+=1
)
