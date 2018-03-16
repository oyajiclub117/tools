set vCmdHome=C:\Users\winridge\Documents\tools\vbs\upd-sheet_ver
set vCmd=%vCmdHome%\main.wsf
set vSrcDir=%~dpnx1

for /r "%vSrcDir%" %%x in ("*.xlsx") do (
	@cscript %vCmd% "%%~dpnxx"
)

pause
