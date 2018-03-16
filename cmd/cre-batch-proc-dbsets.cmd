@echo off

set vcmd=C:\Users\winridge\Documents\tools\vbs\cre-csv-dbset-lists\main.wsf

set vDate=%date: =0%
set vTime=%time: =0%
set vWorkDir=c:\temp\work
set vLog=run_%vDate:~0,4%%vDate:~5,2%%vDate:~8,2%_%vTime:~0,2%%vTime:~3,2%%vTime:~6,2%%vTime:~9,2%.log
set vTargetDir=%vWorkDir%\%~nx1

mkdir %vTargetDir%

xcopy /s "%~1" %vTargetDir%

@echo date_time,file_name >%vLog%
for /r %vTargetDir% %%x in ("ƒoƒbƒ`ˆ—*.xlsx") do (
	@echo *** start %date: =0% %time: =0%,%%~nxx,
	@cscript //Nologo %vcmd% "%%~dpnxx"
	if errorlevel 1 (
		@echo "%date: =0% %time: =0%","%%~nxx" >>%vLog%
	)
	@echo *** end   %date: =0% %time: =0%,%%~nxx,%errorlevel%
)
rmdir /s /q %vTargetDir%
pause
