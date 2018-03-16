@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

pushd "C:\Users\winridge\Documents\workspace\svn-work\ドキュメント\再開発\phase_201411\110_お客様レビュー成果物"

set vId=1

set vDestFile=%USERPROFILE%\desktop\bus_system_flows.csv

set vList[1]="リスト・オーダー共通\08_システムフロー\v4.1"
set vList[2]="勧誘DM管理システム\08_システムフロー\v3.1"
set vList[3]="16_201612納品物\08_システムフロー\v4.0"
set vList[4]="請求入金管理システム\08_システムフロー"
set vList[5]="変更データ管理\08_システムフロー\v4.0"
set vList[6]="サブシステム\08_システムフロー"
set vList[7]="基幹システム共通\08_システムフロー\v4.0"

echo id,bus_id,bus_name >%vDestFile%

(
	call :proc-file-list %vList[1]%
	call :proc-file-list %vList[2]%
	call :proc-file-list %vList[3]%
	call :proc-file-list %vList[4]%
	call :proc-file-list %vList[5]%
	call :proc-file-list %vList[6]%
	call :proc-file-list %vList[7]%
) >>%vDestFile%

pause

goto :eof

popd

:proc-file-list
	for /r %1 %%x in ("システム*.*") do (
		@for /f "tokens=1,2,3 delims=_.（）" %%a in ("%%~nxx") do (
			@echo !vId!,"%%b","%%c"
			set /a vId+=1
		)
	)
	goto :eof
