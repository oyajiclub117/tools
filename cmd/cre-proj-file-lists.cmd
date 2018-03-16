@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

pushd "C:\Users\winridge\Documents\workspace\svn-work\ドキュメント\再開発\phase_201411\110_お客様レビュー成果物"

set vList[1]="リスト・オーダー共通"
set vList[2]="勧誘DM管理システム"
set vList[3]="16_201612納品物"
set vList[4]="請求入金管理システム"
set vList[5]="変更データ管理"
set vList[6]="サブシステム"
set vList[7]="基幹システム共通\08_システムフロー"

set vId=1
set vDestFile=%USERPROFILE%\Documents\mydata\csv\proj_doc_files.csv

echo id,file,path,attr,date,size >%vDestFile%

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
	for /r %1 %%x in ("*.xlsx") do (
		@echo !vId!,"%%~nxx","%%~dpx","%%~ax","%%~tx","%%~zx"
		set /a vId+=1
	)
	goto :eof
