@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

pushd "C:\Users\winridge\Documents\workspace\svn-work\�h�L�������g\�ĊJ��\phase_201411\110_���q�l���r���[���ʕ�"

set vList[1]="���X�g�E�I�[�_�[����"
set vList[2]="���UDM�Ǘ��V�X�e��"
set vList[3]="16_201612�[�i��"
set vList[4]="���������Ǘ��V�X�e��"
set vList[5]="�ύX�f�[�^�Ǘ�"
set vList[6]="�T�u�V�X�e��"
set vList[7]="��V�X�e������\08_�V�X�e���t���["

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
