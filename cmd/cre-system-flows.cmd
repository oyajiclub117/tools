@echo off
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

pushd "C:\Users\winridge\Documents\workspace\svn-work\�h�L�������g\�ĊJ��\phase_201411\110_���q�l���r���[���ʕ�"

set vId=1

set vDestFile=%USERPROFILE%\desktop\bus_system_flows.csv

set vList[1]="���X�g�E�I�[�_�[����\08_�V�X�e���t���[\v4.1"
set vList[2]="���UDM�Ǘ��V�X�e��\08_�V�X�e���t���[\v3.1"
set vList[3]="16_201612�[�i��\08_�V�X�e���t���[\v4.0"
set vList[4]="���������Ǘ��V�X�e��\08_�V�X�e���t���["
set vList[5]="�ύX�f�[�^�Ǘ�\08_�V�X�e���t���[\v4.0"
set vList[6]="�T�u�V�X�e��\08_�V�X�e���t���["
set vList[7]="��V�X�e������\08_�V�X�e���t���[\v4.0"

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
	for /r %1 %%x in ("�V�X�e��*.*") do (
		@for /f "tokens=1,2,3 delims=_.�i�j" %%a in ("%%~nxx") do (
			@echo !vId!,"%%b","%%c"
			set /a vId+=1
		)
	)
	goto :eof
