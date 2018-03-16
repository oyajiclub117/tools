setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

set vHomeDir="C:\Users\winridge\Documents\apps\Firebird-3.0"

set path=%vHomeDir%;%path%


@echo ホームディレクトリへ移動します。[%vHomeDir%]
pushd %vHomeDir%

%comspec%

popd
@echo ホームディレクトリから戻りました。

endlocal
