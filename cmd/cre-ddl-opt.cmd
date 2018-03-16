if "%1" == "" (
	echo Usage:%~nx [src-dir]
	exit /b 1
)

set vSrcDir=%~1
set /p vFileSet="type-in file-set[*.*]:"
for %i in (%vSrcDir%\%vFileSet%) do @echo ^<option value="%~dpnxi"^>%~ni^</option^n^>

