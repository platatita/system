@echo off

set directoryPath=%1
set expectedSize=%2


@echo ----------------------------------------------------------------
@echo Provided arguments:
@echo Arg 1 (directoryPath)		: %1
@echo Arg 2 (expectedSize) [Bytes]	: %2

:: ============================== Validation ==============================
IF /I '%directoryPath%'=='' goto DirectoryNotExist
IF /I '%expectedSize%'=='' set expectedSize=200000

@echo directoryPath: %directoryPath%
@echo expectedSize: %expectedSize%


:: ============================== Processing ==============================
FOR /F %%F in ('dir /S /B %directoryPath%') DO call resize_image.bat %%F %expectedSize%
goto End


:: ============================== End Message ==============================
:DirectoryNotExist
@echo Not provided directory path to get files to resize.

:End
@echo End.