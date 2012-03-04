@echo off

set defaultSize=200000
set defaultResizedFileName=
set fileNameToResize=%1
set expectedSize=%2
set resizedFileName=%3
set resizePercentage=95%%%%

@echo ----------------------------------------------------------------
@echo Provided arguments:
@echo Arg 1 (fileNameToResize)		: %1
@echo Arg 2 (expectedSize) [Bytes]	: %2
@echo Arg 3 (resizedFileName)		: %3

:: ============================== Validation ==============================
IF /I "%fileNameToResize%"=="" goto LacksArgs
IF NOT EXIST %fileNameToResize% goto FileNotExist
FOR %%I in (%fileNameToResize%) DO set defaultResizedFileName=%%~dpIresized_%%~nI%%~xI
IF /I "%expectedSize%"=="" set expectedSize=%defaultSize%
IF /I "%resizedFileName%"=="" set resizedFileName=%defaultResizedFileName%

@echo fileNameToResize: %fileNameToResize%
@echo expectedSize: %expectedSize% [bytes]
@echo resizedFileName: %resizedFileName%


:: ============================== Initialization ==============================
COPY /Y %fileNameToResize% %resizedFileName%
@echo coppied '%fileNameToResize%' to '%resizedFileName%'


:: ============================== Processing ==============================
:ReadSize
FOR /D %%T in (%resizedFileName%) DO set size=%%~zT

@echo File size: %size%

set x="90%%"

IF %size% GTR %expectedSize% (
call convert %resizedFileName% -resize %resizePercentage% %resizedFileName%
goto ReadSize
) ELSE (
goto End
)


:: ============================== End Message ==============================
:LacksArgs
@echo Not provided any parameters.
goto End

:FileNotExist
@echo Provided file: '%fileNameToResize%' does not exist.

:End
@echo End processing.