@echo off
SET ProjectName="CreateZip"
SET TargetSourceFiles="CreateZip\*.cs"

pushd %~dp0
FOR /F "TOKENS=1,2,*" %%I IN ('REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework" /v "InstallRoot"') DO IF "%%I"=="InstallRoot" SET FrameworkPath=%%K
SET PATH="%PATH%;%FrameworkPath%v4.0.30319\;%FrameworkPath%v3.5;%FrameworkPath%v3.0;"
csc /nologo /out:%ProjectName%.exe %TargetSourceFiles% /r:"System.IO.Compression.dll"
if %ERRORLEVEL% NEQ 0 goto FAILURE
%ProjectName%
del %ProjectName%.exe
exit

:FAILURE
cls
%0
