@echo off
REM --- Config          ------------------
SET ProjectName="CreateZip"
SET TargetSourceFiles=%ProjectName%"\*.cs"
REM --- Build            ------------------
pushd %~dp0
FOR /F "TOKENS=1,2,*" %%I IN ('REG QUERY "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\.NETFramework" /v "InstallRoot"') DO IF "%%I"=="InstallRoot" SET FrameworkPath=%%K
SET PATH="%PATH%;%FrameworkPath%v4.0.30319\;%FrameworkPath%v3.5;%FrameworkPath%v3.0;"
SET PATH=%PATH%;%systemroot%\Microsoft.NET\Framework\version
msbuild.exe CreateZip\CreateZip.csproj /p:OutputPath=".."
if %ERRORLEVEL% NEQ 0 goto FAILURE
REM --- Execute Cs    ------------------
%ProjectName%
REM --- Delete Trash ------------------
rd /s /q bin
del %ProjectName%.exe
del %ProjectName%.exe.config
del %ProjectName%.pdb
REM -----------------------------------
goto END

:FAILURE
REM cls
%0

:END
