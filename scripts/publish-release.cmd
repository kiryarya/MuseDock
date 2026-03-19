@echo off
setlocal

set ROOT=%~dp0..
set DOTNET_CLI_HOME=%ROOT%\.dotnet-home
set NUGET_PACKAGES=%ROOT%\.nuget
set PROJECT=%ROOT%\src\MuseDock.Desktop\MuseDock.Desktop.csproj
set PROJECT_OBJ=%ROOT%\src\MuseDock.Desktop\obj
set DIST_ROOT=%ROOT%\dist
set DIST_OUT=%DIST_ROOT%\MuseDock-win-x64

if not exist "%DOTNET_CLI_HOME%" mkdir "%DOTNET_CLI_HOME%"
if not exist "%NUGET_PACKAGES%" mkdir "%NUGET_PACKAGES%"
if not exist "%DIST_ROOT%" mkdir "%DIST_ROOT%"

if exist "%DIST_OUT%" rmdir /s /q "%DIST_OUT%"
if exist "%DIST_ROOT%\portable" rmdir /s /q "%DIST_ROOT%\portable"
if exist "%DIST_ROOT%\self-contained" rmdir /s /q "%DIST_ROOT%\self-contained"
if exist "%DIST_ROOT%\self-contained-toolbar-fix" rmdir /s /q "%DIST_ROOT%\self-contained-toolbar-fix"
if exist "%DIST_ROOT%\self-contained-contrast-fix" rmdir /s /q "%DIST_ROOT%\self-contained-contrast-fix"
if exist "%DIST_ROOT%\self-contained-contrast-fix-2" rmdir /s /q "%DIST_ROOT%\self-contained-contrast-fix-2"
if exist "%DIST_ROOT%\win-x64" rmdir /s /q "%DIST_ROOT%\win-x64"
if exist "%PROJECT_OBJ%" rmdir /s /q "%PROJECT_OBJ%"

dotnet restore "%PROJECT%" -r win-x64
if errorlevel 1 exit /b 1

dotnet publish "%PROJECT%" -c Release -r win-x64 --self-contained true --no-restore -o "%DIST_OUT%"
if errorlevel 1 exit /b 1

echo.
echo Ready: %DIST_OUT%\MuseDock.Desktop.exe
endlocal
