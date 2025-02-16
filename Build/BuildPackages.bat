setlocal

set PackageVersion=%1
set PackageNativeAOTVersion=%2
set DllVersion=%3
set MSBuildPath=%4

set rcfile=..\Source\versioninfo.rc2
PowerShell "(Get-Content %rcfile%) -replace '\d+,\d+,\d+,\d+', '%DllVersion%'.Replace('.',',') -replace '\d+\.\d+\.\d+\.\d+', '%DllVersion%' | Set-Content %rcfile%"
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:restore,build /p:Configuration=Release /p:ContinuousIntegrationBuild=true
@if errorlevel 1 goto end

%MSBuildPath% ..\Source\ExcelDna.sln /t:restore,build /p:Configuration=Release /p:Platform=x64 /p:ContinuousIntegrationBuild=true
@if errorlevel 1 goto end

call build.bat
@if errorlevel 1 goto end

cd ..\Package
call package.cmd %PackageVersion% %PackageNativeAOTVersion%
@if errorlevel 1 goto end

set pushfile=push.cmd
set pushcurrentfile=pushCurrent.cmd
PowerShell "(Get-Content %pushfile%) -replace '_VERSION_', '%PackageVersion%' | Set-Content %pushcurrentfile%"
@if errorlevel 1 goto end

:end
