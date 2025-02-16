setlocal

set PackageVersion=1.9.0-alpha3
set PackageNativeAOTVersion=0.1.0
set PackageReferenceVersion=1.9.0-alpha3
set DllVersion=1.9.0.3

set MSBuildPath="c:\Program Files\Microsoft Visual Studio\2022\Professional\Msbuild\Current\Bin\amd64\MSBuild.exe"

set rootPath=%~dp0..\..

set propsfile=%rootPath%\Directory.Build.props
copy /Y Directory.Build.props %propsfile%
PowerShell "(Get-Content %propsfile%) -replace '_VERSION_', '%DllVersion%' | Set-Content %propsfile%"
@if errorlevel 1 goto end

set targetsfile=%rootPath%\Directory.Build.targets.local
copy /Y PushAll.template.cmd PushAll.cmd
PowerShell "(Get-Content PushAll.cmd) -replace '_VERSION_', '%PackageVersion%' | Set-Content PushAll.cmd
@if errorlevel 1 goto end

set targetsfile=%rootPath%\Directory.Build.targets.local
copy /Y Directory.Build.targets %targetsfile%
PowerShell "(Get-Content %targetsfile%) -replace '_VERSION_', '%PackageReferenceVersion%' | Set-Content %targetsfile%"
@if errorlevel 1 goto end

cd %rootPath%\ExcelDna\Build
call BuildPackages.bat %PackageVersion% %PackageNativeAOTVersion% %DllVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\Registration\Build
copy /Y %targetsfile% %rootPath%\Registration\Source\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\IntelliSense\Build
copy /Y %targetsfile% %rootPath%\IntelliSense\Source\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\ExcelDnaDoc\Build
copy /Y %targetsfile% %rootPath%\ExcelDnaDoc\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\DeveloperTools\ExcelDna.Testing\Build
copy /Y %targetsfile% %rootPath%\DeveloperTools\ExcelDna.Testing\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

:end
