@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg
set ExcelDnaVersion=%1
set ExcelDnaNativeAOTVersion=%2

if exist "%outputPath%\*.nupkg" del "%outputPath%\*.nupkg"

if not exist "%outputPath%" mkdir "%outputPath%"

echo on

nuget.exe pack "%basePath%\ExcelDna.Templates\ExcelDna.Templates.nuspec" -BasePath "%basePath%\ExcelDna.Templates" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.AddIn\ExcelDna.AddIn.nuspec" -BasePath "%basePath%\ExcelDna.AddIn" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.AddIn.NativeAOT\ExcelDna.AddIn.NativeAOT.nuspec" -BasePath "%basePath%\ExcelDna.AddIn.NativeAOT" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaNativeAOTVersion="%ExcelDnaNativeAOTVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.Integration\ExcelDna.Integration.nuspec" -BasePath "%basePath%\ExcelDna.Integration" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%" -Symbols -SymbolPackageFormat snupkg
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.XmlSchemas\ExcelDna.XmlSchemas.nuspec" -BasePath "%basePath%\ExcelDna.XmlSchemas" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\Excel-DNA\Excel-DNA.nuspec" -BasePath "%basePath%\Excel-DNA" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\Excel-DNA.Lib\Excel-DNA.Lib.nuspec" -BasePath "%basePath%\Excel-DNA.Lib" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

:end
