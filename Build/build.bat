if not exist "..\Distribution\net462\" mkdir "..\Distribution\net462\"
if not exist "..\Distribution\net6.0-windows\" mkdir "..\Distribution\net6.0-windows\"

copy /Y ..\Source\ExcelDna\Release\ExcelDna.xll ..\Distribution\net462\
copy /Y ..\Source\ExcelDna\x64\Release\ExcelDna64.xll ..\Distribution\net462\
copy /Y ..\Source\ExcelDna.Host\bin\Release\x86\ExcelDna.Host.x86.xll ..\Distribution\net6.0-windows\ExcelDna.xll
copy /Y ..\Source\ExcelDna.Host\bin\Release\x64\ExcelDna.Host.x64.xll ..\Distribution\net6.0-windows\ExcelDna64.xll

copy /Y ..\Source\ExcelDna.Integration\bin\Release\net462\ExcelDna.Integration.dll ..\Distribution\net462\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net462\ExcelDna.Integration.xml ..\Distribution\net462\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net462\ExcelDna.Integration.pdb ..\Distribution\net462\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.dll ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.xml ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.pdb ..\Distribution\net6.0-windows\

copy /Y ..\Source\ExcelDnaPack\bin\Release\net462\ExcelDnaPack.exe ..\Distribution\net462\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net462\ExcelDnaPack.exe.config ..\Distribution\net462\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net6.0-windows\ExcelDnaPack.exe ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net6.0-windows\ExcelDnaPack.dll ..\Distribution\net6.0-windows\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net6.0-windows\ExcelDnaPack.runtimeconfig.json ..\Distribution\net6.0-windows\

if not exist "..\Package\ExcelDna.AddIn\tools\net462\" mkdir "..\Package\ExcelDna.AddIn\tools\net462\"
if not exist "..\Package\ExcelDna.AddIn\tools\net6.0-windows\" mkdir "..\Package\ExcelDna.AddIn\tools\net6.0-windows\"
if not exist "..\Package\ExcelDna.AddIn.NativeAOT\tools\" mkdir "..\Package\ExcelDna.AddIn.NativeAOT\tools\"
if not exist "..\Package\ExcelDna.AddIn.NativeAOT\lib\net6.0-windows\" mkdir "..\Package\ExcelDna.AddIn.NativeAOT\lib\net6.0-windows\"
if not exist "..\Package\ExcelDna.AddIn.NativeAOT\analyzers\dotnet\cs" mkdir "..\Package\ExcelDna.AddIn.NativeAOT\analyzers\dotnet\cs"

copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net462\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net462\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\Newtonsoft.Json.dll ..\Package\ExcelDna.AddIn\tools\net462\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\Microsoft.Extensions.DependencyModel.dll ..\Package\ExcelDna.AddIn\tools\net462\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\System.Reflection.Metadata.dll ..\Package\ExcelDna.AddIn\tools\net462\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net462\System.Collections.Immutable.dll ..\Package\ExcelDna.AddIn\tools\net462\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net6.0-windows\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net6.0-windows\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net6.0-windows\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net6.0-windows\

copy /Y "..\Source\ExcelDna.Integration\bin\Release\net6.0-windows\ExcelDna.Integration.dll" "..\Package\ExcelDna.AddIn.NativeAOT\lib\net6.0-windows\"
copy /Y "..\Source\ExcelDna.ManagedHost\bin\Release\net6.0-windows\ExcelDna.ManagedHost.dll" "..\Package\ExcelDna.AddIn.NativeAOT\lib\net6.0-windows\"
copy /Y "..\Source\ExcelDna.Loader\bin\Release\net6.0-windows\ExcelDna.Loader.dll" "..\Package\ExcelDna.AddIn.NativeAOT\lib\net6.0-windows\"
copy /Y "..\Source\ExcelDna.SourceGenerator.NativeAOT\bin\Release\netstandard2.0\ExcelDna.SourceGenerator.NativeAOT.dll" "..\Package\ExcelDna.AddIn.NativeAOT\analyzers\dotnet\cs\"
copy /Y "..\Source\ExcelDna.Host.NativeAOT\bin\Release\x64\ExcelDna.Host.NativeAOT.x64.xll" "..\Package\ExcelDna.AddIn.NativeAOT\tools\ExcelDnaNativeAOT64.xll"

copy /Y ..\Distribution\net462\ ..\Distribution\

