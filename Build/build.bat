copy /Y ..\Source\ExcelDna\Release\ExcelDna.xll ..\Distribution\
copy /Y ..\Source\ExcelDna\Release64\ExcelDna64.xll ..\Distribution\
copy /Y ..\Source\ExcelDna.Integration\bin\Release\net452\ExcelDna.Integration.dll ..\Distribution\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net452\ExcelDnaPack.exe ..\Distribution\
copy /Y ..\Source\ExcelDnaPack\bin\Release\net452\ExcelDnaPack.exe.config ..\Distribution\
if not exist "..\Source\ExcelDna.AddIn.Tasks\bin\Release\net472\" mkdir "..\Source\ExcelDna.AddIn.Tasks\bin\Release\net472\"
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net472\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net472\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net472\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net472\
if not exist "..\Source\ExcelDna.AddIn.Tasks\bin\Release\net5.0-windows\" mkdir "..\Source\ExcelDna.AddIn.Tasks\bin\Release\net5.0-windows\"
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net5.0-windows\ExcelDna.AddIn.Tasks.dll ..\Package\ExcelDna.AddIn\tools\net5.0-windows\
copy /Y ..\Source\ExcelDna.AddIn.Tasks\bin\Release\net5.0-windows\ExcelDna.AddIn.Tasks.pdb ..\Package\ExcelDna.AddIn\tools\net5.0-windows\
pause
