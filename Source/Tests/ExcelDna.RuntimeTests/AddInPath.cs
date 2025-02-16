namespace ExcelDna.RuntimeTests
{
    internal class AddInPath
    {
#if DEBUG
        public const string RuntimeTestsAOT = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\Debug\net8.0-windows\win-x64\publish\ExcelDna.AddIn.RuntimeTestsAOT";
        public const string RuntimeTests = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Debug\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn";
        public const string RegistrationSample = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn";
        public const string RegistrationSampleFS = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleFS\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleFS-AddIn";
        public const string RegistrationSampleVB = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleVB\bin\Debug\net6.0-windows\ExcelDna.AddIn.RegistrationSampleVB-AddIn";
#else
        public const string RuntimeTestsAOT = @"..\..\..\..\ExcelDna.AddIn.RuntimeTestsAOT\bin\Release\net8.0-windows\win-x64\publish\ExcelDna.AddIn.RuntimeTestsAOT";
        public const string RuntimeTests = @"..\..\..\..\ExcelDna.AddIn.RuntimeTests\bin\Release\net6.0-windows\ExcelDna.AddIn.RuntimeTests-AddIn";
        public const string RegistrationSample = @"..\..\..\..\ExcelDna.AddIn.RegistrationSample\bin\Release\net6.0-windows\ExcelDna.AddIn.RegistrationSample-AddIn";
        public const string RegistrationSampleFS = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleFS\bin\Release\net6.0-windows\ExcelDna.AddIn.RegistrationSampleFS-AddIn";
        public const string RegistrationSampleVB = @"..\..\..\..\ExcelDna.AddIn.RegistrationSampleVB\bin\Release\net6.0-windows\ExcelDna.AddIn.RegistrationSampleVB-AddIn";
#endif
    }
}
