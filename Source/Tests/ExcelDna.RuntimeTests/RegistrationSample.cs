using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class RegistrationSample
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void SayHello()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaSayHello(\"world\")";
                Assert.Equal("Hello world!", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaSayHello(\"Bang!\")";
                Assert.True(functionRange.Value.ToString().StartsWith("!!! ERROR: System.ArgumentException: Bad name!"));
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void FunctionExecutionHandler()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            {
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                functionRange.Formula = "=dnaSayHello(\"FunctionExecutionHandler\")";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

                Assert.True(functionRange.Value.ToString().Contains("FunctionLoggingHandler dnaSayHello - OnEntry - Args: FunctionExecutionHandler"));
            }
            {
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                functionRange.Formula = "=dnaSayHelloTiming()";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

                Assert.True(functionRange.Value.ToString().Contains("TimingFunctionExecutionHandler dnaSayHelloTiming"));
            }
            {
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";

                functionRange.Formula = "=dnaSayHelloCache(\"123\")";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                Assert.True(functionRange.Value.ToString().Contains("CacheFunctionExecutionHandler dnaSayHelloCache result not in cache"));

                functionRange.Formula = "=dnaSayHelloCache(\"123\")";
                functionRange.Formula = "=MyRegistrationSampleFunctionExecutionLog()";
                Assert.True(functionRange.Value.ToString().Contains("CacheFunctionExecutionHandler dnaSayHelloCache result in cache"));
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void InstanceMemberRegistration()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=GetContent(\"world\")";
                Assert.Equal("Content is world", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void ParameterConversion()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaConversionTest(B2)";
                Assert.True(functionRange.Value.ToString().StartsWith("Reference: "));
                Assert.True(functionRange.Value.ToString().EndsWith("!$B$2"));
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Value = "Hello There!";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=dnaConversionToString(D1)";
                Assert.Equal("Hello There!", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D3"];
                functionRange3.Formula = "=dnaDirectString(D1)";
                Assert.Equal("Hello There!", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Value = "3.5";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=dnaConversionToDouble(C1)";
                Assert.Equal("3.5", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C3"];
                functionRange3.Formula = "=dnaDirectDouble(C1)";
                Assert.Equal("3.5", functionRange3.Value.ToString());

                Range functionRange4 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C4"];
                functionRange4.Formula = "=dnaDirectDouble()";
                Assert.Equal("0", functionRange4.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Value = "3.5";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=dnaConversionToInt32(E1)";
                Assert.Equal("4", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E3"];
                functionRange3.Formula = "=dnaDirectInt32(E1)";
                Assert.Equal("4", functionRange3.Value.ToString());

                Range functionRange4 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E4"];
                functionRange4.Formula = "=dnaDirectInt32()";
                Assert.Equal("0", functionRange4.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E11"];
                functionRange1.Value = "1";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E12"];
                functionRange2.Formula = "=dnaConversionToInt32(E11)";
                Assert.Equal("1", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E13"];
                functionRange3.Formula = "=dnaDirectInt32(E11)";
                Assert.Equal("1", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Value = "3.5";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaConversionToInt64(F1)";
                Assert.Equal("4", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F3"];
                functionRange3.Formula = "=dnaDirectInt64(F1)";
                Assert.Equal("4", functionRange3.Value.ToString());

                Range functionRange4 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F4"];
                functionRange4.Formula = "=dnaDirectInt64()";
                Assert.Equal("0", functionRange4.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F11"];
                functionRange1.Value = "1";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F12"];
                functionRange2.Formula = "=dnaConversionToInt64(F11)";
                Assert.Equal("1", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F13"];
                functionRange3.Formula = "=dnaDirectInt64(F11)";
                Assert.Equal("1", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G1"];
                functionRange1.Value = "12/18/1979  1:12:00 PM";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G2"];
                functionRange2.Formula = "=dnaConversionToDateTime(G1)";
                Assert.Equal("29207.55", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G3"];
                functionRange3.Formula = "=dnaDirectDateTime(G1)";
                Assert.Equal("29207.55", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H1"];
                functionRange1.Value = "TRUE";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H2"];
                functionRange2.Formula = "=dnaConversionToBoolean(H1)";
                Assert.Equal("True", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H3"];
                functionRange3.Formula = "=dnaDirectBoolean(H1)";
                Assert.Equal("True", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H11"];
                functionRange1.Value = "FALSE";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H12"];
                functionRange2.Formula = "=dnaConversionToBoolean(H11)";
                Assert.Equal("False", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H13"];
                functionRange3.Formula = "=dnaDirectBoolean(H11)";
                Assert.Equal("False", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["K1"];
                functionRange1.Value = "abc";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["K2"];
                functionRange2.Value = "de";

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["K3"];
                functionRange3.Formula = "=dnaJoinStrings(\";\", K1:K2)";
                Assert.Equal("abc;de", functionRange3.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["M1"];
                functionRange.Formula = "=GetErrorNA(\"abc\")";
                Assert.Equal("abc", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void ParameterConversionNullableOptional()
        {
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=dnaParameterConvertTest()";
                Assert.Equal("NULL!!!", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=dnaParameterConvertTest(2.5)";
                Assert.Equal("2.5", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=dnaDoubleNullableOptional()";
                Assert.Equal("NaN", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=dnaDoubleNullableOptional(2.5)";
                Assert.Equal("2.5", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=dnaParameterConvertOptionalTest()";
                Assert.Equal("VALUE: 42.0", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=dnaParameterConvertOptionalTest(2.5)";
                Assert.Equal("VALUE: 2.5", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=dnaMultipleOptional()";
                Assert.Equal("VALUES: 3.1415927 & @42@", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=dnaMultipleOptional(2.5)";
                Assert.Equal("VALUES: 2.5000000 & @42@", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E3"];
                functionRange3.Formula = "=dnaMultipleOptional(2.5, \"abc\")";
                Assert.Equal("VALUES: 2.5000000 & abc", functionRange3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=dnaOptionalInt()";
                Assert.Equal("VALUE: 42.0", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaOptionalInt(2)";
                Assert.Equal("VALUE: 2.0", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G1"];
                functionRange1.Formula = "=dnaOptionalString()";
                Assert.Equal("Hello World!", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G2"];
                functionRange2.Formula = "=dnaOptionalString(\"abc\")";
                Assert.Equal("abc", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H1"];
                functionRange1.Formula = "=dnaNullableOptionalDateTime()";
                Assert.Equal("NULL", functionRange1.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void ParameterConversionNullable()
        {
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=dnaNullableDouble()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=dnaNullableDouble(12.3)";
                Assert.Equal("VAL: 12.3", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=dnaNullableInt()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=dnaNullableInt(12)";
                Assert.Equal("VAL: 12", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=dnaNullableLong()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=dnaNullableLong(12345)";
                Assert.Equal("VAL: 12345", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=dnaNullableDateTime()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=dnaNullableDateTime(12.3)";
                Assert.Equal("VAL: 1/11/1900 7:12:00 AM", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=dnaNullableBoolean()";
                Assert.Equal("NULL", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaNullableBoolean(2)";
                Assert.Equal("VAL: True", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void ParameterConversionEnum()
        {
            {
                Range functionRangeB1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRangeB1.Formula = "=NA()";

                Range functionRangeC1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRangeC1.Formula = "=NA()";

                Range functionRangeB2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRangeB2.Formula = "negative";

                Range functionRangeC2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRangeC2.Formula = "imaginary";

                Range functionRange11 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B11:C11"];
                functionRange11.FormulaArray = "=dnaNullableEnum()";

                Range functionRangeB11 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B11"];
                Assert.Equal("0", functionRangeB11.Value.ToString());

                Range functionRangeC11 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C11"];
                Assert.Equal("0", functionRangeC11.Value.ToString());

                Range functionRange12 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B12:C12"];
                functionRange12.FormulaArray = "=dnaNullableEnum(B5, C5)";

                Range functionRangeB12 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B12"];
                Assert.Equal("0", functionRangeB12.Value.ToString());

                Range functionRangeC12 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C12"];
                Assert.Equal("0", functionRangeC12.Value.ToString());

                Range functionRange13 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B13:C13"];
                functionRange13.FormulaArray = "=dnaNullableEnum(B1, C1)";

                Range functionRangeB13 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B13"];
                Assert.Equal("0", functionRangeB13.Value.ToString());

                Range functionRangeC13 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C13"];
                Assert.Equal("0", functionRangeC13.Value.ToString());

                Range functionRange14 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B14:C14"];
                functionRange14.FormulaArray = "=dnaNullableEnum(B2, C2)";

                Range functionRangeB14 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B14"];
                Assert.Equal("-1", functionRangeB14.Value.ToString());

                Range functionRangeC14 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C14"];
                Assert.Equal("1", functionRangeC14.Value.ToString());

                Range functionRange24 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B24:C24"];
                functionRange24.FormulaArray = "=dnaEnumParameters(B2, C2)";

                Range functionRangeB24 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B24"];
                Assert.Equal("-1", functionRangeB24.Value.ToString());

                Range functionRangeC24 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C24"];
                Assert.Equal("1", functionRangeC24.Value.ToString());
            }
            {
                Range functionRangeD1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRangeD1.Value = "imaginary";

                Range functionRangeD2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRangeD2.Value = "real";

                Range functionRangeD3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D3"];
                functionRangeD3.Value = "imaginary";

                Range functionRangeE = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1:E3"];
                functionRangeE.FormulaArray = "=dnaEnumsEnumerated(D1:D3)";

                Range functionRangeE1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                Assert.Equal("Negative", functionRangeE1.Value.ToString());

                Range functionRangeE2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                Assert.Equal("Positive", functionRangeE2.Value.ToString());

                Range functionRangeE3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E3"];
                Assert.Equal("Negative", functionRangeE3.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=dnaReturnEnum1(\"Negative\")";
                Assert.Equal("Negative", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=dnaReturnEnum1(\"Positive\")";
                Assert.Equal("Positive", functionRange2.Value.ToString());
            }
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G1"];
                functionRange1.Formula = "=dnaReturnEnum2(\"Real\")";
                Assert.Equal("Real", functionRange1.Value.ToString());

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G2"];
                functionRange2.Formula = "=dnaReturnEnum2(\"Imaginary\")";
                Assert.Equal("Imaginary", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void ParameterConversionComplex()
        {
            {
                Range functionRangeB1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRangeB1.Formula = "=NA()";

                Range functionRangeC1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRangeC1.Formula = "=NA()";

                Range functionRangeB2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRangeB2.Formula = "1";

                Range functionRangeC2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRangeC2.Formula = "2";

                Range functionRange11 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B11:C11"];
                functionRange11.FormulaArray = "=dnaComplex(B2:C2)";

                Range functionRangeB11 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B11"];
                Assert.Equal("1", functionRangeB11.Value.ToString());

                Range functionRangeC11 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C11"];
                Assert.Equal("2", functionRangeC11.Value.ToString());

                Range functionRange12 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B12:C12"];
                functionRange12.FormulaArray = "=dnaNullableComplex()";

                Range functionRangeB12 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B12"];
                Assert.Equal("111", functionRangeB12.Value.ToString());

                Range functionRangeC12 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C12"];
                Assert.Equal("222", functionRangeC12.Value.ToString());

                Range functionRange13 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B13:C13"];
                functionRange13.FormulaArray = "=dnaNullableComplex(B5:C5)";

                Range functionRangeB13 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B13"];
                Assert.Equal("0", functionRangeB13.Value.ToString());

                Range functionRangeC13 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C13"];
                Assert.Equal("0", functionRangeC13.Value.ToString());

                Range functionRange14 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B14:C14"];
                functionRange14.FormulaArray = "=dnaNullableComplex(B1)";

                Range functionRangeB14 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B14"];
                Assert.Equal("111", functionRangeB14.Value.ToString());

                Range functionRangeC14 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C14"];
                Assert.Equal("222", functionRangeC14.Value.ToString());

                Range functionRange15 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B15:C15"];
                functionRange15.FormulaArray = "=ERROR.TYPE(dnaNullableComplex(B1:C1))";

                Range functionRangeB15 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B15"];
                Assert.Equal("6", functionRangeB15.Value.ToString());

                Range functionRangeC15 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C15"];
                Assert.Equal("6", functionRangeC15.Value.ToString());

                Range functionRange16 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B16:C16"];
                functionRange16.FormulaArray = "=dnaNullableComplex(B2:C2)";

                Range functionRangeB16 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B16"];
                Assert.Equal("1", functionRangeB16.Value.ToString());

                Range functionRangeC16 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C16"];
                Assert.Equal("2", functionRangeC16.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void ParameterConversionType()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaTestFunction1(\"abc\")";
                Assert.Equal("The Test (1) value is abc", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaTestFunction2(\"abc\")";
                Assert.Equal("The Test (2) value is abc", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange.Formula = "=dnaTestFunction2Ret1(\"abc\")";
                Assert.Equal("From Type 1 with The Test (2) value is abc", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void Params()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaParamsFunc2(\"a\",,\"c\",\"d\",,\"f\")";
                Assert.Equal("a,,c, [3: d,ExcelDna.Integration.ExcelMissing,f]", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaParamsFunc(\"a\",\"b\",\"c\",\"d\")";
                Assert.Equal("a,b, : 2", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange.Formula = "=dnaJoinStringParams(\"|\",\"a\",\"b\",\"c\")";
                Assert.Equal("a|b|c", functionRange.Value.ToString());
            }
        }
    }
}
