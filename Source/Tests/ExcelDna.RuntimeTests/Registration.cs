using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class Registration
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void SayHello()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=SayHello(\"world\")";
            Assert.Equal("Hello world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Double()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyDouble(3.4)";
            Assert.Equal("3.4", functionRange.Value.ToString());

            functionRange.Formula = "=MyDouble()";
            Assert.Equal("0", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void DateTime()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyDateTime(\"2024/10/20\")";
            Assert.Equal("10/20/2024 12:00:00 AM", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void NullableDouble()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyNullableDouble(1.2)";
            Assert.Equal("Nullable VAL: 1.2", functionRange.Value.ToString());

            functionRange.Formula = "=MyNullableDouble()";
            Assert.Equal("Nullable VAL: NULL", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void NullableDateTime()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyNullableDateTime(\"2024/11/21\")";
            Assert.Equal("Nullable DateTime: 11/21/2024 12:00:00 AM", functionRange.Value.ToString());

            functionRange.Formula = "=MyNullableDateTime()";
            Assert.Equal("Nullable DateTime: NULL", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void OptionalDouble()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyOptionalDouble(2.3)";
            Assert.Equal("Optional VAL: 2.3", functionRange.Value.ToString());

            functionRange.Formula = "=MyOptionalDouble()";
            Assert.Equal("Optional VAL: 1.23", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void OptionalDateTime()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyOptionalDateTime(\"2024/11/21\")";
            Assert.Equal("Optional DateTime: 11/21/2024 12:00:00 AM", functionRange.Value.ToString());

            functionRange.Formula = "=MyOptionalDateTime()";
            Assert.Equal("Optional DateTime: 1/1/0001 12:00:00 AM", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Enum()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyEnum(\"Unspecified\")";
            Assert.Equal("Enum VAL: Unspecified", functionRange.Value.ToString());

            functionRange.Formula = "=MyEnum(\"Local\")";
            Assert.Equal("Enum VAL: Local", functionRange.Value.ToString());

            functionRange.Formula = "=MyEnum(1)";
            Assert.Equal("Enum VAL: Utc", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void EnumReturn()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyEnumReturn(\"Unspecified\")";
            Assert.Equal("Unspecified", functionRange.Value.ToString());

            functionRange.Formula = "=MyEnumReturn(\"Local\")";
            Assert.Equal("Local", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void MapArray()
        {
            Range a1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1:A1"];
            a1.Value = "Utc";

            Range a2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2:A2"];
            a2.Value = "Local";

            Range a3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3:A3"];
            a3.Value = "Unspecified";

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B3"];
            functionRange.FormulaArray = "=MyMapArray(A1:A3)";

            Range b1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            Assert.Equal("Array element VAL: Utc", b1.Value.ToString());

            Range b2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2:B2"];
            Assert.Equal("Array element VAL: Local", b2.Value.ToString());

            Range b3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3:B3"];
            Assert.Equal("Array element VAL: Unspecified", b3.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void AsyncTaskInstant()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyAsyncTaskHello(\"world\", 0)";

            Assert.Equal("Hello async task world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void DefaultAsyncReturnValue()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyAsyncTaskHello(\"world\", 1000)";

            Assert.Equal(-2146826246, functionRange.Value); // #N/A
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void GettingDataAsyncReturnValue()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyAsyncGettingData(\"world\", 1000)";

            Assert.Equal(-2146826245, functionRange.Value); // #GETTING_DATA
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void TaskInstant()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyTaskHello(\"world\")";

            Assert.Equal("Hello task world", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void StringArray()
        {
            Range a1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1:A1"];
            a1.Value = "01";

            Range a2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2:A2"];
            a2.Value = "2.30";

            Range a3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3:A3"];
            a3.Value = "World";

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];
            functionRange.Formula = "=MyStringArray(A1:A3)";

            Assert.Equal("StringArray VALS: 12.3World", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void StringArray2D()
        {
            Range a1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
            a1.Value = "01";

            Range a2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
            a2.Value = "2.30";

            Range a3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A3"];
            a3.Value = "Hello";

            Range b1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
            b1.Value = "5";

            Range b2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
            b2.Value = "6.7";

            Range b3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3"];
            b3.Value = "World";

            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
            functionRange.Formula = "=MyStringArray2D(A1:B3)";

            Assert.Equal("StringArray2D VALS: 15 2.36.7 HelloWorld ", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void UserDefinedParameterConversions()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1:B1"];

            functionRange.Formula = "=MyVersion2(\"4.3.2.1\")";
            Assert.Equal("The Version value with field count 2 is 4.3", functionRange.Value.ToString());

            functionRange.Formula = "=MyTestType1(\"world\")";
            Assert.Equal("The TestType1 value is world", functionRange.Value.ToString());

            functionRange.Formula = "=MyTestType2(\"world2\")";
            Assert.Equal("The TestType2 value is From TestType1 world2", functionRange.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void FunctionExecutionHandlerExtended()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=MyVersion2(\"5.4.3.2\")";
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("MyVersion2 - OnSuccess - Result: The Version value with field count 2 is 5.4"));
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void FunctionExecutionHandlerStandard()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=SayHello(\"FunctionExecutionHandlerStandard\")";
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("SayHello - OnSuccess - Result: Hello FunctionExecutionHandlerStandard"));
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void FunctionExecutionHandlerWithAttribute()
        {
            Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

            functionRange.Formula = "=SayHelloWithLoggingID(\"FunctionExecutionHandlerWithAttribute\")";
            functionRange.Formula = "=MyFunctionExecutionLog()";
            Assert.True(functionRange.Value.ToString().Contains("ID=7 SayHelloWithLoggingID - OnSuccess - Result: Hello FunctionExecutionHandlerWithAttribute"));
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Observable()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];

                functionRange.Formula = "=MyStringObservable(\"s1\")";
                Assert.Equal("s1", functionRange.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateCalc(12, 13)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyCalcSumObservable(C1)";

                Assert.Equal("25", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyCalcObservable(14, 15)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyCalcSum(D1)";

                Assert.Equal("29", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=MyCalcExcelHandleObservable(0.5, 8)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=MyCalcExcelHandleMul(E1)";

                Assert.Equal("4", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void ObjectHandles()
        {
            string b1;
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateCalc(45, 0)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyCalcSum(B1)";

                b1 = functionRange1.Value.ToString();
                Assert.Equal("45", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateCalc(45, 0)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyCalcSum(C1)";

                Assert.Equal(b1, functionRange1.Value.ToString());
                Assert.Equal("45", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyCreateCalc(54, 0)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyCalcSum(D1)";

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal("54", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=MyCreateCalc2(45, 0)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=MyCalcSum(E1)";

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal("90", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=MyCreateCalc(1.2, 3.4)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=MyCalcSum(F1)";

                Assert.Equal("4.6", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G1"];
                functionRange1.Formula = "=MyCreateSquareIntObject(3)";
                Assert.True(functionRange1.Value.ToString().StartsWith("MyCreateSquareIntObject"));

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G2"];
                functionRange2.Formula = "=MyPrintIntObject(G1)";

                Assert.Equal("IntObject value=9", functionRange2.Value.ToString());

                Range functionRange3 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["G3"];
                functionRange3.Formula = "=MyPrintMixedIntObject(4.5, G1)";

                Assert.Equal("double value=4.5, IntObject value=9", functionRange3.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H1"];
                functionRange1.Formula = "=MyCreateCalcExcelHandle(1.2, 3.5)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["H2"];
                functionRange2.Formula = "=MyCalcExcelHandleMul(H1)";

                Assert.Equal("4.2", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B10"];
                functionRange1.Formula = "=MyCreateCalcStructExcelHandle(1.5, 0.5)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B11"];
                functionRange2.Formula = "=MyCalcStructExcelHandleMul(B10)";

                Assert.Equal("0.75", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C10"];
                functionRange1.Formula = "=MyCreateCalcExcelHandleExternal(2.5, 0.2)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C11"];
                functionRange2.Formula = "=MyCalcExcelHandleExternalMul(C10)";

                Assert.Equal("0.5", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D10"];
                functionRange1.Formula = "=MyGetExecutingAssembly()";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D11"];
                functionRange2.Formula = "=MyGetAssemblyName(D10)";

                Assert.Equal("ExcelDna.AddIn.RuntimeTests", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void TaskObjectHandles()
        {
            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateCalc(8, 9)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyTaskCalcSum(B1)";

                Assert.Equal("17", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateCalc(10, 11)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyTaskCalcDoubleSumWithCancellation(C1)";

                Assert.Equal("42", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyTaskCreateCalc(0, 12, 13)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyCalcSum(D1)";

                Assert.Equal("25", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                functionRange1.Formula = "=MyTaskCreateCalcWithCancellation(0, 14, 15)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                functionRange2.Formula = "=MyCalcSum(E1)";

                Assert.Equal("29", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                functionRange1.Formula = "=MyTaskCreateCalcExcelHandle(0, 0.1, 30)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                functionRange2.Formula = "=MyCalcExcelHandleMul(F1)";

                Assert.Equal("3", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void ObjectHandlesDisposable()
        {
            string b1;
            {
                Range functionRangeC1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRangeC1.Formula = "=MyGetCreatedDisposableObjectsCount()";
                int initialCreatedObjectsCount = (int)functionRangeC1.Value;

                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyCreateDisposableObject(1)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Range functionRangeC2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRangeC2.Formula = "=MyGetCreatedDisposableObjectsCount()";
                int finalCreatedObjectsCount = (int)functionRangeC2.Value;

                Assert.Equal(1, finalCreatedObjectsCount - initialCreatedObjectsCount);

                b1 = functionRange1.Value.ToString();
                Assert.Equal("1", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateDisposableObject(5)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal("2", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyCreateDisposableObject(1)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal(b1, functionRange1.Value.ToString());
                Assert.Equal("2", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void TaskObjectHandlesDisposable()
        {
            string b1;
            {
                Range functionRangeC1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRangeC1.Formula = "=MyGetCreatedDisposableObjectsCount()";
                int initialCreatedObjectsCount = (int)functionRangeC1.Value;

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";
                int initialDisposableObjectsCount = (int)functionRange2.Value;

                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Formula = "=MyTaskCreateDisposableObject(0, 1)";
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";
                int finalDisposableObjectsCount = (int)functionRange2.Value;

                Range functionRangeC2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRangeC2.Formula = "=MyGetCreatedDisposableObjectsCount()";
                int finalCreatedObjectsCount = (int)functionRangeC2.Value;

                Assert.Equal(1, finalCreatedObjectsCount - initialCreatedObjectsCount);

                b1 = functionRange1.Value.ToString();
                Assert.Equal(1, finalDisposableObjectsCount - initialDisposableObjectsCount);
            }

            {
                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";
                int initialDisposableObjectsCount = (int)functionRange2.Value;

                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyTaskCreateDisposableObject(0, 5)";

                functionRange2.Formula = "=MyGetDisposableObjectsCount()";
                int finalDisposableObjectsCount = (int)functionRange2.Value;

                Assert.NotEqual(b1, functionRange1.Value.ToString());
                Assert.Equal(1, finalDisposableObjectsCount - initialDisposableObjectsCount);
            }

            {
                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";
                int initialDisposableObjectsCount = (int)functionRange2.Value;

                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange1.Formula = "=MyTaskCreateDisposableObject(0, 1)";

                functionRange2.Formula = "=MyGetDisposableObjectsCount()";
                int finalDisposableObjectsCount = (int)functionRange2.Value;

                Assert.Equal(b1, functionRange1.Value.ToString());
                Assert.Equal(0, finalDisposableObjectsCount - initialDisposableObjectsCount);
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Range()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyRange(B2)";
                Assert.Equal("$B$2", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyRange(B2:C4)";
                Assert.Equal("$B$2:$C$4", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyRange((B2,D5:E6))";
                Assert.Equal("$B$2,$D$5:$E$6", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Params()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=MyParamsFunc1(1,\"2\",4,5)";
                Assert.Equal("1,2, : 2", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange.Formula = "=MyParamsFunc2(\"a\",,\"c\",\"d\",,\"f\")";
                Assert.Equal("a,,c, [3: d,ExcelDna.Integration.ExcelMissing,f]", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B3"];
                functionRange.Formula = "=MyParamsJoinString(\"//\",\"5\",\"4\",\"3\")";
                Assert.Equal("5//4//3", functionRange.Value.ToString());
            }
        }
    }
}
