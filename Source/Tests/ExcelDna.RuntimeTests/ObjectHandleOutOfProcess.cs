using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class ObjectHandleOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void ThreadSafe()
        {
            for (int i = 1; i <= 5; ++i)
            {
                {
                    Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"C{i}"];
                    functionRange.Formula = $"=MyCreateCalcTS({(i - 1) * 20}, 0)";
                }
                {
                    Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"D{i}"];
                    functionRange.Formula = $"=MyCalcSumTS(C{i})";
                }
            }

            for (int i = 1; i <= 5; ++i)
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range[$"D{i}"];
                Assert.Equal($"{(i - 1) * 20}", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void Disposable()
        {
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

                Assert.Equal("1", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange1.Formula = "=MyCreateDisposableObject(5)";

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal("2", functionRange2.Value.ToString());
            }

            {
                Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange1.Clear();

                Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                Assert.Equal("1", functionRange2.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void TaskObjectStableCreate()
        {
            string v1;
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A1"];
                functionRange.Formula = $"=MyTaskCreateCalc(100, 1, 2)";
                {
                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A2"];
                    functionRange2.Formula = $"=MyCalcSum(A1)";
                    Automation.WaitFor(() => functionRange2.Value?.ToString() == "3", 3000);
                    Assert.Equal("3", functionRange2.Value.ToString());
                }
                v1 = functionRange.Value.ToString();
            }

            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = $"=MyTaskCreateCalc(100, 1, 2)";
                {
                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                    functionRange2.Formula = $"=MyCalcSum(B1)";
                    Automation.WaitFor(() => functionRange2.Value?.ToString() == "3", 3000);
                    Assert.Equal("3", functionRange2.Value.ToString());
                }
                string v2 = functionRange.Value.ToString();
                Assert.Equal(v1, v2);
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void TaskDisposable()
        {
            foreach (int delay in new[] { 0, 500 })
            {
                {
                    Range functionRangeC1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                    functionRangeC1.Formula = "=MyGetCreatedDisposableObjectsCount()";
                    int initialCreatedObjectsCount = (int)functionRangeC1.Value;

                    Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                    functionRange1.Formula = $"=MyTaskCreateDisposableObject({delay}, 1)";
                    Automation.WaitFor(() => functionRange1.Value.ToString().Contains("DisposableObject"), 3000);

                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                    functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                    Range functionRangeC2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                    functionRangeC2.Formula = "=MyGetCreatedDisposableObjectsCount()";
                    int finalCreatedObjectsCount = (int)functionRangeC2.Value;

                    Assert.Equal(1, finalCreatedObjectsCount - initialCreatedObjectsCount);

                    Assert.Equal("1", functionRange2.Value.ToString());
                }

                {
                    Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                    functionRange1.Formula = $"=MyTaskCreateDisposableObject({delay}, 5)";
                    Automation.WaitFor(() => functionRange1.Value.ToString().Contains("DisposableObject"), 3000);

                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C2"];
                    functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                    Assert.Equal("2", functionRange2.Value.ToString());
                }

                {
                    Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                    functionRange1.Clear();

                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B2"];
                    functionRange2.Formula = "=MyGetDisposableObjectsCount()";

                    Assert.Equal("1", functionRange2.Value.ToString());
                }
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RuntimeTests)]
        public void AsyncObjectCreate()
        {
            foreach (int delay in new[] { 0, 500 })
            {
                {
                    Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E1"];
                    functionRange1.Formula = $"=MyAsyncCreateCalc({delay}, 14, 15)";
                    Automation.WaitFor(() => functionRange1.Value.ToString().Contains("Calc"), 3000);

                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["E2"];
                    functionRange2.Formula = "=MyCalcSum(E1)";

                    Assert.Equal("29", functionRange2.Value.ToString());
                }

                {
                    Range functionRange1 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F1"];
                    functionRange1.Formula = $"=MyAsyncCreateCalcWithCancellation({delay}, 1, 2)";
                    Automation.WaitFor(() => functionRange1.Value.ToString().Contains("Calc"), 3000);

                    Range functionRange2 = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["F2"];
                    functionRange2.Formula = "=MyCalcSum(F1)";

                    Assert.Equal("3", functionRange2.Value.ToString());
                }
            }
        }
    }
}
