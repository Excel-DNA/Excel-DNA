using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    [ExcelTestSettings(OutOfProcess = true)]
    [Collection("OutOfProcess")]
    public class RegistrationSampleOutOfProcess
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void AsyncSleep()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaDelayedHello(\"world\", 0)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello world!", 1000);

                Assert.Equal("Hello world!", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaDelayedHello(\"world\", 200)";

                Automation.WaitFor(() => functionRange.Value?.ToString() == "Hello world!", 1000);

                Assert.Equal("Hello world!", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSample)]
        public void GettingData()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaDelayedHelloAsync(\"a\", 2000)";
                Assert.Equal(-2146826245, functionRange.Value);
            }
        }
    }
}
