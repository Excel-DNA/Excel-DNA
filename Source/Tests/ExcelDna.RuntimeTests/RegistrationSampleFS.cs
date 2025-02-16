using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class RegistrationSampleFS
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSampleFS)]
        public void Optional()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D3"];
                functionRange.Formula = "=dnaFSharpOptional(A3,B3,C3)";
                Assert.Equal("Value: 0.000000, String: , Bool: false", functionRange.Value.ToString());
            }
            {
                Range functionRangeA = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A4"];
                functionRangeA.Value = "0";

                Range functionRangeB = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B4"];
                functionRangeB.Value = "zero";

                Range functionRangeC = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C4"];
                functionRangeC.Value = "0";

                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D4"];
                functionRange.Formula = "=dnaFSharpOptional(A4,B4,C4)";
                Assert.Equal("Value: 0.000000, String: zero, Bool: false", functionRange.Value.ToString());
            }
            {
                Range functionRangeA = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["A5"];
                functionRangeA.Value = "1";

                Range functionRangeB = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B5"];
                functionRangeB.Value = "one";

                Range functionRangeC = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C5"];
                functionRangeC.Value = "1";

                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D5"];
                functionRange.Formula = "=dnaFSharpOptional(A5,B5,C5)";
                Assert.Equal("Value: 1.000000, String: one, Bool: true", functionRange.Value.ToString());
            }
            {
                Range functionRangeC = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C6"];
                functionRangeC.Value = "FALSE";

                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D6"];
                functionRange.Formula = "=dnaFSharpOptional(A6,B6,C6)";
                Assert.Equal("Value: 0.000000, String: , Bool: false", functionRange.Value.ToString());
            }
            {
                Range functionRangeC = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C7"];
                functionRangeC.Value = "TRUE";

                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D7"];
                functionRange.Formula = "=dnaFSharpOptional(A7,B7,C7)";
                Assert.Equal("Value: 0.000000, String: , Bool: true", functionRange.Value.ToString());
            }
        }
    }
}
