using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelDna.RuntimeTests
{
    public class RegistrationSampleVB
    {
        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSampleVB)]
        public void Optional()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaOptionalAnswer()";
                Assert.Equal("The answer is 42", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaOptionalAnswer(7.5)";
                Assert.Equal("The answer is 7.5", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSampleVB)]
        public void Params()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaAddValues(3,6,12)";
                Assert.Equal(21, functionRange.Value);
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaConcatStrings(\"p\",\"-\",\"a\",\"b\")";
                Assert.Equal("pa-b", functionRange.Value.ToString());
            }
        }

        [ExcelFact(Workbook = "", AddIn = AddInPath.RegistrationSampleVB)]
        public void Range()
        {
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["B1"];
                functionRange.Formula = "=dnaVbRangeTest(B2)";
                Assert.Equal("$B$2", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["C1"];
                functionRange.Formula = "=dnaVbRangeTest(B2:C4)";
                Assert.Equal("$B$2:$C$4", functionRange.Value.ToString());
            }
            {
                Range functionRange = ((Worksheet)ExcelDna.Testing.Util.Workbook.Sheets[1]).Range["D1"];
                functionRange.Formula = "=dnaVbRangeTest((B2,D5:E6))";
                Assert.Equal("$B$2,$D$5:$E$6", functionRange.Value.ToString());
            }
        }
    }
}
