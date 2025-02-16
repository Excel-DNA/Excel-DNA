using ExcelDna.Integration;

namespace ExcelDna.AddIn.RuntimeTestsAOT
{
    public class Functions
    {
        [ExcelFunction]
        public static string NativeHello(string name)
        {
            return $"Hello {name}!";
        }

        [ExcelFunction]
        public static int NativeSum(int i1, int i2)
        {
            return i1 + i2;
        }
    }
}
