using System.Collections.Generic;
using System.Reflection;

namespace ExcelDna.Integration
{
    public class NativeAOT
    {
        public static bool IsActive { get; set; }

        public static List<MethodInfo> MethodsForRegistration { get; } = new List<MethodInfo>();
        public static List<ITypeHelper> ExcelAddIns { get; } = new List<ITypeHelper>();
    }
}
