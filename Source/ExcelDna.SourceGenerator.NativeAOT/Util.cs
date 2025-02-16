using Microsoft.CodeAnalysis;
using System.Collections.Generic;

namespace ExcelDna.SourceGenerator.NativeAOT
{
    public class Util
    {
        public static string GetFullTypeName(ITypeSymbol type)
        {
            return type.ToDisplayString();
        }

        public static string GetFullMethodName(IMethodSymbol method)
        {
            return $"{GetFullTypeName(method.ContainingType)}.{method.Name}";
        }

        public static IEnumerable<ITypeSymbol> GetBaseTypesAndThis(ITypeSymbol type)
        {
            var current = type;
            while (current != null)
            {
                yield return current;
                current = current.BaseType;
            }
        }
    }
}
