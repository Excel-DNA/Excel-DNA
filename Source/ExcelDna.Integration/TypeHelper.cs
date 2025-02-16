using System;
using System.Linq;
using System.Reflection;

namespace ExcelDna.Integration
{
    public interface ITypeHelper
    {
        Type Type { get; }
        object CreateInstance();
    }

    public class TypeHelper<T> : ITypeHelper
    {
        public Type Type => typeof(T);

        public object CreateInstance()
        {
            return Activator.CreateInstance<T>();
        }
    }

    public class TypeHelperDynamic : ITypeHelper
    {
        public TypeHelperDynamic(Type t)
        {
            Type = t;
        }

        public Type Type { get; }

        public object CreateInstance()
        {
            return Activator.CreateInstance(Type);
        }
    }
}
