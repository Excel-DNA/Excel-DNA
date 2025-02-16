//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Logging;

namespace ExcelDna.Integration
{
    // Loads the managed assembly, finds all the methods to be exported to Excel
    // and build the method information.

    // DOCUMENT: There is a lot of magic here, and many arbitrary decisions about what to register, and how.
    internal class AssemblyLoader
    {
        // We consolidate the TraceSources for both ExcelDna.Integration and ExcelDna.Loader under the Excel.Integration name 
        // (since it is the public contract for ExcelDna).
        // For the first version we don't make separate TraceSources for each class, though in future we might specialize under 
        // the ExcelDna.Integration namespace, so listening to ExcelDna.Integration* will be the forward-compatible pattern. 

        // Consolidated processing so we only have a single pass through the types.
        // CONSIDER: This is pretty ugly right now (both the flow and the names.)
        //           Try some fancy visitor pattern?
        public static void ProcessAssemblies(
                    List<ExportedAssembly> assemblies,
                    List<MethodInfo> methods,
                    List<ExtendedRegistration.ExcelParameterConversion> excelParameterConversions,
                    List<ExtendedRegistration.ExcelFunctionProcessor> excelFunctionProcessors,
                    List<Registration.ExcelFunctionRegistration> excelFunctionsExtendedRegistration,
                    List<Registration.FunctionExecutionHandlerSelector> excelFunctionExecutionHandlerSelectors,
                    List<ExcelAddInInfo> addIns,
                    List<Type> rtdServerTypes,
                    List<ExcelComClassType> comClassTypes)
        {
            bool loadRibbons = (ExcelDnaUtil.ExcelVersion >= 12.0);

            foreach (ExportedAssembly assembly in assemblies)
            {
                int initialObjectsCount = methods.Count +
                    excelParameterConversions.Count +
                    excelFunctionProcessors.Count +
                    excelFunctionsExtendedRegistration.Count +
                    excelFunctionExecutionHandlerSelectors.Count +
                    addIns.Count +
                    rtdServerTypes.Count +
                    comClassTypes.Count;
                Logger.Initialization.Verbose("Processing assembly {0}. ExplicitExports {1}, ExplicitRegistration {2}, ComServer {3}, IsDynamic {4}",
                    assembly.Assembly.FullName, assembly.ExplicitExports, assembly.ExplicitRegistration, assembly.ComServer, assembly.IsDynamic);

                int assemblyAttributes = ObjectHandles.ObjectHandleRegistration.ProcessAssemblyAttributes(assembly.Assembly.GetCustomAttributes());

                // Patch contributed by y_i on CodePlex:
                // http://stackoverflow.com/questions/11915389/assembly-gettypes-throwing-an-exception
                Type[] types;
                try
                {
                    // NOTE: The fact that NonPublic types are returned here, and processed as-if they were public
                    //       was a mistake. But it would be a serious breaking change to go back, so we'll leave it as is.
                    types = assembly.Assembly.GetTypes();
                }
                catch (ReflectionTypeLoadException e)
                {
                    Logger.Initialization.Verbose("Assembly.GetTypes() error: {0}", e.ToString());
                    // From MSDN:
                    // [...]contains the array of classes (Types property) that were defined in the module and were loaded. 
                    // The array can contain some null values.
                    types = e.Types;
                }

                bool explicitExports = assembly.ExplicitExports;
                bool explicitRegistration = assembly.ExplicitRegistration;
                foreach (Type type in types)
                {
                    if (type == null) continue; // We might get nulls from ReflectionTypeLoadException above

                    Logger.Initialization.Verbose("Processing type {0}", type.FullName);
                    try
                    {
                        object[] attribs = type.GetCustomAttributes(false);
                        bool isRtdServer;

                        if (!explicitRegistration)
                        {
                            GetExcelParameterConversions(type, excelParameterConversions);
                            GetExcelFunctionProcessors(type, excelFunctionProcessors);
                            GetExcelMethods(type, explicitExports, methods, excelFunctionsExtendedRegistration);
                            GetExcelFunctionExecutionHandlerSelectors(type, excelFunctionExecutionHandlerSelectors);
                        }
                        GetExcelAddIns(assembly, new TypeHelperDynamic(type), loadRibbons, addIns);
                        GetRtdServerTypes(type, rtdServerTypes, out isRtdServer);
                        GetComClassTypes(assembly, type, attribs, isRtdServer, comClassTypes);
                    }
                    catch (Exception e)
                    {
                        Logger.Initialization.Warn("Type {0} could not be processed. Error: {1}", type.FullName, e.ToString());
                    }
                }

                if (methods.Count +
                    excelParameterConversions.Count +
                    excelFunctionProcessors.Count +
                    excelFunctionsExtendedRegistration.Count +
                    excelFunctionExecutionHandlerSelectors.Count +
                    addIns.Count +
                    rtdServerTypes.Count +
                    comClassTypes.Count +
                    assemblyAttributes
                    == initialObjectsCount)
                {
                    Logger.Initialization.Error("No objects loaded from {0}", assembly.Assembly.FullName);
                }
            }
        }

        static void GetExcelParameterConversions(Type t, List<ExtendedRegistration.ExcelParameterConversion> excelParameterConversions)
        {
            MethodInfo[] mis = t.GetMethods(BindingFlags.Public | BindingFlags.Static);
            foreach (MethodInfo mi in mis)
            {
                if (IsParameterConversion(mi))
                {
                    excelParameterConversions.Add(new ExtendedRegistration.ExcelParameterConversion(mi));
                }
            }
        }

        static void GetExcelFunctionProcessors(Type t, List<ExtendedRegistration.ExcelFunctionProcessor> excelFunctionProcessors)
        {
            MethodInfo[] mis = t.GetMethods(BindingFlags.Public | BindingFlags.Static);
            foreach (MethodInfo mi in mis)
            {
                if (IsFunctionProcessor(mi))
                {
                    excelFunctionProcessors.Add(new ExtendedRegistration.ExcelFunctionProcessor(mi));
                }
            }
        }

        static void GetExcelMethods(Type t, bool explicitExports, List<MethodInfo> excelMethods, List<Registration.ExcelFunctionRegistration> excelFunctionsExtendedRegistration)
        {
            // DOCUMENT: Exclude if not a class, not public, /*abstract,*/ an array,  
            // open generic type or in "My" namespace.
            // Some basic checks -- what else?
            // TODO: Sort out exactly which types to export.
            if (!t.IsClass || !t.IsPublic ||
                /*t.IsAbstract ||*/ t.IsArray ||
                (t.IsGenericType && t.ContainsGenericParameters) ||
                t.Namespace == "My")
            {
                // Ignored cases
                Logger.Initialization.Info("Type ignored: {0}", t.FullName);
                return;
            }

            MethodInfo[] mis = t.GetMethods(BindingFlags.Public | BindingFlags.Static);
            // Filter list first - LINQ would be nice here :-)
            foreach (MethodInfo mi in mis)
            {
                bool isSupported = IsMethodSupported(mi, explicitExports);

                if (!isSupported && IsMethodMarkedForExport(mi))
                {
                    excelFunctionsExtendedRegistration.Add(new Registration.ExcelFunctionRegistration(mi));
                }
                else if (!isSupported)
                {
                    // CONSIDER: More detailed logging
                    Logger.Initialization.Info("Method not registered - unsupported signature, abstract or generic: '{0}.{1}'", mi.DeclaringType.Name, mi.Name);
                }

                if (isSupported)
                    excelMethods.Add(mi);
            }
        }

        static void GetExcelFunctionExecutionHandlerSelectors(Type type, List<Registration.FunctionExecutionHandlerSelector> excelFunctionExecutionHandlerSelectors)
        {
            MethodInfo[] mis = type.GetMethods(BindingFlags.Public | BindingFlags.Static);
            foreach (MethodInfo mi in mis)
            {
                if (IsExcelFunctionExecutionHandlerSelector(mi))
                    excelFunctionExecutionHandlerSelectors.Add((IExcelFunctionInfo functionInfo) => (Registration.IFunctionExecutionHandler)mi.Invoke(null, new object[] { functionInfo }));
            }
        }

        static bool IsMethodSupported(MethodInfo mi, bool explicitExports)
        {
            var isSupported = true;

            // Skip generic methods - these may appear even though we have skipped generic types, 
            // e.g. in F# --standalone assemblies
            if (mi.IsAbstract || mi.IsGenericMethod)
            {
                isSupported = false;
            }
            // if explicitexports - check that this method is marked
            else if (explicitExports && !IsMethodMarkedForExport(mi))
            {
                isSupported = false;
            }
            else if (ObjectHandles.ObjectHandleRegistration.IsMethodSupported(new ExcelDna.Registration.ExcelFunctionRegistration(mi)))
            {
                isSupported = false;
            }
            else if (!IsPrimitiveParameterType(mi.ReturnType))
            {
                isSupported = false;
            }
            else if (IsExcelAsyncFunction(mi))
            {
                isSupported = false;
            }
            else if (ExcelDna.Registration.ParamsRegistration.IsParamsMethod(new ExcelDna.Registration.ExcelFunctionRegistration(mi)))
            {
                isSupported = false;
            }
            else
            {
                foreach (ParameterInfo pi in mi.GetParameters())
                {
                    if (pi.IsOptional || !IsPrimitiveParameterType(pi.ParameterType))
                        isSupported = false;
                }
            }

            return isSupported;
        }

        static bool IsExcelAsyncFunction(MethodInfo mi)
        {
            return HasCustomAttribute(mi, "ExcelDna.Integration.ExcelAsyncFunctionAttribute");
        }

        static bool IsExcelFunctionExecutionHandlerSelector(MethodInfo mi)
        {
            return HasCustomAttribute(mi, "ExcelDna.Integration.ExcelFunctionExecutionHandlerSelectorAttribute");
        }

        // CAUTION: This check needs to match the usage in ExcelDna.Loader.XlMethodInfo.SetAttributeInfo()
        static bool IsMethodMarkedForExport(MethodInfo mi)
        {
            object[] atts = mi.GetCustomAttributes(false);
            foreach (object att in atts)
            {
                Type attType = att.GetType();
                if (TypeHasAncestorWithFullName(attType, "ExcelDna.Integration.ExcelFunctionAttribute") ||
                    TypeHasAncestorWithFullName(attType, "ExcelDna.Integration.ExcelCommandAttribute"))
                {
                    return true;
                }
            }
            return false;
        }

        static readonly List<Type> _primitiveParameterTypes = new List<Type>
        {
            typeof(double),
            typeof(string),
            typeof(DateTime),
            typeof(double[]),
            typeof(double[,]),
            typeof(object),
            typeof(object[]),
            typeof(object[,]),
            typeof(bool),
            typeof(int),
            typeof(short),
            typeof(ushort),
            typeof(decimal),
            typeof(long),
            typeof(void)
        };

        static public bool IsPrimitiveParameterType(Type type)
        {
            return _primitiveParameterTypes.Contains(type) ||
                   (ExcelDnaUtil.ExcelVersion >= 14.0 && type == typeof(ExcelAsyncHandle));    // Only Excel 2010+
        }

        // Some support for creating add-ins that are notified of open and close
        // this allows the add-in to add menus, toolbar buttons etc.
        // Also records whether this class should be loaded as a ComAddIn (for the Ribbon).
        public class ExcelAddInInfo
        {
            public MethodInfo AutoOpenMethod;
            public MethodInfo AutoCloseMethod;
            public bool IsCustomUI;
            public object Instance;
            public DnaLibrary ParentDnaLibrary;
        }

        static public void GetExcelAddIns(ExportedAssembly assembly, ITypeHelper t, bool loadRibbons, List<ExcelAddInInfo> addIns)
        {
            // NOTE: We probably should have restricted this to public types, but didn't. Now it's too late.
            //       So internal classes that implement IExcelAddIn are also loaded.
            try
            {
                if (!t.Type.IsClass || t.Type.IsAbstract)
                {
                    Logger.Registration.Verbose("GetExcelAddIns - Skipped add-in object of type: {0}", t.Type.FullName);
                    return;
                }

                Type addInType = t.Type.GetInterface("ExcelDna.Integration.IExcelAddIn");
                bool isRibbon = IsRibbonType(t.Type);
                if (addInType != null || (isRibbon && loadRibbons))
                {
                    ExcelAddInInfo info = new ExcelAddInInfo();
                    if (addInType != null)
                    {
                        info.AutoOpenMethod = typeof(IExcelAddIn).GetMethod("AutoOpen");
                        info.AutoCloseMethod = typeof(IExcelAddIn).GetMethod("AutoClose");
                    }
                    info.IsCustomUI = isRibbon;
                    info.Instance = t.CreateInstance();
                    info.ParentDnaLibrary = assembly?.ParentDnaLibrary;
                    addIns.Add(info);
                    Logger.Registration.Verbose("GetExcelAddIns - Created add-in object of type: {0}", t.Type.FullName);
                }
            }
            catch (Exception e) // I think only CreateInstance can throw an exception here...
            {
                Logger.Initialization.Warn("GetExcelAddIns CreateInstance problem for type: {0} - exception: {1}", t.Type.FullName, e);
            }

        }

        // DOCUMENT: We register types that implement an interface with the IRtdServer Guid. These include
        //           "Microsoft.Office.Interop.Excel.IRtdServer" and
        //           "ExcelDna.Integration.Rtd.IRtdServer".
        // The RTD server can be accessed using the ExcelDnaUtil.RTD function under the 
        // FullName of the type, or under the ProgId defined in an attribute, if there is one.
        static public void GetRtdServerTypes(Type t, List<Type> rtdServerTypes, out bool isRtdServer)
        {
            isRtdServer = false;
            Type[] itfs = t.GetInterfaces();
            foreach (Type itf in itfs)
            {
                if (itf.GUID == ComAPI.guidIRtdServer)
                {
                    //object[] attrs = t.GetCustomAttributes(typeof(ProgIdAttribute), false);
                    //if (attrs.Length >= 1)
                    //{
                    //    ProgIdAttribute progIdAtt = (ProgIdAttribute)attrs[0];
                    //    rtdServerTypes[progIdAtt.Value] = t;
                    //}
                    //rtdServerTypes[t.FullName] = t;
                    rtdServerTypes.Add(t);
                    isRtdServer = true;
                    Logger.Initialization.Verbose("GetRtdServerTypes - Found RTD server type: {0}", t.FullName);
                }
            }
        }

        // DOCUMENT: We register ComVisible types that
        //           (implement IRtdServer OR are in ExternalLibraries marked ComServer='true'
        static public void GetComClassTypes(ExportedAssembly assembly, Type type, object[] attributes, bool isRtdServer, List<ExcelComClassType> comClassTypes)
        {
            if (!Marshal.IsTypeVisibleFromCom(type))
            {
                return;
            }

            if (isRtdServer || assembly.ComServer)
            {
                string progId;
                Guid clsId;

                // Check for public default constructor
                if (type.GetConstructor(BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null) == null)
                {
                    // No use to us here - won't be able to construct.
                    return;
                }

                if (assembly.IsDynamic)
                {
                    // Check that we have an explicit Guid attribute
                    object[] attrs = type.GetCustomAttributes(typeof(GuidAttribute), false);
                    if (attrs.Length == 0)
                    {
                        // No Guid attribute - skip this type.
                        return;
                    }
                    else
                    {
                        GuidAttribute guidAtt = (GuidAttribute)attrs[0];
                        clsId = new Guid(guidAtt.Value);
                    }
                }
                else
                {
                    clsId = Marshal.GenerateGuidForType(type);
                }

                progId = Marshal.GenerateProgIdForType(type);

                ExcelComClassType comClassType = new ExcelComClassType
                {
                    Type = type,
                    ClsId = clsId,
                    ProgId = progId,
                    IsRtdServer = isRtdServer,
                    TypeLibPath = assembly.TypeLibPath
                };
                comClassTypes.Add(comClassType);
                Logger.Initialization.Verbose("GetComClassTypes - Found type {0}, with ProgId {1}", type.FullName, progId);
            }
        }

        static bool IsRibbonType(Type type)
        {
            // Ribbon type is one that has ExcelRibbon as an ancestor (but is not ExcelRibbon itself), is not abstract, and it's parent is not a ribbon type

            // We are trying to prevent loading multiple copies of a ribbon along the inheritance hierarchy, 
            // while still allowing some abstraction of Ribbon handling classes.

            // Current design will load only the least-derived concrete class.

            bool isRibbon =
                    type != null &&
                    TypeHasAncestorWithFullName(type.BaseType, "ExcelDna.Integration.CustomUI.ExcelRibbon") &&
                    !type.IsAbstract &&
                    !IsRibbonType(type.BaseType);

            return isRibbon;
        }

        private static bool IsParameterConversion(MethodInfo methodInfo)
        {
            return HasCustomAttribute(methodInfo, "ExcelDna.Integration.ExcelParameterConversionAttribute");
        }

        private static bool IsFunctionProcessor(MethodInfo methodInfo)
        {
            return HasCustomAttribute(methodInfo, "ExcelDna.Integration.ExcelFunctionProcessorAttribute");
        }

        private static bool HasCustomAttribute(MethodInfo methodInfo, string attributeName)
        {
            object[] atts = methodInfo.GetCustomAttributes(false);
            foreach (object att in atts)
            {
                Type attType = att.GetType();
                if (TypeHasAncestorWithFullName(attType, attributeName))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool TypeHasAncestorWithFullName(Type type, string fullName)
        {
            if (type == null) return false;
            if (type.FullName == fullName) return true;
            return TypeHasAncestorWithFullName(type.BaseType, fullName);
        }
    }
}
