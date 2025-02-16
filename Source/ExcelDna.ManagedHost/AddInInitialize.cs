using ExcelDna.Loader;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelDna.ManagedHost
{
    public unsafe class AddInInitialize
    {
#if NETFRAMEWORK
        public static bool Initialize32(int xlAddInExportInfoAddress, int hModuleXll, string pathXll)
        {
            // NOTE: The sequence here is important - we install the AssemblyManage which can resolve packed assemblies
            //       before calling LoadIntegration, which will be the first time we try to resolve ExcelDna.Integration
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll);
            // TODO: Load up the DnaFile and Assembly names ?
            AppDomain.CurrentDomain.AssemblyResolve += (object sender, ResolveEventArgs args) => AssemblyManager.AssemblyResolve(new AssemblyName(args.Name), true);
            return XlAddInInitialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, AssemblyManager.GetResourceBytes, AssemblyManager.LoadFromAssemblyPath, AssemblyManager.LoadFromAssemblyBytes, Logger.SetIntegrationTraceSource);
        }

        public static bool Initialize64(long xlAddInExportInfoAddress, long hModuleXll, string pathXll)
        {
            // NOTE: The sequence here is important - we install the AssemblyManage which can resolve packed assemblies
            //       before calling LoadIntegration, which will be the first time we try to resolve ExcelDna.Integration
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll);
            AppDomain.CurrentDomain.AssemblyResolve += (object sender, ResolveEventArgs args) => AssemblyManager.AssemblyResolve(new AssemblyName(args.Name), true);
            return XlAddInInitialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, AssemblyManager.GetResourceBytes, AssemblyManager.LoadFromAssemblyPath, AssemblyManager.LoadFromAssemblyBytes, Logger.SetIntegrationTraceSource);
        }
#endif

#if NETCOREAPP
        static ExcelDnaAssemblyLoadContext _alc;

        [UnmanagedCallersOnly]
        public static short Initialize(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL, byte disableAssemblyContextUnload, void* pTempDirPath)
        {
            UnloadALC();
            ProcessStartupHooks();

            string pathXll = Marshal.PtrToStringUni((IntPtr)pPathXLL);
            string tempDirPath = Marshal.PtrToStringUni((IntPtr)pTempDirPath);
            _alc = new ExcelDnaAssemblyLoadContext(pathXll, disableAssemblyContextUnload == 0);
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll, _alc, Path.Combine(tempDirPath, "ExcelDna.ManagedHost"));
            var loaderAssembly = _alc.LoadFromAssemblyName(new AssemblyName("ExcelDna.Loader"));
            var xlAddInType = loaderAssembly.GetType("ExcelDna.Loader.XlAddIn");
            var initOK = (bool)xlAddInType.InvokeMember("Initialize", BindingFlags.Public | BindingFlags.Static | BindingFlags.InvokeMethod, null, null,
                new object[] { (IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, tempDirPath,
                    (Func<string, int, byte[]>)AssemblyManager.GetResourceBytes,
                    (Func<string, Assembly>)_alc.LoadFromAssemblyPath,
                    (Func<byte[], byte[], Assembly>)_alc.LoadFromAssemblyBytes,
                    (Action<TraceSource>)Logger.SetIntegrationTraceSource });

            return initOK ? (short)1 : (short)0;
        }

        public static short InitializeNativeAOT(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL, byte disableAssemblyContextUnload, void* pTempDirPath)
        {
            UnloadALC();
            ProcessStartupHooks();

            string pathXll = Marshal.PtrToStringUni((IntPtr)pPathXLL);
            string tempDirPath = Marshal.PtrToStringUni((IntPtr)pTempDirPath);
            _alc = new ExcelDnaAssemblyLoadContext(pathXll, disableAssemblyContextUnload == 0);
            AssemblyManager.Initialize((IntPtr)hModuleXll, pathXll, _alc, Path.Combine(tempDirPath, "ExcelDna.ManagedHost.NativeAOT"));
            var initOK = (bool)ExcelDna.Loader.XlAddIn.Initialize((IntPtr)xlAddInExportInfoAddress, (IntPtr)hModuleXll, pathXll, tempDirPath,
                    (Func<string, int, byte[]>)AssemblyManager.GetResourceBytes,
                    (Func<string, Assembly>)_alc.LoadFromAssemblyPath,
                    (Func<byte[], byte[], Assembly>)_alc.LoadFromAssemblyBytes,
                    (Action<TraceSource>)Logger.SetIntegrationTraceSource);

            return initOK ? (short)1 : (short)0;
        }

        private static void UnloadALC()
        {
            if (_alc == null)
                return;

            WeakReference alcWeakRef = StartUnloadALC();
            for (int i = 0; alcWeakRef.IsAlive && (i < 10); i++)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            if (alcWeakRef.IsAlive)
                Logger.Initialization.Warn("Failed to unload AssemblyLoadContext");
        }

        private static WeakReference StartUnloadALC()
        {
            AssemblyManager.ResetALC();

            WeakReference alcWeakRef = new WeakReference(_alc);
            _alc.Unload();
            _alc = null;
            return alcWeakRef;
        }

        private static void ProcessStartupHooks()
        {
            try
            {
                Type StartupHookProviderType = Type.GetType($"System.StartupHookProvider, System.Private.CoreLib, Version=6.0.0.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e");
                MethodInfo ProcessStartupHooksMethod = StartupHookProviderType.GetMethod("ProcessStartupHooks", BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.InvokeMethod);
                ProcessStartupHooksMethod.Invoke(null, null);
            }
            catch
            {
            }
        }
    }
#endif

#if !NETCOREAPP
        // NOTE: We need this code to be in a separate method, so that the assembly resolution for ExcelDna.Loader runs after the AssemblyManager is installed.
        [MethodImpl(MethodImplOptions.NoInlining)]
        static bool XlAddInInitialize(IntPtr xlAddInExportInfoAddress, IntPtr hModuleXll, string pathXll,
                Func<string, int, byte[]> getResourceBytes,
                Func<string, Assembly> loadFromAssemblyPath,
                Func<byte[], byte[], Assembly> loadFromAssemblyBytes,
                Action<TraceSource> setIntegrationTraceSource)
        {
            return XlAddIn.Initialize(xlAddInExportInfoAddress, hModuleXll, pathXll, null, getResourceBytes, loadFromAssemblyPath, loadFromAssemblyBytes, setIntegrationTraceSource);
        }
    }
#endif
}
