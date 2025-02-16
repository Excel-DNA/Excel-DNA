#if NETCOREAPP
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.Loader;

namespace ExcelDna.ManagedHost
{
    public class ExcelDnaAssemblyLoadContext : AssemblyLoadContext
    {
        readonly string _basePath;
        readonly AssemblyDependencyResolver _resolver;
        private Dictionary<string, string> unmanagedDllsResolutionCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public ExcelDnaAssemblyLoadContext(string basePath, bool isCollectible)
            : base($"ExcelDnaAssemblyLoadContext_{Path.GetFileNameWithoutExtension(basePath)}", isCollectible: isCollectible)
        {
            _basePath = basePath;

            if (System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported)
                _resolver = new AssemblyDependencyResolver(basePath);

#if DEBUG
            this.Resolving += ExcelDnaAssemblyLoadContext_Resolving;
            this.ResolvingUnmanagedDll += ExcelDnaAssemblyLoadContext_ResolvingUnmanagedDll;
#endif
        }

        protected override Assembly Load(AssemblyName assemblyName)
        {
            // CONSIDER: Should we consider priorities for packed vs local files?

            // First try the regular load path
            string assemblyPath = _resolver?.ResolveAssemblyToPath(assemblyName);
            if (assemblyPath != null)
            {
                return LoadFromAssemblyPath(assemblyPath);
            }

            // Finally we try the AssemblyManager
            return AssemblyManager.AssemblyResolve(assemblyName, false);
        }

        protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
        {
            string libraryPath = null;
            if (unmanagedDllsResolutionCache.TryGetValue(unmanagedDllName, out string cachedValue))
                libraryPath = cachedValue;

            if (libraryPath == null)
                libraryPath = _resolver?.ResolveUnmanagedDllToPath(unmanagedDllName);

            if (libraryPath == null)
                libraryPath = ResolveDllFromBaseDirectory(unmanagedDllName);

            if (libraryPath == null)
                libraryPath = AssemblyManager.NativeLibraryResolve(unmanagedDllName);

            unmanagedDllsResolutionCache[unmanagedDllName] = libraryPath;
            if (libraryPath != null)
            {
                return LoadUnmanagedDllFromPath(libraryPath);
            }

            return IntPtr.Zero;
        }

#if DEBUG
        // NOTE: The resolving events whould not be used if we are handling Load
        //       Added here for extra debugging
        Assembly ExcelDnaAssemblyLoadContext_Resolving(AssemblyLoadContext arg1, AssemblyName arg2)
        {
            Debug.Print($"Resolving event in {arg1.Name} for {arg2}");
            return null;
        }

        IntPtr ExcelDnaAssemblyLoadContext_ResolvingUnmanagedDll(Assembly arg1, string arg2)
        {
            Debug.Print($"ResolvingUnmanagedDll event from assembly {arg1.FullName} for {arg2}");
            return IntPtr.Zero;
        }
#endif

        internal Assembly LoadFromAssemblyBytes(byte[] assemblyBytes, byte[] pdbBytes)
        {
            if (pdbBytes == null)
            {
                return LoadFromStream(new MemoryStream(assemblyBytes));
            }
            else
            {
                return LoadFromStream(new MemoryStream(assemblyBytes), new MemoryStream(pdbBytes));
            }
        }

        private string ResolveDllFromBaseDirectory(string dllName)
        {
            string result = Path.Combine(Path.GetDirectoryName(_basePath), dllName);
            if (File.Exists(result))
                return result;

            return null;
        }
    }
}
#endif
