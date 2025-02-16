//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using SevenZip.Compression.LZMA;
using System.Threading;
using System.Globalization;
using ExcelDna.PackedResources;
using ExcelDna.PackedResources.Logging;

internal static class ResourceHelper
{
    internal enum TypeName
    {
        CONFIG = -1,
        ASSEMBLY = 0,
        DNA = 1,
        IMAGE = 2,
        SOURCE = 3,
        PDB = 4,
        NATIVE_LIBRARY = 5,
        DOC = 6,
        NATIVE_ASSEMBLY = 7,
    }

    // TODO: Learn about locales
    private const ushort localeNeutral = 0;
    private const ushort localeEnglishUS = 1033;
    private const ushort localeEnglishSA = 7177;

    [DllImport("version.dll", SetLastError = true)]
    private static extern uint GetFileVersionInfoSize(
        string lptstrFilename,
        out uint lpdwHandle);

    [DllImport("version.dll", SetLastError = true)]
    private static extern bool GetFileVersionInfo(
        string lptstrFilename,
        uint dwHandle,
        uint dwLen,
        byte[] lpData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr FindResource(
    IntPtr hModule,
    string lpName,
    string lpType);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadResource(
        IntPtr hModule,
        IntPtr hResInfo);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LockResource(
        IntPtr hResData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SizeofResource(
        IntPtr hModule,
        IntPtr hResInfo);

    [DllImport("kernel32.dll")]
    private static extern uint GetLastError();

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hReservedNull, LoadLibraryFlags dwFlags);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FreeLibrary(IntPtr hModule);

    [System.Flags]
    enum LoadLibraryFlags : uint
    {
        None = 0,
        DONT_RESOLVE_DLL_REFERENCES = 0x00000001,
        LOAD_IGNORE_CODE_AUTHZ_LEVEL = 0x00000010,
        LOAD_LIBRARY_AS_DATAFILE = 0x00000002,
        LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = 0x00000040,
        LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x00000020,
        LOAD_LIBRARY_SEARCH_APPLICATION_DIR = 0x00000200,
        LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = 0x00001000,
        LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = 0x00000100,
        LOAD_LIBRARY_SEARCH_SYSTEM32 = 0x00000800,
        LOAD_LIBRARY_SEARCH_USER_DIRS = 0x00000400,
        LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008
    }

    public static IntPtr LoadXllResources(string xllPath)
    {
        return LoadLibraryEx(xllPath, IntPtr.Zero, LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE);
    }

    public static bool FreeXllResources(IntPtr hModule)
    {
        return FreeLibrary(hModule);
    }

    internal class ResourceUpdater
    {
        private readonly IBuildLogger buildLogger;
        private readonly IResourceResolver resourceResolver;
        private readonly object lockResource = new object();

        private int typelibIndex = 0;
        private List<object> updateData = new List<object>();
        private Queue<ManualResetEvent> finishedTask = new Queue<ManualResetEvent>();

        public ResourceUpdater(string fileName, bool useManagedResourceResolver, IBuildLogger buildLogger)
        {
            this.buildLogger = buildLogger;
            if (useManagedResourceResolver)
            {
#if ASMRESOLVER
                resourceResolver = new ResourceResolverManaged();
#endif
            }
            else
            {
                resourceResolver = new ResourceResolverWin();
            }

            resourceResolver.Begin(fileName);
        }

        private void CompressDoUpdateHelper(byte[] content, string name, TypeName typeName, string source, bool compress)
        {
            if (compress)
            {
                content = SevenZipHelper.Compress(content);
            }

            DoUpdateResource(typeName.ToString() + (compress ? "_LZMA" : ""), name, source, content);
        }

        public string AddFile(byte[] content, string name, TypeName typeName, string source, bool compress, bool multithreading)
        {
            Debug.Assert(name == name.ToUpperInvariant());

            if (multithreading)
            {
                var mre = new ManualResetEvent(false);
                finishedTask.Enqueue(mre);
                ThreadPool.QueueUserWorkItem(delegate
                {
                    CompressDoUpdateHelper(content, name, typeName, source, compress);
                    mre.Set();
                }
                );
            }
            else
            {
                CompressDoUpdateHelper(content, name, typeName, source, compress);
            }

            return name;
        }

        public string AddAssembly(string path, string source, bool compress, bool multithreading, bool includePdb)
        {
            try
            {
                byte[] assemblyBytes = File.ReadAllBytes(path);
                AssemblyName assemblyName = AssemblyName.GetAssemblyName(path);
                CultureInfo cultureInfo = assemblyName.CultureInfo;
                string name = assemblyName.Name.ToUpperInvariant(); // .ToUpperInvariant().Replace(".", "_");

                // For .resources assemblies we add the Culture name to the packed name
                if (name.EndsWith(".RESOURCES") && cultureInfo != null && !string.IsNullOrEmpty(cultureInfo.Name))
                {
                    name += "." + cultureInfo.Name.ToUpperInvariant();
                }

                AddFile(assemblyBytes, name, TypeName.ASSEMBLY, source, compress, multithreading);

                string pdbFile = Path.ChangeExtension(path, "pdb");
                if (includePdb && File.Exists(pdbFile))
                {
                    byte[] pdbBytes = File.ReadAllBytes(pdbFile);
                    AddFile(pdbBytes, name, TypeName.PDB, source, compress, multithreading);
                }

                return name;
            }
            catch (Exception e)
            {
                buildLogger.Error(e, "Assembly at {0} could not be packed. Possibly a mixed assembly? (These are not supported yet.)\r\nException: {1}", path, e);
                return null;
            }
        }

        public int AddTypeLib(byte[] data)
        {
            lock (lockResource)
            {
                string typeName = "TYPELIB";
                typelibIndex++;

                buildLogger.Information("  ->  Updating typelib: Type: {0}, Index: {1}, Length: {2}", typeName, typelibIndex, data.Length);
                GCHandle pinHandle = GCHandle.Alloc(data, GCHandleType.Pinned);
                updateData.Add(pinHandle);

                bool result = resourceResolver.Update(typeName, (IntPtr)typelibIndex, localeNeutral, pinHandle.AddrOfPinnedObject(), (uint)data.Length);
                if (!result)
                {
                    throw new Win32Exception();
                }
            }

            return typelibIndex;
        }

        public void DoUpdateResource(string typeName, string name, string source, byte[] data)
        {
            lock (lockResource)
            {
                string sourceInfo = (source != null) ? $" Source: {source}," : null;
                buildLogger.Information($"  ->  Updating resource: Type: {typeName}, Name: {name},{sourceInfo} Length: {data.Length}");

                bool result = resourceResolver.Update(typeName, name, localeNeutral, data);
                if (!result)
                {
                    throw new Win32Exception();
                }
            }
        }

        public void RemoveResource(string typeName, string name)
        {
            lock (lockResource)
            {
                bool result = resourceResolver.Update(typeName, name, localeNeutral, null);
                if (!result)
                {
                    throw new Win32Exception();
                }
            }
        }

        public void CopyFileVersion(string fromFile)
        {
            uint ignored;
            uint versionSize = ResourceHelper.GetFileVersionInfoSize(fromFile, out ignored);
            if (versionSize == 0)
            {
                throw new Win32Exception();
            }

            byte[] versionBuf = new byte[versionSize];
            bool result = ResourceHelper.GetFileVersionInfo(fromFile, ignored, versionSize, versionBuf);
            if (!result)
            {
                throw new Win32Exception();
            }

            GCHandle versionBufHandle = GCHandle.Alloc(versionBuf, GCHandleType.Pinned);
            try
            {
                lock (lockResource)
                {
                    uint versionResourceType = 16;
                    uint versionResourceId = 1;
                    result = resourceResolver.Update(
                        versionResourceType,
                        versionResourceId,
                        localeNeutral,
                        versionBufHandle.AddrOfPinnedObject(),
                        versionSize);
                    if (!result)
                    {
                        throw new Win32Exception();
                    }
                }
            }
            finally
            {
                versionBufHandle.Free();
            }
        }

        public void EndUpdate()
        {
            EndUpdate(false);
        }

        public void EndUpdate(bool discard)
        {
            if (finishedTask.Count > 0)
            {
                while (finishedTask.Count > 0)
                {
                    int cnt = finishedTask.Count;
                    // WaitAll accepts a maximum of 64 WaitHandles
                    if (cnt > 64) cnt = 64;

                    ManualResetEvent[] mre = new ManualResetEvent[cnt];

                    for (int i = 0; i < cnt; i++)
                        mre[i] = finishedTask.Dequeue();

                    WaitHandle.WaitAll(mre);
                }
            }

            resourceResolver.End(discard);
        }

        // Load the resource, trying also as compressed if no uncompressed version is found.
        // If the resource type ends with "_LZMA", we decompress from the LZMA format.
        internal static byte[] LoadResourceBytes(IntPtr hModule, string typeName, string resourceName)
        {
            // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
            // Logger.Initialization.Info("LoadResourceBytes for resource {0} of type {1}", resourceName, typeName);
            IntPtr hResInfo = FindResource(hModule, resourceName, typeName);
            if (hResInfo == IntPtr.Zero)
            {
                // We expect this null result value when the resource does not exists.

                if (!typeName.EndsWith("_LZMA"))
                {
                    // Try the compressed name.
                    typeName += "_LZMA";
                    hResInfo = FindResource(hModule, resourceName, typeName);
                }
                if (hResInfo == IntPtr.Zero)
                {
                    // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
                    // Logger.Initialization.Info("Resource not found - resource {0} of type {1}", resourceName, typeName);
                    Debug.Print("ResourceHelper.LoadResourceBytes - Resource not found - resource {0} of type {1}", resourceName, typeName);
                    // Return null to indicate that the resource was not found.
                    return null;
                }
            }
            IntPtr hResData = LoadResource(hModule, hResInfo);
            if (hResData == IntPtr.Zero)
            {
                // Unexpected error - this should not happen
                // CAREFUL: Can't log here yet as this method is called during Integration.Initialize()
                //Logger.Initialization.Error("Unexpected errror loading resource {0} of type {1}", resourceName, typeName);
                Debug.Print("ResourceHelper.LoadResourceBytes - Unexpected errror loading resource {0} of type {1}", resourceName, typeName);
                throw new Win32Exception();
            }
            uint size = SizeofResource(hModule, hResInfo);
            IntPtr pResourceBytes = LockResource(hResData);
            byte[] resourceBytes = new byte[size];
            Marshal.Copy(pResourceBytes, resourceBytes, 0, (int)size);

            byte[] resultBytes;
            if (typeName.EndsWith("_LZMA"))
                resultBytes = Decompress(resourceBytes);
            else
                resultBytes = resourceBytes;
            return resultBytes;
        }

        private static byte[] Decompress(byte[] inputBytes)
        {
            MemoryStream newInStream = new MemoryStream(inputBytes);
            Decoder decoder = new Decoder();
            newInStream.Seek(0, 0);
            MemoryStream newOutStream = new MemoryStream();
            byte[] properties2 = new byte[5];
            if (newInStream.Read(properties2, 0, 5) != 5)
                throw (new Exception("input .lzma is too short"));
            long outSize = 0;
            for (int i = 0; i < 8; i++)
            {
                int v = newInStream.ReadByte();
                if (v < 0)
                    throw (new Exception("Can't Read 1"));
                outSize |= ((long)(byte)v) << (8 * i);
            }
            decoder.SetDecoderProperties(properties2);
            long compressedSize = newInStream.Length - newInStream.Position;
            decoder.Code(newInStream, newOutStream, compressedSize, outSize, null);
            byte[] b = newOutStream.ToArray();
            return b;
        }
    }
}
