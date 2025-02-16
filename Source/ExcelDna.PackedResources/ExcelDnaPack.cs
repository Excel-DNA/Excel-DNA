using System;
using System.Collections.Generic;
using System.IO;
using ExcelDna.Integration;
using ExcelDna.PackedResources.Logging;
using System.Reflection;
using System.Linq;

namespace ExcelDna.PackedResources
{
    internal class ExcelDnaPack
    {
        public static int Pack(string dnaPath, string xllOutputPathParam, bool compress, bool multithreading, bool overwrite, string usageInfo, List<string> filesToPublish, bool packNativeLibraryDependencies, bool packManagedDependencies, string excludeDependencies, bool useManagedResourceResolver, string outputBitness, string docPath, IBuildLogger buildLogger)
        {
            string HostXLL = outputBitness == "64" ? "ExcelDna64.xll" : "ExcelDna.xll";
            string dnaDirectory = Path.GetDirectoryName(dnaPath);
            string dnaFilePrefix = Path.GetFileNameWithoutExtension(dnaPath);
            string configPath = Path.ChangeExtension(dnaPath, ".xll.config");
            string xllInputPath = Path.Combine(dnaDirectory, dnaFilePrefix + ".xll");
            string xllOutputPath = Path.Combine(dnaDirectory, dnaFilePrefix + "-packed.xll");
            if (xllOutputPathParam != null)
                xllOutputPath = xllOutputPathParam;

            if (!File.Exists(dnaPath))
            {
                buildLogger.Error(typeof(ExcelDnaPack), "ERROR: Add-in .dna file {0} not found.\r\n\r\n{1}", dnaPath, usageInfo);
                return 1;
            }

            if (filesToPublish == null && File.Exists(xllOutputPath))
            {
                if (overwrite == false)
                {
                    Console.Write("Output .xll file " + xllOutputPath + " already exists. Overwrite? [Y/N] ");
                    string response = Console.ReadLine();
                    if (!response.Equals("Y", StringComparison.CurrentCultureIgnoreCase))
                    {
                        Console.WriteLine("\r\nNot overwriting existing file.\r\nExiting ExcelDnaPack.");
                        return 1;
                    }
                }

                try
                {
                    File.Delete(xllOutputPath);
                }
                catch
                {
                    buildLogger.Error(typeof(ExcelDnaPack), "ERROR: Existing output .xll file {0} could not be deleted. (Perhaps loaded in Excel?)\r\n\r\nExiting ExcelDnaPack.", xllOutputPath);
                    return 1;
                }
            }

            string outputDirectory = Path.GetDirectoryName(xllOutputPath);
            if (outputDirectory == String.Empty)
            {
                outputDirectory = ".";  // https://github.com/Excel-DNA/ExcelDna/issues/7
            }

            if (filesToPublish == null && !Directory.Exists(outputDirectory))
            {
                try
                {
                    Directory.CreateDirectory(outputDirectory);
                }
                catch (Exception ex)
                {
                    buildLogger.Error(typeof(ExcelDnaPack), "ERROR: Output directory {0} could not be created. Error: {1}\r\n\r\nExiting ExcelDnaPack.", outputDirectory, ex.Message);
                    return 1;
                }
            }

            // Find ExcelDna.xll to use.
            // First try <MyAddin>.xll
            if (!File.Exists(xllInputPath))
            {
                // CONSIDER: Maybe the next two (old) search locations should be deprecated?
                // Then try one called ExcelDna.xll next to the .dna file
                xllInputPath = Path.Combine(dnaDirectory, HostXLL);
                if (!File.Exists(xllInputPath))
                {
                    // Then try one called ExcelDna.xll next to the ExcelDnaPack.exe
                    xllInputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, HostXLL);
                    if (!File.Exists(xllInputPath))
                    {
                        buildLogger.Error(typeof(ExcelDnaPack), "ERROR: Base add-in not found.\r\n\r\n {0}", usageInfo);
                        return 1;
                    }
                }
            }

            buildLogger.Information("Using base add-in " + xllInputPath);

            if (useManagedResourceResolver)
            {
                buildLogger.Information("Using managed resource packing.");
            }

            ResourceHelper.ResourceUpdater ru = null;
            if (filesToPublish == null)
            {
                File.Copy(xllInputPath, xllOutputPath, false);
                ru = new ResourceHelper.ResourceUpdater(Path.Combine(Directory.GetCurrentDirectory(), xllOutputPath), useManagedResourceResolver, buildLogger);
            }
            else
            {
                filesToPublish.Add(xllInputPath);
            }

            if (File.Exists(configPath))
            {
                if (filesToPublish == null)
                    ru.AddFile(File.ReadAllBytes(configPath), "__MAIN__", ResourceHelper.TypeName.CONFIG, null, false, multithreading);  // Name here must exactly match name in ExcelDnaLoad.cpp.
                else
                    filesToPublish.Add(configPath);
            }

            string[] dependenciesToExclude = (excludeDependencies ?? "").Split(';');
            if (packNativeLibraryDependencies && outputBitness != null && filesToPublish == null)
            {
                foreach (string nativeLibrary in FindNativeLibrariesDeps(dnaPath, outputBitness))
                {
                    if (dependenciesToExclude.Contains(Path.GetFileName(nativeLibrary), StringComparer.OrdinalIgnoreCase))
                        continue;

                    ru.AddFile(File.ReadAllBytes(nativeLibrary), Path.GetFileName(nativeLibrary).ToUpperInvariant(), ResourceHelper.TypeName.NATIVE_LIBRARY, "Native deps.json", compress, multithreading);
                }
            }

            if (!string.IsNullOrWhiteSpace(docPath) && filesToPublish == null)
            {
                ru.AddFile(File.ReadAllBytes(docPath), Path.GetFileName(docPath).ToUpperInvariant(), ResourceHelper.TypeName.DOC, null, compress, multithreading);
            }

            byte[] dnaBytes = File.ReadAllBytes(dnaPath);
            byte[] dnaContentForPacking = PackDnaLibrary(dnaPath, dnaBytes, dnaDirectory, ru, compress, multithreading, filesToPublish, packManagedDependencies, dependenciesToExclude, outputBitness, buildLogger);
            if (filesToPublish == null)
            {
                ru.AddFile(dnaContentForPacking, "__MAIN__", ResourceHelper.TypeName.DNA, null, false, multithreading); // Name here must exactly match name in DnaLibrary.Initialize.
                ru.EndUpdate();
            }
            else
            {
                filesToPublish.Add(dnaPath);
            }

            buildLogger.Information("Completed Packing {0}.", xllOutputPath);

            // All OK - set process exit code to 'Success'
            return 0;
        }

        public static int PackNativeAOT(string mainNativeAssembly, string xllOutput, bool multithreading, bool useManagedResourceResolver, IBuildLogger buildLogger)
        {
            ResourceHelper.ResourceUpdater ru = new ResourceHelper.ResourceUpdater(xllOutput, useManagedResourceResolver, buildLogger);
            ru.AddFile(File.ReadAllBytes(mainNativeAssembly), "__MAIN__", ResourceHelper.TypeName.NATIVE_ASSEMBLY, null, false, multithreading);  // Name here must exactly match name in host.cpp.
            ru.EndUpdate();

            // All OK - set process exit code to 'Success'
            return 0;
        }

        static private byte[] PackDnaLibrary(string dnaPath, byte[] dnaContent, string dnaDirectory, ResourceHelper.ResourceUpdater ru, bool compress, bool multithreading, List<string> filesToPublish, bool packManagedDependencies, string[] dependenciesToExcludeParam, string outputBitness, IBuildLogger buildLogger)
        {
            List<string> dependenciesToExclude = new List<string>();
            if (dependenciesToExcludeParam != null)
                dependenciesToExclude.AddRange(dependenciesToExcludeParam);

            string errorMessage;
            DnaLibrary dna = DnaLibrary.LoadFrom(dnaContent, dnaDirectory);
            if (dna == null)
            {
                // TODO: Better error handling here.
                errorMessage = "ERROR: .dna file could not be loaded. Possibly malformed xml content? ABORTING.";
                buildLogger.Error(typeof(ExcelDnaPack), errorMessage);
                throw new InvalidOperationException(errorMessage);
            }
            if (dna.ExternalLibraries != null)
            {
                bool copiedVersion = false;
                foreach (ExternalLibrary ext in dna.ExternalLibraries)
                {
                    var path = dna.ResolvePath(ext.Path);
                    if (!File.Exists(path))
                    {
                        var format = "!!! ERROR: ExternalLibrary `{0}` not found. ABORTING.";
                        errorMessage = string.Format(format, ext.Path);
                        buildLogger.Error(typeof(ExcelDnaPack), format, ext.Path);
                        throw new InvalidOperationException(errorMessage);
                    }

                    if (ext.Pack)
                    {
                        buildLogger.Information("  ~~> ExternalLibrary path {0} resolved to {1}.", ext.Path, path);
                        if (Path.GetExtension(path).Equals(".DNA", StringComparison.OrdinalIgnoreCase))
                        {
                            string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + ".DNA";
                            byte[] dnaContentForPacking = PackDnaLibrary(path, File.ReadAllBytes(path), Path.GetDirectoryName(path), ru, compress, multithreading, filesToPublish, false, null, outputBitness, buildLogger);
                            if (filesToPublish == null)
                            {
                                ru.AddFile(dnaContentForPacking, name, ResourceHelper.TypeName.DNA, null, compress, multithreading);
                                ext.Path = "packed:" + name;
                            }
                            else
                            {
                                filesToPublish.Add(path);
                            }
                        }
                        else
                        {
                            if (filesToPublish == null)
                            {
                                string packedName = ru.AddAssembly(path, null, compress, multithreading, ext.IncludePdb);
                                if (packedName != null)
                                {
                                    ext.Path = "packed:" + packedName;
                                }
                                dependenciesToExclude.Add(Path.GetFileName(path));
                            }
                            else
                            {
                                filesToPublish.Add(path);
                            }
                        }
                        if (ext.ComServer == true)
                        {
                            // Check for a TypeLib to pack
                            //string tlbPath = dna.ResolvePath(ext.TypeLibPath);
                            string resolvedTypeLibPath = null;
                            if (!string.IsNullOrEmpty(ext.TypeLibPath))
                            {
                                resolvedTypeLibPath = dna.ResolvePath(ext.TypeLibPath); // null is unresolved
                                if (resolvedTypeLibPath == null)
                                {
                                    // Try relative to .dll
                                    resolvedTypeLibPath = DnaLibrary.ResolvePath(ext.TypeLibPath, System.IO.Path.GetDirectoryName(path)); // null is unresolved
                                    if (resolvedTypeLibPath == null)
                                    {
                                        var format = "!!! ERROR: ExternalLibrary TypeLib path {0} could not be resolved.";
                                        errorMessage = string.Format(format, ext.TypeLibPath);
                                        buildLogger.Error(typeof(ExcelDnaPack), format, ext.TypeLibPath);
                                        throw new InvalidOperationException(errorMessage);
                                    }
                                }
                            }
                            else
                            {
                                // Check for .tlb
                                string tlbCheck = System.IO.Path.ChangeExtension(path, "tlb");
                                if (System.IO.File.Exists(tlbCheck))
                                {
                                    resolvedTypeLibPath = tlbCheck;
                                }
                            }
                            if (resolvedTypeLibPath != null)
                            {
                                if (filesToPublish == null)
                                {
                                    buildLogger.Information("  ~~> ExternalLibrary typelib path {0} resolved to {1}.", ext.TypeLibPath, resolvedTypeLibPath);
                                    int packedIndex = ru.AddTypeLib(File.ReadAllBytes(resolvedTypeLibPath));
                                    ext.TypeLibPath = "packed:" + packedIndex.ToString();
                                }
                                else
                                {
                                    filesToPublish.Add(resolvedTypeLibPath);
                                }
                            }
                        }
                    }
                    if (filesToPublish == null && ext.UseVersionAsOutputVersion)
                    {
                        if (copiedVersion)
                        {
                            buildLogger.Information("  ~~> Assembly version already copied from previous ExternalLibrary; ignoring 'UseVersionAsOutputVersion' attribute.");
                            continue;
                        }
                        try
                        {
                            ru.CopyFileVersion(Path.Combine(Directory.GetCurrentDirectory(), path));
                            copiedVersion = true;
                        }
                        catch (Exception e)
                        {
                            var format = "  ~~> ERROR: Error copying version to output version: {0}";
                            errorMessage = string.Format(format, e.Message);
                            buildLogger.Error(typeof(ExcelDnaPack), format, e.Message);
                            throw new InvalidOperationException(errorMessage);
                        }
                    }
                }
            }
            // Collect the list of all the references.
            List<Reference> refs = new List<Reference>();
            foreach (Project proj in dna.GetProjects())
            {
                if (proj.References != null)
                {
                    refs.AddRange(proj.References);
                }
            }
            // Fix-up if Reference is not part of a project, but just used to add an assembly for packing.
            foreach (Reference rf in dna.References)
            {
                if (!refs.Contains(rf))
                    refs.Add(rf);
            }

            // Expand asterisk in filename of reference path, e.g. "./*.dll"
            List<Reference> expandedReferences = new List<Reference>();
            for (int i = refs.Count - 1; i >= 0; i--)
            {
                string path = refs[i].Path;
                if (path != null && path.Contains("*"))
                {
                    var files = Directory.GetFiles(Path.GetDirectoryName(path), Path.GetFileName(path));
                    foreach (var file in files)
                    {
                        expandedReferences.Add(new Reference(Path.GetFullPath(file)) { Pack = true });
                    }
                    refs.RemoveAt(i);
                }
            }
            refs.AddRange(expandedReferences);

            // Now pack the references
            foreach (Reference rf in refs)
            {
                if (rf.Pack)
                {
                    string path = null;
                    if (rf.Path != null)
                    {
                        if (rf.Path.StartsWith("packed:"))
                        {
                            continue;
                        }

                        path = dna.ResolvePath(rf.Path);
                        buildLogger.Information("  ~~> Assembly path {0} resolved to {1}.", rf.Path, path);
                    }
                    if (path == null && rf.Name != null)
                    {
                        // Try Load as as last resort (and opportunity to load by FullName)
                        try
                        {
#pragma warning disable 0618
                            Assembly ass = Assembly.LoadWithPartialName(rf.Name);
#pragma warning restore 0618
                            if (ass != null)
                            {
                                path = ass.Location;
                                buildLogger.Information("  ~~> Assembly {0} 'Load'ed from location {1}.", rf.Name, path);
                            }
                        }
                        catch (Exception e)
                        {
                            buildLogger.Error(e, "  ~~> Assembly {0} not 'Load'ed. Exception: {1}", rf.Name, e);
                        }
                    }
                    if (path == null)
                    {
                        var format = "  ~~> ERROR: Reference with Path: {0} and Name: {1} NOT FOUND.";
                        errorMessage = string.Format(format, rf.Path, rf.Name);
                        buildLogger.Error(typeof(ExcelDnaPack), format, rf.Path, rf.Name);
                        throw new InvalidOperationException(errorMessage);
                    }

                    // It worked!
                    if (filesToPublish == null)
                    {
                        bool isNativeLibrary = false;
#if PACKEXTENDEDDEPS
                        isNativeLibrary = IsNativeLibrary(path);
#endif
                        if (isNativeLibrary)
                        {
                            ru.AddFile(File.ReadAllBytes(path), Path.GetFileName(path).ToUpperInvariant(), ResourceHelper.TypeName.NATIVE_LIBRARY, null, compress, multithreading);
                        }
                        else
                        {
                            string packedName = ru.AddAssembly(path, null, compress, multithreading, rf.IncludePdb);
                            if (packedName != null)
                            {
                                rf.Path = "packed:" + packedName;
                            }
                        }
                        dependenciesToExclude.Add(Path.GetFileName(path));
                    }
                    else
                    {
                        filesToPublish.Add(path);
                    }
                }
            }
            foreach (Image image in dna.Images)
            {
                if (image.Pack)
                {
                    string path = dna.ResolvePath(image.Path);
                    if (path == null)
                    {
                        var format = "  ~~> ERROR: Image path {0} NOT RESOLVED.";
                        errorMessage = string.Format(format, image.Path);
                        buildLogger.Error(typeof(ExcelDnaPack), format, image.Path);
                        throw new InvalidOperationException(errorMessage);
                    }
                    if (filesToPublish == null)
                    {
                        string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                        byte[] imageBytes = File.ReadAllBytes(path);
                        ru.AddFile(imageBytes, name, ResourceHelper.TypeName.IMAGE, null, compress, multithreading);
                        image.Path = "packed:" + name;
                    }
                    else
                    {
                        filesToPublish.Add(path);
                    }
                }
            }
            foreach (Project project in dna.Projects)
            {
                foreach (SourceItem source in project.SourceItems)
                {
                    if (source.Pack && !string.IsNullOrEmpty(source.Path))
                    {
                        string path = dna.ResolvePath(source.Path);
                        if (path == null)
                        {
                            var format = "  ~~> ERROR: Source path {0} NOT RESOLVED.";
                            errorMessage = string.Format(format, source.Path);
                            buildLogger.Error(typeof(ExcelDnaPack), format, source.Path);
                            throw new InvalidOperationException(errorMessage);
                        }
                        if (filesToPublish == null)
                        {
                            string name = Path.GetFileNameWithoutExtension(path).ToUpperInvariant() + "_" + lastPackIndex++ + Path.GetExtension(path).ToUpperInvariant();
                            byte[] sourceBytes = File.ReadAllBytes(path);
                            ru.AddFile(sourceBytes, name, ResourceHelper.TypeName.SOURCE, null, compress, multithreading);
                            source.Path = "packed:" + name;
                        }
                        else
                        {
                            filesToPublish.Add(path);
                        }
                    }
                }
            }

            if (packManagedDependencies && filesToPublish == null)
            {
                foreach (string assembly in FindManagedDeps(dnaPath, outputBitness))
                {
                    if (dependenciesToExclude.Contains(Path.GetFileName(assembly), StringComparer.OrdinalIgnoreCase))
                        continue;

                    ru.AddAssembly(assembly, "Managed deps.json", compress, multithreading, false);
                }
            }

            if (!string.IsNullOrWhiteSpace(dna.CustomRuntimeConfiguration))
            {
                string path = dna.ResolvePath(dna.CustomRuntimeConfiguration);
                if (path == null)
                {
                    var format = "  ~~> ERROR: Custom runtime configuration path {0} NOT RESOLVED.";
                    errorMessage = string.Format(format, dna.CustomRuntimeConfiguration);
                    buildLogger.Error(typeof(ExcelDnaPack), format, dna.CustomRuntimeConfiguration);
                    throw new InvalidOperationException(errorMessage);
                }

                if (filesToPublish == null)
                {
                    ru.AddFile(File.ReadAllBytes(path), "__CUSTOM_RUNTIMECONFIG__", ResourceHelper.TypeName.SOURCE, null, false, multithreading);
                }
                else
                {
                    filesToPublish.Add(path);
                }
            }

            return DnaLibrary.Save(dna);
        }

        static private List<string> FindNativeLibrariesDeps(string dnaPath, string outputBitness)
        {
            List<string> result = new List<string>();
#if PACKEXTENDEDDEPS
            string basePath = Path.GetDirectoryName(dnaPath);
            foreach (Microsoft.Extensions.DependencyModel.RuntimeAssetGroup asset in FindDepsAssets(dnaPath))
            {
                if (!MatchArchitecture(asset.Runtime, outputBitness))
                    continue;

                foreach (var path in asset.AssetPaths)
                {
                    string fullPath = Path.Combine(basePath, path);
                    if (File.Exists(fullPath) && IsNativeLibrary(fullPath))
                        result.Add(fullPath);
                }
            }
#endif
            return result;
        }

        static private List<string> FindManagedDeps(string dnaPath, string outputBitness)
        {
            List<string> result = new List<string>();
#if PACKEXTENDEDDEPS
            string basePath = Path.GetDirectoryName(dnaPath);

            var assets = FindDepsAssets(dnaPath);
            assets.RemoveAll(i => i.Runtime.Length > 0 && !MatchArchitecture(i.Runtime, outputBitness));

            foreach (Microsoft.Extensions.DependencyModel.RuntimeAssetGroup asset in assets)
            {
                foreach (var path in asset.AssetPaths)
                {
                    if (asset.Runtime.Length == 0 && FindAssetHavingRuntime(assets, Path.GetFileName(path)))
                        continue;

                    string fullPath = Path.Combine(basePath, path);
                    if (!File.Exists(fullPath))
                        fullPath = Path.Combine(basePath, Path.GetFileName(path));

                    if (File.Exists(fullPath) && IsManagedAssembly(fullPath))
                        result.Add(fullPath);
                }
            }
#endif
            return result;
        }

#if PACKEXTENDEDDEPS
        static private bool FindAssetHavingRuntime(List<Microsoft.Extensions.DependencyModel.RuntimeAssetGroup> assets, string fileName)
        {
            foreach (Microsoft.Extensions.DependencyModel.RuntimeAssetGroup asset in assets)
            {
                if (asset.Runtime.Length == 0)
                    continue;

                foreach (var path in asset.AssetPaths)
                {
                    if (string.Compare(fileName, Path.GetFileName(path), StringComparison.OrdinalIgnoreCase) == 0)
                        return true;
                }
            }

            return false;
        }

        static private List<Microsoft.Extensions.DependencyModel.RuntimeAssetGroup> FindDepsAssets(string dnaPath)
        {
            var result = new List<Microsoft.Extensions.DependencyModel.RuntimeAssetGroup>();
            string depsJsonPath = Path.ChangeExtension(dnaPath, "deps.json");
            if (!File.Exists(depsJsonPath))
                return result;

            var reader = new Microsoft.Extensions.DependencyModel.DependencyContextJsonReader();
            using (FileStream depsStream = new FileStream(depsJsonPath, FileMode.Open))
            {
                Microsoft.Extensions.DependencyModel.DependencyContext dc = reader.Read(depsStream);
                foreach (Microsoft.Extensions.DependencyModel.RuntimeLibrary library in dc.RuntimeLibraries)
                {
                    foreach (Microsoft.Extensions.DependencyModel.RuntimeAssetGroup asset in library.RuntimeAssemblyGroups.Concat(library.NativeLibraryGroups))
                    {
                        result.Add(asset);
                    }
                }
            }

            return result;
        }

        static private bool MatchArchitecture(string runtimeID, string requiredBitness)
        {
            if (!runtimeID.StartsWith("win"))
                return false;

            string matchingArchitecture = "-x" + (requiredBitness == "64" ? "64" : "86");
            string mismatchingArchitecture = "-x" + (requiredBitness == "64" ? "86" : "64");
            if (runtimeID.Contains(matchingArchitecture))
                return true;
            if (runtimeID.Contains(mismatchingArchitecture))
                return false;
            if (runtimeID.Contains("-arm"))
                return false;

            return true; // runtimeID doesn't specify specific architecture
        }

        static private bool IsNativeLibrary(string path)
        {
            bool isPE;
            if (IsAssembly(path, out isPE))
                return false;

            return isPE;
        }

        static private bool IsManagedAssembly(string path)
        {
            bool isPE;
            return IsAssembly(path, out isPE);
        }

        static private bool IsAssembly(string path, out bool isPE)
        {
            isPE = false;
            using (FileStream file = File.OpenRead(path))
            {
                try
                {
                    var peReader = new System.Reflection.PortableExecutable.PEReader(file);
                    System.Reflection.PortableExecutable.CorHeader corHeader = peReader.PEHeaders.CorHeader;

                    isPE = true; // If peReader.PEHeaders doesn't throw, it is a valid PEImage
                    return corHeader != null;
                }
                catch (BadImageFormatException)
                {
                }
            }
            return false;
        }
#endif

        static private int lastPackIndex = 0;
    }
}
