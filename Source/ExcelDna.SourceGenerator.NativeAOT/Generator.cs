﻿using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelDna.SourceGenerator.NativeAOT
{
    [Generator]
    public class Generator : ISourceGenerator
    {
        public void Execute(GeneratorExecutionContext context)
        {
            if (!(context.SyntaxContextReceiver is SyntaxReceiver receiver))
                return;

            context.AnalyzerConfigOptions.GlobalOptions.TryGetValue("build_property.PublishAOT", out string? publishAOT);

            string source = """
// <auto-generated/>
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ExcelDna.SourceGenerator.NativeAOT
{
    public unsafe class AddInInitialize
    {
        [UnmanagedCallersOnly(EntryPoint = "Initialize", CallConvs = new[] { typeof(CallConvCdecl) })]
        public static short Initialize(void* xlAddInExportInfoAddress, void* hModuleXll, void* pPathXLL, byte disableAssemblyContextUnload, void* pTempDirPath)
        {
            ExcelDna.Integration.NativeAOT.IsActive = true;

            [FUNCTIONS]

            return ExcelDna.ManagedHost.AddInInitialize.InitializeNativeAOT(xlAddInExportInfoAddress, hModuleXll, pPathXLL, disableAssemblyContextUnload, pTempDirPath);
        }
    }
}
""";
            string functions = "";
            foreach (var i in receiver.Functions)
            {
                functions += $"ExcelDna.Integration.NativeAOT.MethodsForRegistration.Add(typeof({i.ContainingType}).GetMethod(\"{i.Name}\")!);\r\n";
            }

            source = source.Replace("[FUNCTIONS]", functions);

            context.AddSource($"ExcelDna.SG.NAOT.Init.g.cs", source);
        }

        public void Initialize(GeneratorInitializationContext context)
        {
            context.RegisterForSyntaxNotifications(() => new SyntaxReceiver());
        }

        class SyntaxReceiver : ISyntaxContextReceiver
        {
            public List<IMethodSymbol> Functions { get; } = new List<IMethodSymbol>();

            public void OnVisitSyntaxNode(GeneratorSyntaxContext context)
            {
                if (context.Node is MethodDeclarationSyntax methodSyntax)
                {
                    IMethodSymbol methodSymbol = (context.SemanticModel.GetDeclaredSymbol(methodSyntax) as IMethodSymbol)!;
                    if (methodSymbol.GetAttributes().Any(a => a.AttributeClass?.ToDisplayString(fullNameFormat) == "ExcelDna.Integration.ExcelFunctionAttribute"))
                        Functions.Add(methodSymbol);
                }
            }

            private static SymbolDisplayFormat fullNameFormat = new SymbolDisplayFormat(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces);
        }
    }
}
