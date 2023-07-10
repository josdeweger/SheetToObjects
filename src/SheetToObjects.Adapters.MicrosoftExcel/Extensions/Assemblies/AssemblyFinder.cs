using Microsoft.Extensions.DependencyModel;
using SheetToObjects.Configuration.Settings.Configuration;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Extensions.Assemblies;

abstract class AssemblyFinder
{
    public abstract IReadOnlyList<AssemblyName> FindAssembliesContainingName(string nameToFind);

    protected static bool IsCaseInsensitiveMatch(string? text, string textToFind)
    {
        return text != null && text.ToLowerInvariant().Contains(textToFind.ToLowerInvariant());
    }

    public static AssemblyFinder Auto()
    {
        return new CompositeAssemblyFinder(new DependencyContextAssemblyFinder(DependencyContext.Default), new DllScanningAssemblyFinder());
    }

    public static AssemblyFinder ForSource(ConfigurationAssemblySource configurationAssemblySource)
    {
        return configurationAssemblySource switch
        {
            ConfigurationAssemblySource.UseLoadedAssemblies => Auto(),
            ConfigurationAssemblySource.AlwaysScanDllFiles => new DllScanningAssemblyFinder(),
            _ => throw new ArgumentOutOfRangeException(nameof(configurationAssemblySource), configurationAssemblySource, null),
        };
    }

    public static AssemblyFinder ForDependencyContext(DependencyContext dependencyContext)
    {
        return new DependencyContextAssemblyFinder(dependencyContext);
    }

    protected internal abstract IEnumerable<Assembly> GetAssemblies();
}
