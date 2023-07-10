using Microsoft.Extensions.DependencyModel;
using System.Reflection;

namespace SheetToObjects.Configuration.Settings.Configuration;

public abstract class ConfigurationReaderOptions
{
    private Assembly[]? _assemblies;
    private readonly DependencyContext? _dependencyContext;
    private readonly ConfigurationAssemblySource _source;

    public abstract string SectionPath { get; set; }

    protected ConfigurationReaderOptions(IEnumerable<Assembly> assemblies)
    {
        _assemblies = assemblies as Assembly[] ?? assemblies?.ToArray();
    }
    protected ConfigurationReaderOptions() : this(dependencyContext: null) { }

    public ConfigurationReaderOptions(DependencyContext? dependencyContext)
    {
        _dependencyContext = dependencyContext;
    }
    public ConfigurationReaderOptions(ConfigurationAssemblySource source)
    {
        _source = source;
    }

    //public Assembly[]? Assemblies => _assemblies;
    public Assembly[]? Assemblies { get => _assemblies; set => _assemblies = value; }
    public DependencyContext? DependencyContext => _dependencyContext;
    public ConfigurationAssemblySource? ConfigurationAssemblySource => _source;

}


