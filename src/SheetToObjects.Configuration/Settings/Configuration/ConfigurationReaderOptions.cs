using Microsoft.Extensions.DependencyModel;
using System.Reflection;

namespace SheetToObjects.Configuration.Settings.Configuration;

public abstract class ConfigurationReaderOptions
{
    private readonly Assembly[]? _assemblies;

    public string SectionName { get; init; } = "SheetToObjects:MicrosoftExcel";

    protected ConfigurationReaderOptions()
    {

    }
    protected ConfigurationReaderOptions(IEnumerable<Assembly> assemblies)
    {
        _assemblies = assemblies as Assembly[] ?? assemblies?.ToArray();
    }
    public Assembly[]? Assemblies => _assemblies;
    public DependencyContext? DependencyContext { get; }
    public ConfigurationAssemblySource? ConfigurationAssemblySource { get; }
}


