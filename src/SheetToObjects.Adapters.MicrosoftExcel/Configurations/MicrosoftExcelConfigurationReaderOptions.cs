using Microsoft.Extensions.DependencyModel;
using SheetToObjects.Configuration.Settings.Configuration;
using System.Collections.Generic;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Configurations;
public class MicrosoftExcelConfigurationReaderOptions : ConfigurationReaderOptions
{
    private const string _defaultSectionPath = "SheetToObjects:MicrosoftExcel";
    private string _sectionPath = _defaultSectionPath;

    public MicrosoftExcelConfigurationReaderOptions()
        : base()
    {
    }
    public MicrosoftExcelConfigurationReaderOptions(IEnumerable<Assembly> assemblies)
        : base(assemblies)
    {
    }

    public MicrosoftExcelConfigurationReaderOptions(DependencyContext? dependencyContext) : base(dependencyContext)
    {
    }

    public MicrosoftExcelConfigurationReaderOptions(ConfigurationAssemblySource source) : base(source)
    {
    }

    public override string SectionPath { get => _sectionPath; set => _sectionPath = value; }
}
