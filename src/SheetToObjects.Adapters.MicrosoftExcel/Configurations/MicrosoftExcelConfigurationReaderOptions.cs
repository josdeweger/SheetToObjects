using SheetToObjects.Configuration.Settings.Configuration;
using System.Collections.Generic;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Configurations;
public class MicrosoftExcelConfigurationReaderOptions : ConfigurationReaderOptions
{
    public MicrosoftExcelConfigurationReaderOptions()
        : base()
    {
        SectionName = "SheetToObjects:MicrosoftExcel";
    }
    public MicrosoftExcelConfigurationReaderOptions(IEnumerable<Assembly> assemblies)
        : base(assemblies)
    {
        SectionName = "SheetToObjects:MicrosoftExcel";
    }
}
