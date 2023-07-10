using Microsoft.Extensions.Configuration;
using SheetToObjects.Configuration.Settings.Configuration.Columns;

namespace SheetToObjects.Configuration.Settings.Configuration;
public abstract class ConfigurationReader<T>
{
    protected readonly IConfiguration _configSection;
    protected readonly ConfigurationReaderOptions _readerOptions;
    protected readonly IConfiguration? _configuration;
    protected readonly MappingConfigBuilderConfigurator<T> _configurator;

    public ConfigurationReader(IConfiguration configSection, ConfigurationReaderOptions readerOptions, IConfiguration? configuration = null)//AssemblyFinder assemblyFinder, 
    {
        _configSection = configSection;
        _readerOptions = readerOptions;
        _configuration = configuration;

        _configurator = new MappingConfigBuilderConfigurator<T>();
    }

    public virtual void Configure(SheetMapConfig sheetMapConfig)
    {
        ApplyColumnConfigs();
        //ApplyRangeConfig();
        ApplySkipOnEmptryRowConfig();

        sheetMapConfig.MappingConfig = _configurator.ConfigBuilder.Build();
    }

    private void ApplySkipOnEmptryRowConfig()
    {
        bool skip = _configSection.GetSection(nameof(MappingConfiguration.SkipOnEmptyRow)).Get<bool>();
        _configurator.SkipOnEmptryRow(skip);
    }

    private void ApplyColumnConfigs()
    {
        IConfigurationSection columnSection = _configSection.GetSection(ColumnMappingConfig.SectionName);
        List<ColumnMappingConfig> configs = columnSection.GetRequired<List<ColumnMappingConfig>>();

        foreach (ColumnMappingConfig config in configs)
        {
            _configurator.ConfigureColumn(config);
        }
    }
}
