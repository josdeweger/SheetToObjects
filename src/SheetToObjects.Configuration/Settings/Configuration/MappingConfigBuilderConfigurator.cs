using SheetToObjects.Configuration.Settings.Configuration.Columns;
using SheetToObjects.Configuration.Settings.Configuration.Ranges;
using SheetToObjects.Lib.FluentConfiguration;

namespace SheetToObjects.Configuration.Settings.Configuration;

public class MappingConfigBuilderConfigurator<T>
{
    private readonly MappingConfigBuilder<T> _configBuilder;
    private readonly ColumnMappingBuilderConfigure<T> _columnConfigure;

    internal MappingConfigBuilderConfigurator()
    {
        _configBuilder = new MappingConfigBuilder<T>();
        _columnConfigure = new ColumnMappingBuilderConfigure<T>(_configBuilder);
    }

    public MappingConfigBuilder<T> ConfigBuilder => _configBuilder;

    public void ConfigureColumn(ColumnMappingConfig columnMappingConfig)
    {
        _columnConfigure.Configure(columnMappingConfig);
    }

    public void SkipOnEmptryRow(bool skipOnEmptryRow)
    {
        if (!skipOnEmptryRow)
            return;

        _configBuilder.StopParsingAtFirstEmptyRow();
    }

    public void ApplyRange(RangeMappingConfig rangeConfig)
    {

    }
}


