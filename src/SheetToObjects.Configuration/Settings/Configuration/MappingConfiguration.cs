using SheetToObjects.Configuration.Settings.Configuration.Columns;
using SheetToObjects.Configuration.Settings.Configuration.Ranges;

namespace SheetToObjects.Configuration.Settings.Configuration;
public class MappingConfiguration
{
    public required RangeMappingConfig RangeConfiguration { get; init; }
    public required List<ColumnMappingConfig> ColumnConfigurations { get; init; }
    public bool SkipOnEmptyRow { get; init; } = true;
}
