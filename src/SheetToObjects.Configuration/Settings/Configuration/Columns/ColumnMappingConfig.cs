namespace SheetToObjects.Configuration.Settings.Configuration.Columns;

public class ColumnMappingConfig
{
    public const string SectionName = "MapColumns";

    public string? WithHeader { get; init; }
    public int? WithColumnIndex { get; init; }
    public string? WithColumnLetter { get; init; }
    public string? UsingFormat { get; init; }
    public required string MapTo { get; init; }
    public bool IsRequired { get; init; }
}