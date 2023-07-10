namespace SheetToObjects.Configuration.Settings.Configuration.Ranges;
public class RangeMappingConfig
{
    public const string SectionName = "Range";

    public int? HeaderRow { get; init; }
    public string FromColumn { get; init; } = "A";
    public string ToColumn { get; init; } = "Z";
}
