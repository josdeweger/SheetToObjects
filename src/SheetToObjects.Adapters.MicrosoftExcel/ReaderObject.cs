using OfficeOpenXml;

namespace SheetToObjects.Adapters.MicrosoftExcel;

public class ReaderObject<T>
    where T : new()
{
    public required ExcelRow Row { get; init; }
    public required MappingRowResult<T> MappingRowResult { get; init; }
}
