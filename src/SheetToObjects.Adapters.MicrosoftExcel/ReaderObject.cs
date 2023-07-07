using OfficeOpenXml;
using SheetToObjects.Lib;

namespace SheetToObjects.Adapters.MicrosoftExcel;

public class ReaderObject<T>
    where T : new()
{
    public required Row Row { get; init; }
    public required ExcelRow ExcelRow { get; init; }
    public required MappingRowResult<T> MappingRowResult { get; init; }
}
