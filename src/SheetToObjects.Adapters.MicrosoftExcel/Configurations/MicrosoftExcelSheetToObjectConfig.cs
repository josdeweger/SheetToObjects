using SheetToObjects.Lib.FluentConfiguration;

namespace SheetToObjects.Adapters.MicrosoftExcel.Configurations;
public class MicrosoftExcelSheetToObjectConfig : SheetToObjectConfig
{
    public ExcelRange Range { get; private set; }
    public string? SheetName { get; private set; }
    public int? SheetIndex { get; private set; }

    public void ConfigureRange(ExcelRange range)
    {
        Range = range;
    }

    public void ConfigureSheetName(string sheetName)
    {
        SheetName = sheetName;
    }

    public void ConfigureSheetIndex(int sheetIndex)
    {
        SheetIndex = sheetIndex;
    }
}
