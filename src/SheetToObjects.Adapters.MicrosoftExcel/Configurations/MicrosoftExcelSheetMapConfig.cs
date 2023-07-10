using SheetToObjects.Configuration.Settings.Configuration;

namespace SheetToObjects.Adapters.MicrosoftExcel.Configurations;
public class MicrosoftExcelSheetMapConfig : SheetMapConfig
{
    public ExcelRange Range { get; set; }
}
