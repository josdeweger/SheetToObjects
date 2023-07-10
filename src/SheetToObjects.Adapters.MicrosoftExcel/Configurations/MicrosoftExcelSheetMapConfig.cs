using SheetToObjects.Configuration.Settings.Configuration;

namespace SheetToObjects.Adapters.MicrosoftExcel.Configurations;
public class MicrosoftExcelSheetMapConfig : SheetMapConfig
{
    //public MappingConfig MappingConfig { get; set; }
    public ExcelRange Range { get; set; }
}
