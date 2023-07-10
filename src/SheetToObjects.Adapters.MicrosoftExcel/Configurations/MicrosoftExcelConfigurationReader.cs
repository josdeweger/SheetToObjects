using Microsoft.Extensions.Configuration;
using SheetToObjects.Configuration.Settings.Configuration;
using SheetToObjects.Configuration.Settings.Configuration.Ranges;
using System;

namespace SheetToObjects.Adapters.MicrosoftExcel.Configurations;
internal class MicrosoftExcelConfigurationReader<T> : ConfigurationReader<T>
{
    public MicrosoftExcelConfigurationReader(IConfiguration configSection, ConfigurationReaderOptions readerOptions, IConfiguration? configuration = null)
        : base(configSection, readerOptions, configuration)
    {
    }

    public void Configure(MicrosoftExcelSheetMapConfig sheetMapConfig)
    {
        base.Configure(sheetMapConfig);
        //ApplyRange(sheetMapConfig);
        ApplyRangeConfig(sheetMapConfig);
    }

    private void ApplyRange(MicrosoftExcelSheetMapConfig sheetMapConfig)
    {
        ExcelRange excelRange = _configSection.GetSection(nameof(MicrosoftExcelSheetMapConfig.Range)).GetRequired<ExcelRange>();
        sheetMapConfig.Range = excelRange;
    }

    private void ApplyRangeConfig(MicrosoftExcelSheetMapConfig sheetMapConfig)
    {
        IConfigurationSection rangeSection = _configSection.GetSection(RangeMappingConfig.SectionName);
        RangeMappingConfig rangeMappingConfig = rangeSection.GetRequired<RangeMappingConfig>();

        if (rangeMappingConfig.HeaderRow is null)
            return;

        //ExcelRange excelRange = new()
        int startRow = 1;
        if (rangeMappingConfig.HeaderRow is not null)
        {
            startRow = rangeMappingConfig.HeaderRow.Value;
        }

        ExcelCell from = new(rangeMappingConfig.FromColumn, startRow);
        ExcelCell to = new(rangeMappingConfig.ToColumn, Int32.MaxValue);
        ExcelRange excelRange = new(from, to);

        sheetMapConfig.Range = excelRange;
    }
}
