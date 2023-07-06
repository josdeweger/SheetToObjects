using OfficeOpenXml;
using SheetToObjects.Lib;
using SheetToObjects.Lib.FluentConfiguration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel;
internal class MicrosoftExcelDataReaderOld<T> : IDisposable
    where T : new()
{
    private bool _disposedValue;
    private readonly IConvertDataToRow<List<string>> _excelRowConverter;
    private readonly Func<MappingConfigBuilder<T>, MappingConfigBuilder<T>> _mappingConfigFunc;

    static MicrosoftExcelDataReaderOld()
    {
        FieldInfo licenseField = typeof(ExcelPackage).GetField("_licenseSet", BindingFlags.NonPublic | BindingFlags.Static);
        licenseField.SetValue(null, true);
    }

    private MicrosoftExcelDataReaderOld(
        IConvertDataToRow<List<string>> excelRowConverter,
        Func<MappingConfigBuilder<T>, MappingConfigBuilder<T>> mappingConfigFunc)
    {
        _excelRowConverter = excelRowConverter;
        _mappingConfigFunc = mappingConfigFunc;
    }

    public static MicrosoftExcelDataReaderOld<T> CreateReader(Func<MappingConfigBuilder<T>, MappingConfigBuilder<T>> mappingConfigFunc)
    {
        MicrosoftExcelToRowConverter rowConverter = new();
        return new MicrosoftExcelDataReaderOld<T>(rowConverter, mappingConfigFunc);
    }

    public IEnumerable<MappingRowResult<T>> GetFromStream(Stream fileStream, string sheetName, ExcelRange range, bool stopReadingOnEmptyRow = false)
    {
        using var excelPackage = new ExcelPackage(fileStream);
        ExcelWorkbook workBook = excelPackage.Workbook;
        ExcelWorksheet workSheet = GetSheetFromWorkBook(workBook, sheetName);

        IMapSheetToObjects mapper = new SheetMapper().AddConfigFor(_mappingConfigFunc);

        for (var rowNumber = range.From.RowNumber; rowNumber <= range.To.RowNumber; rowNumber++)
        {
            List<string> rowData = new();

            for (var columnNumber = range.From.ColumnNumber; columnNumber <= range.To.ColumnNumber; columnNumber++)
            {
                rowData.Add(workSheet.Cells[rowNumber, columnNumber].Text);

                if (stopReadingOnEmptyRow && rowData.All(x => string.IsNullOrEmpty(x)))
                    yield break;
            }

            Row row = _excelRowConverter.ConvertRow(rowData, rowNumber);

            MappingRowResult<T> mappingRowResult = mapper.MapRow<T>(row);
            yield return mappingRowResult;
        }
    }

    private static ExcelWorksheet GetSheetFromWorkBook(ExcelWorkbook excelWorkbook, string sheetName)
    {
        var normalizedSheetName = sheetName.Replace(" ", "").ToLowerInvariant();

        for (var i = 1; i <= excelWorkbook.Worksheets.Count; i++)
        {
            if (excelWorkbook.Worksheets[i].Name.Replace(" ", "").ToLowerInvariant().Equals(normalizedSheetName))
            {
                return excelWorkbook.Worksheets[i];
            }
        }

        throw new ArgumentException($"Workbook does not contain Worksheet with name {sheetName}");
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {

                // TODO: освободить управляемое состояние (управляемые объекты)
            }

            // TODO: освободить неуправляемые ресурсы (неуправляемые объекты) и переопределить метод завершения
            // TODO: установить значение NULL для больших полей
            _disposedValue = true;
        }
    }

    // // TODO: переопределить метод завершения, только если "Dispose(bool disposing)" содержит код для освобождения неуправляемых ресурсов
    // ~MicrosoftExcelDataReader()
    // {
    //     // Не изменяйте этот код. Разместите код очистки в методе "Dispose(bool disposing)".
    //     Dispose(disposing: false);
    // }

    public void Dispose()
    {
        // Не изменяйте этот код. Разместите код очистки в методе "Dispose(bool disposing)".
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
