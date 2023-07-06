using OfficeOpenXml;
using SheetToObjects.Lib;
using SheetToObjects.Lib.FluentConfiguration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel;

public class MicrosoftExcelDataReader<T> : IReadOnlyCollection<ReaderObject<T>>, IEnumerator<ReaderObject<T>>, IEnumerator
    where T : new()
{
    private int _rowNumber = -1;
    private ReaderObject<T>? _current;
    private bool _disposedValue;

    private readonly IConvertDataToRow<List<string>> _excelRowConverter;
    private readonly ExcelPackage _excelPackage;
    private readonly ExcelWorksheet _workSheet;
    private readonly ExcelRange _range;
    private readonly bool _stopReadingOnEmptyRow;
    private readonly IMapSheetToObjects _mapper;

    static MicrosoftExcelDataReader()
    {
        FieldInfo licenseField = typeof(ExcelPackage).GetField("_licenseSet", BindingFlags.NonPublic | BindingFlags.Static);
        licenseField.SetValue(null, true);
    }

    object IEnumerator.Current => ((IEnumerator<ReaderObject<T>>)this).Current;

    public int Count => _workSheet.Rows.Count();
    public ExcelPackage ExcelPackage => _excelPackage;

    ReaderObject<T> IEnumerator<ReaderObject<T>>.Current => GetCurrent();

    internal MicrosoftExcelDataReader(
        ExcelPackage excelPackage,
        string sheetName,
        ExcelRange range,
        Func<MappingConfigBuilder<T>, MappingConfigBuilder<T>> mappingConfigFunc,
        IConvertDataToRow<List<string>> excelRowConverter,
        bool stopReadingOnEmptyRow)
    {
        _excelPackage = excelPackage;
        _workSheet = GetSheetFromWorkBook(excelPackage.Workbook, sheetName);
        _excelRowConverter = excelRowConverter;
        _range = range;
        _stopReadingOnEmptyRow = stopReadingOnEmptyRow;

        _mapper = new SheetMapper().AddConfigFor(mappingConfigFunc);
    }

    public static MicrosoftExcelDataReader<T> CreateReader(
        Stream stream,
        string sheetName,
        ExcelRange range,
        Func<MappingConfigBuilder<T>, MappingConfigBuilder<T>> mappingConfigFunc,
        bool stopReadingOnEmptyRow = false)
    {
        MicrosoftExcelToRowConverter rowConverter = new();
        ExcelPackage excelPackage = new(stream);
        ExcelWorkbook workBook = excelPackage.Workbook;
        ExcelWorksheet workSheet = GetSheetFromWorkBook(workBook, sheetName);

        return new MicrosoftExcelDataReader<T>(excelPackage, sheetName, range, mappingConfigFunc, rowConverter, stopReadingOnEmptyRow);
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

    public IEnumerator<ReaderObject<T>> GetEnumerator()
    {
        return this;
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this;
    }

    public bool MoveNext()
    {
        _rowNumber++;

        if (_rowNumber >= _range.To.RowNumber)
            return false;

        if (!_stopReadingOnEmptyRow)
            return true;

        for (var columnNumber = _range.From.ColumnNumber; columnNumber <= _range.To.ColumnNumber; columnNumber++)
        {
            string item = _workSheet.Cells[_rowNumber, columnNumber].Text;
            if (!string.IsNullOrEmpty(item))
                return true;
        }

        return false;
    }

    public void Reset()
    {
        _rowNumber = -1;
        _current = null;
    }

    private ReaderObject<T> GetCurrent()
    {
        if (_current is not null)
            return _current;

        List<string> rowData = new();

        for (var columnNumber = _range.From.ColumnNumber; columnNumber <= _range.To.ColumnNumber; columnNumber++)
        {
            rowData.Add(_workSheet.Cells[_rowNumber, columnNumber].Text);
        }

        Row row = _excelRowConverter.ConvertRow(rowData, _rowNumber);
        MappingRowResult<T> mappingRowResult = _mapper.MapRow<T>(row);

        _current = new ReaderObject<T>()
        {
            Row = _workSheet.Row(_rowNumber),
            MappingRowResult = mappingRowResult
        };

        return _current;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                _excelPackage?.Dispose();
                // TODO: освободить управляемое состояние (управляемые объекты)
            }

            // TODO: освободить неуправляемые ресурсы (неуправляемые объекты) и переопределить метод завершения
            // TODO: установить значение NULL для больших полей
            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        // Не изменяйте этот код. Разместите код очистки в методе "Dispose(bool disposing)".
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
