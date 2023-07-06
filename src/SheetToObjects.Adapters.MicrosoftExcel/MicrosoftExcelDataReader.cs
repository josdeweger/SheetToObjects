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

public class MicrosoftExcelDataReader<T> : IReadOnlyCollection<ReaderObject<T>>, IDisposable
    where T : new()
{
    private int _rowNumber;
    private bool _disposedValue;

    private readonly IConvertDataToRow<List<string>> _excelRowConverter;
    private readonly ExcelPackage _excelPackage;
    private readonly ExcelWorksheet _workSheet;
    private readonly ExcelRange _range;
    private readonly IMapSheetToObjects _mapper;
    private readonly Enumerator _enumerator;

    static MicrosoftExcelDataReader()
    {
        FieldInfo licenseField = typeof(ExcelPackage).GetField("_licenseSet", BindingFlags.NonPublic | BindingFlags.Static);
        licenseField.SetValue(null, true);
    }

    public int Count => _workSheet.Rows.Count();
    public ExcelPackage ExcelPackage => _excelPackage;
    public ExcelWorksheet Worksheet => _workSheet;

    internal MicrosoftExcelDataReader(
        ExcelPackage excelPackage,
        string sheetName,
        ExcelRange range,
        IMapSheetToObjects mapper,
        IConvertDataToRow<List<string>> excelRowConverter,
        bool stopReadingOnEmptyRow)
    {
        _excelPackage = excelPackage;
        _workSheet = GetSheetFromWorkBook(excelPackage.Workbook, sheetName);
        _excelRowConverter = excelRowConverter;
        _range = range;

        _mapper = mapper;

        _rowNumber = _range.From.RowNumber - 1;
        _enumerator = new(_workSheet, range, excelRowConverter, mapper, stopReadingOnEmptyRow);
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
        IMapSheetToObjects mapper = new SheetMapper().AddConfigFor(mappingConfigFunc);

        return new MicrosoftExcelDataReader<T>(excelPackage, sheetName, range, mapper, rowConverter, stopReadingOnEmptyRow);
    }

    public static MicrosoftExcelDataReader<T> CreateReader(
        Stream stream,
        string sheetName,
        ExcelRange range,
        SheetToObjectConfig sheetToObjectConfig,
        bool stopReadingOnEmptyRow = false)
    {
        MicrosoftExcelToRowConverter rowConverter = new();
        ExcelPackage excelPackage = new(stream);
        IMapSheetToObjects mapper = new SheetMapper().AddSheetToObjectConfig(sheetToObjectConfig);

        return new MicrosoftExcelDataReader<T>(excelPackage, sheetName, range, mapper, rowConverter, stopReadingOnEmptyRow);
    }

    private static ExcelWorksheet GetSheetFromWorkBook(ExcelWorkbook excelWorkbook, string sheetName)
    {
        var normalizedSheetName = sheetName.Replace(" ", "").ToLowerInvariant();

        for (var i = 0; i < excelWorkbook.Worksheets.Count; i++)
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
        List<string> rowData = new();

        for (var columnNumber = _range.From.ColumnNumber; columnNumber <= _range.To.ColumnNumber; columnNumber++)
        {
            var excelRange = _workSheet.Cells[_rowNumber, columnNumber];
            rowData.Add(excelRange.Text);
        }

        Row row = _excelRowConverter.ConvertRow(rowData, _rowNumber);
        _mapper.MapHeadersToIndex<T>(row);

        return _enumerator;
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable<ReaderObject<T>>)this).GetEnumerator();
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

    internal struct Enumerator : IEnumerator<ReaderObject<T>>, IEnumerator
    {
        private readonly ExcelWorksheet _workSheet;
        private readonly ExcelRange _range;
        private readonly bool _stopReadingOnEmptyRow;
        private ReaderObject<T>? _current;
        private readonly IConvertDataToRow<List<string>> _excelRowConverter;
        private readonly IMapSheetToObjects _mapper;
        private int _rowNumber;

        public Enumerator(ExcelWorksheet workSheet, ExcelRange range, IConvertDataToRow<List<string>> excelRowConverter, IMapSheetToObjects mapper, bool stopReadingOnEmptyRow)
        {
            _workSheet = workSheet;
            _range = range;
            _excelRowConverter = excelRowConverter;
            _mapper = mapper;
            _stopReadingOnEmptyRow = stopReadingOnEmptyRow;

            _rowNumber = _range.From.RowNumber - 1;
        }

        public ReaderObject<T> Current => GetCurrent();

        object IEnumerator.Current => ((IEnumerator<ReaderObject<T>>)this).Current;

        public void Dispose()
        {
            _rowNumber = -1;
            _current = null;
        }

        public bool MoveNext()
        {
            _current = null;
            _rowNumber++;

            if (_rowNumber > _range.To.RowNumber)
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

        private ReaderObject<T> GetCurrent()
        {
            if (_current is not null)
                return _current;

            List<string> rowData = new();

            for (var columnNumber = _range.From.ColumnNumber; columnNumber <= _range.To.ColumnNumber; columnNumber++)
            {
                var excelRange = _workSheet.Cells[_rowNumber, columnNumber];
                rowData.Add(excelRange.Text);
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

        public void Reset()
        {
            Dispose();
        }
    }
}
