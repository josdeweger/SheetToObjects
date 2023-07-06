namespace SheetToObjects.Lib
{
    public interface IConvertDataToSheet<in TData>
    {
        Sheet Convert(TData sheetData);
    }

    public interface IConvertDataToRow<in TData>
    {
        Row ConvertRow(TData rowData, int rowIndex);
    }
}