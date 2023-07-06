using SheetToObjects.Lib;
using System.Collections.Generic;

namespace SheetToObjects.Adapters.MicrosoftExcel
{
    internal class MicrosoftExcelToRowConverter : IConvertDataToRow<List<string>>
    {
        public Row ConvertRow(List<string> rowData, int rowIndex)
        {
            List<Cell> cells = rowData.RowDataToCells(rowIndex);
            return new Row(cells, rowIndex);
        }
    }
}
