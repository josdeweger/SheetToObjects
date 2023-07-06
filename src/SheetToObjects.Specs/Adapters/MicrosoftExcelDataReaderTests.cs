using FluentAssertions;
using SheetToObjects.Adapters.MicrosoftExcel;
using SheetToObjects.Specs.SheetToObjectConfigs;
using SheetToObjects.Specs.TestModels;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace SheetToObjects.Specs.Adapters;
public class MicrosoftExcelDataReaderTests
{
    private const string TestModelFilePath = @"./TestFiles/TestModel.xlsx";
    private const string TestModelWithUnstandardHeaderFilePath = @"./TestFiles/TestModelWithUnstandardHeader.xlsx";

    [Fact]
    public void ReadExcel_WhenHasEmpty_ShouldBeEquals()
    {
        int max_number_of_excel_rows = 4;

        using var fileStream = new FileStream(TestModelFilePath, FileMode.Open);
        using var sr = new StreamReader(fileStream, Encoding.UTF8, false, 1024, true);

        var mapConfig = new TestModelMap();

        var excelRange = new ExcelRange(new ExcelCell("A", 2), new ExcelCell("J", max_number_of_excel_rows));
        MicrosoftExcelDataReader<TestModel> dataReader = MicrosoftExcelDataReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: false);

        var list = dataReader.ToList();

        list[0].MappingRowResult.ParsedModel?.Value.StringProperty.Should().Be("My string value");
        list.Count.Should().Be(3);
    }

    [Fact]
    public void ReadExcel_WhenHasEmpty_ShouldBeLess()
    {
        int max_number_of_excel_rows = 6;

        using var fileStream = new FileStream(TestModelFilePath, FileMode.Open);
        using var sr = new StreamReader(fileStream, Encoding.UTF8, false, 1024, true);

        var mapConfig = new TestModelMap();

        var excelRange = new ExcelRange(new ExcelCell("A", 2), new ExcelCell("J", max_number_of_excel_rows));
        MicrosoftExcelDataReader<TestModel> dataReader = MicrosoftExcelDataReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

        var list = dataReader.ToList();

        list[0].MappingRowResult.ParsedModel?.Value.StringProperty.Should().Be("My string value");
        list.Count.Should().Be(2);
    }

    [Fact]
    public void ReadExcelUnstandardPath_WhenHasEmpty_ShouldBeLess()
    {
        int max_number_of_excel_rows = 5;

        using var fileStream = new FileStream(TestModelWithUnstandardHeaderFilePath, FileMode.Open);
        using var sr = new StreamReader(fileStream, Encoding.UTF8, false, 1024, true);

        var mapConfig = new TestModelMap();

        var excelRange = new ExcelRange(new ExcelCell("A", 3), new ExcelCell("J", max_number_of_excel_rows));
        MicrosoftExcelDataReader<TestModel> dataReader = MicrosoftExcelDataReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

        var list = dataReader.ToList();

        list[0].MappingRowResult.ParsedModel?.Value.StringProperty.Should().Be("My string value");
        list.Count.Should().Be(2);
        list.ForEach(x => x.MappingRowResult.IsSuccess.Should().Be(true));
    }

    [Fact]
    public void WriteTo_Null_ShouldOk()
    {
        int max_number_of_excel_rows = 5;

        using var fileStream = new FileStream(TestModelFilePath, FileMode.Open);

        var mapConfig = new TestModelMap();

        var excelRange = new ExcelRange(new ExcelCell("A", 2), new ExcelCell("J", max_number_of_excel_rows));
        using var dataReader = MicrosoftExcelDataReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

        dataReader.Worksheet.Cells[2, 2].Value = 55;

        dataReader.ExcelPackage.SaveAsAsync(@"./TestFiles/TestModel2.xlsx");
    }
}

public class MyClass
{

}
