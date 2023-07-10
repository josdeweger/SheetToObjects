using FluentAssertions;
using SheetToObjects.Adapters.MicrosoftExcel;
using SheetToObjects.Specs.SheetToObjectConfigs;
using SheetToObjects.Specs.TestModels;
using System;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
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
        MicrosoftExcelSheetReader<TestModel> dataReader = MicrosoftExcelSheetReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: false);

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
        MicrosoftExcelSheetReader<TestModel> dataReader = MicrosoftExcelSheetReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

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
        MicrosoftExcelSheetReader<TestModel> dataReader = MicrosoftExcelSheetReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

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
        using var dataReader = MicrosoftExcelSheetReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

        dataReader.Worksheet.Cells[2, 2].Value = 55;

        dataReader.ExcelPackage.SaveAsAsync(@"./TestFiles/TestModel2.xlsx");
    }

    [Fact]
    public void TestSave()
    {
        int max_number_of_excel_rows = 5;

        using var fileStream = new FileStream(TestModelFilePath, FileMode.Open);

        var mapConfig = new TestModelMap();

        var excelRange = new ExcelRange(new ExcelCell("A", 2), new ExcelCell("J", max_number_of_excel_rows));
        using var dataReader = MicrosoftExcelSheetReader<TestModel>.CreateReader(fileStream, "sheet1", excelRange, mapConfig, stopReadingOnEmptyRow: true);

        foreach (var obj in dataReader)
        {
            TestModel testModel = obj.MappingRowResult.ParsedModel!.Value;
            testModel.StringProperty = testModel.StringProperty + " новое значение";
            testModel.IntProperty = testModel.IntProperty + 500;

            int rowNumber = obj.ExcelRow.Row;
            dataReader.SetValue(testModel, rowNumber);
        }

        dataReader.ExcelPackage.SaveAsAsync(@"./TestFiles/TestModel2.xlsx");
    }

    [Fact]
    public void PropertyExpression()
    {
        TestModel model = new TestModel()
        {
            IntProperty = 1,
        };

        Expression<Func<TestModel, int>> exp = o => o.IntProperty;

        ParameterExpression parameterExpr = Expression.Parameter(typeof(TestModel), "o");
        PropertyInfo propertyInfo = typeof(TestModel).GetProperty("IntProperty")!;
        MemberExpression memberExpression = Expression.Property(parameterExpr, propertyInfo);
        LambdaExpression lambdaExpression = Expression.Lambda(memberExpression, parameterExpr);

        var @delegate = lambdaExpression.Compile();
        var type = @delegate.GetType();
        int result = (int)@delegate.Method.Invoke(model, null);
    }
}

public class MyClass
{

}
