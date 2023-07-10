using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SheetToObjects.Adapters.MicrosoftExcel.Configurations;
using SheetToObjects.Adapters.MicrosoftExcel.ServiceCollectionExt;
using SheetToObjects.ConfigurationTests.Support;
using SheetToObjects.ConfigurationTests.TestModels;
using Xunit;

namespace SheetToObjects.ConfigurationTests.Settings.Configuration;
public class MicrosoftExcelConfigurationReaderTests
{
    [Fact]
    public void ConfigureTypes()
    {
        string json = """
{
  "SheetToObjects": {
    "MicrosoftExcel": {
      "TestModel": {
        "MapColumns": [
          {
            "WithHeader": "StringProperty",
            "IsRequired": true,
            "MapTo": "StringProperty"
          }
        ],
        "Range": {
          "HeaderRow": 2,
          "FromColumn": "A",
          "ToColumn":  "Z"
        },
        "SkipOnEmptyRow": true
      }
    }
  }
}
""";

        MicrosoftExcelSheetMapConfig config = new();

        IConfigurationRoot root = new ConfigurationBuilder()
            .Add(new JsonStringConfigSource(json))
            .Build();

        var assemblies = AppDomain.CurrentDomain.GetAssemblies();
        MicrosoftExcelConfigurationReaderOptions options = new(assemblies);

        //config.AddConfiguration(root, options);

        var section = JsonStringConfigSource.LoadSection(json, "SheetToObjects:MicrosoftExcel:TestModel");
        var reader = new MicrosoftExcelConfigurationReader<TestModel>(section, new MicrosoftExcelConfigurationReaderOptions(), section);
        reader.Configure(config);
    }

    [Fact]
    public void AddToSErviceCollection()
    {
        string json = """
{
  "SheetToObjects": {
    "MicrosoftExcel": {
      "TestModel": {
        "MapColumns": [
          {
            "WithHeader": "StringProperty",
            "IsRequired": true,
            "MapTo": "StringProperty"
          }
        ],
        "Range": {
          "HeaderRow": 2,
          "FromColumn": "A",
          "ToColumn":  "Z"
        },
        "SkipOnEmptyRow": true
      }
    }
  }
}
""";


        IConfigurationRoot root = new ConfigurationBuilder()
            .Add(new JsonStringConfigSource(json))
            .Build();

        var assemblies = AppDomain.CurrentDomain.GetAssemblies();
        MicrosoftExcelConfigurationReaderOptions options = new(assemblies);

        //config.AddConfiguration(root, options);
        ServiceCollection services = new();

        services.AddMicrosoftExcelSheetMapConfigs(root, configure =>
        {
            configure.Assemblies
        });

        var section = JsonStringConfigSource.LoadSection(json, "SheetToObjects:MicrosoftExcel:TestModel");
        var reader = new MicrosoftExcelConfigurationReader<TestModel>(section, new MicrosoftExcelConfigurationReaderOptions(), section);
        //reader.Configure(config);
    }

    [Fact]
    public void Test1()
    {
        string json = """
{
  "SheetToObjects": {
    "MicrosoftExcel": {
      "TestModel": {
        "MapColumns": [
          {
            "WithHeader": "StringProperty",
            "IsRequired": true,
            "MapTo": "StringProperty"
          }
        ],
        "Range": {
          "HeaderRow": 2,
          "FromColumn": "A",
          "ToColumn":  "Z"
        },
        "SkipOnEmptyRow": true
      }
    }
  }
}
""";

        MicrosoftExcelSheetMapConfig config = new();

        var t = new ConfigurationBuilder()
            .Add(new JsonStringConfigSource(json))
            .Build();

        var t1 = t.GetSection("SheetToObjects");

        var section = JsonStringConfigSource.LoadSection(json, "SheetToObjects:MicrosoftExcel:TestModel");
        var reader = new MicrosoftExcelConfigurationReader<TestModel>(section, new MicrosoftExcelConfigurationReaderOptions(), section);
        reader.Configure(config);
    }
}
