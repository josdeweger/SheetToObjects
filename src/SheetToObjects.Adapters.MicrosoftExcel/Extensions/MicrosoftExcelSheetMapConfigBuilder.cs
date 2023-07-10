using Microsoft.Extensions.Configuration;
using SheetToObjects.Adapters.MicrosoftExcel.Configurations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Extensions;
internal class MicrosoftExcelSheetMapConfigBuilder
{
    private readonly IConfiguration _configuration;
    private readonly MicrosoftExcelConfigurationReaderOptions _options;
    private readonly Assembly[] _assemblies;

    public MicrosoftExcelSheetMapConfigBuilder(
        IConfiguration configuration,
        IEnumerable<Assembly> assemblies,
        MicrosoftExcelConfigurationReaderOptions options)
    {
        _configuration = configuration;
        _options = options;
        _assemblies = assemblies as Assembly[] ?? assemblies.ToArray();
    }

    public MicrosoftExcelSheetMapConfig Build(IConfigurationSection section)
    {
        MicrosoftExcelSheetMapConfig sheetMapConfig = new();

        string sheetModel = section.Key;
        Type sheetModelType = GetSheetModelType(sheetModel);

        var configReaderType = typeof(MicrosoftExcelConfigurationReader<>).MakeGenericType(sheetModelType);
        object[] constructorParams = new object[]
        {
            section,
            _options,
            _configuration
        };

        object configReader = Activator.CreateInstance(configReaderType, constructorParams)!;
        MethodInfo configureMethod = configReaderType.GetMethods()
            .FirstOrDefault(x =>
            {
                ParameterInfo[] parameters = x.GetParameters();
                if (parameters.Length != 1)
                    return false;

                return parameters[0].ParameterType == typeof(MicrosoftExcelSheetMapConfig);
            }) ?? throw new InvalidOperationException("Method not found");

        configureMethod.Invoke(configReader, new object[] { sheetMapConfig });
        return sheetMapConfig;
    }

    private Type GetSheetModelType(string name)
    {
        List<Type> types = new();
        foreach (Assembly assembly in _assemblies)
            foreach (Type type in assembly.GetTypes())
            {
                if (type.Name != name)
                    continue;

                types.Add(type);
            }

        if (types.Count > 1)
            throw new InvalidOperationException($"Found more than 1 type {name}");

        if (types.Count == 0)
            throw new InvalidOperationException($"Not found type {name}");

        return types[0];
    }
}
