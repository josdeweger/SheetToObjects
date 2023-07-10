using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SheetToObjects.Adapters.MicrosoftExcel.Configurations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.ServiceCollectionExt;
public static class ServiceCollectionExtensionsOld
{
    public static IServiceCollection AddMicrosoftExcelSheetMapConfigsOld(
        this IServiceCollection services,
        IConfiguration configuration,
        Action<MicrosoftExcelConfigurationReaderOptions>? configure = null)
    {
        MicrosoftExcelSheetMapConfig sheetMapConfig = new();
        MicrosoftExcelConfigurationReaderOptions options = new();
        configure?.Invoke(options);

        sheetMapConfig.AddConfigurationOld(configuration, options);

        return services;
    }

    public static void AddConfigurationOld(
        this MicrosoftExcelSheetMapConfig sheetMapConfig,
        IConfiguration configuration,
        MicrosoftExcelConfigurationReaderOptions options)
    {
        var parentSection = configuration.GetSection(options.SectionName);
        var childSections = parentSection.GetChildren();

        foreach (IConfigurationSection childSection in childSections)
        {
            string sheetModel = childSection.Key;
            Type sheetModelType = GetSheetModelType(options.Assemblies, sheetModel);

            var configReaderType = typeof(MicrosoftExcelConfigurationReader<>).MakeGenericType(sheetModelType);
            object[] constructorParams = new object[]
            {
                childSection,
                options,
                configuration
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

            IServiceCollection services = null;
            services.AddOptions<MicrosoftExcelSheetMapConfig>(sheetModel)
                .Configure(options =>
                {
                    options.MappingConfig = sheetMapConfig.MappingConfig;
                    options.Range = sheetMapConfig.Range;
                });
        }
    }

    private static void T(MicrosoftExcelSheetMapConfig sheetMapConfig, IConfigurationSection childSection, IConfiguration configuration, MicrosoftExcelConfigurationReaderOptions options)
    {
        string sheetModel = childSection.Key;
        Type sheetModelType = GetSheetModelType(options.Assemblies, sheetModel);

        var configReaderType = typeof(MicrosoftExcelConfigurationReader<>).MakeGenericType(sheetModelType);
        object[] constructorParams = new object[]
        {
                childSection,
                options,
                configuration
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

        IServiceCollection services = null;
        services.AddOptions<MicrosoftExcelSheetMapConfig>(sheetModel)
            .Configure(options =>
            {
                options.MappingConfig = sheetMapConfig.MappingConfig;
                options.Range = sheetMapConfig.Range;
            });
    }

    private static Type GetSheetModelType(Assembly[] assemblies, string name)
    {
        List<Type> types = new();
        foreach (Assembly assembly in assemblies)
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
