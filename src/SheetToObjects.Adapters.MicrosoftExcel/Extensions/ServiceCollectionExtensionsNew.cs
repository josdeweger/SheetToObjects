using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SheetToObjects.Adapters.MicrosoftExcel.Configurations;
using SheetToObjects.Adapters.MicrosoftExcel.Extensions.Assemblies;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace SheetToObjects.Adapters.MicrosoftExcel.Extensions;
public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddMicrosoftExcelSheetMapConfigs(
        this IServiceCollection services,
        IConfiguration configuration,
        Action<MicrosoftExcelConfigurationReaderOptions>? configure = null)
    {
        MicrosoftExcelConfigurationReaderOptions options = new();
        configure?.Invoke(options);

        IEnumerable<Assembly> assemblies = Assemblies(options);

        MicrosoftExcelSheetMapConfigBuilder builder = new(configuration, assemblies, options);
        var parentSection = configuration.GetSection(options.SectionPath);
        var childSections = parentSection.GetChildren();

        foreach (IConfigurationSection childSection in childSections)
        {
            string sheetModel = childSection.Key;
            MicrosoftExcelSheetMapConfig sheetMapConfig = builder.Build(childSection);

            services.AddOptions<MicrosoftExcelSheetMapConfig>(sheetModel)
                .Configure(options =>
                {
                    options.MappingConfig = sheetMapConfig.MappingConfig;
                    options.Range = sheetMapConfig.Range;
                });
        }

        return services;
    }

    internal static IEnumerable<Assembly> Assemblies(MicrosoftExcelConfigurationReaderOptions options)
    {
        return options.Assemblies ?? GetAssemblyFinder(options).GetAssemblies();
    }

    internal static AssemblyFinder GetAssemblyFinder(MicrosoftExcelConfigurationReaderOptions options)
    {
        AssemblyFinder assemblyFinder;
        if (options.ConfigurationAssemblySource is not null)
        {
            assemblyFinder = AssemblyFinder.ForSource(options.ConfigurationAssemblySource.Value);
        }
        else
        {
            assemblyFinder = options.DependencyContext == null ? AssemblyFinder.Auto() : AssemblyFinder.ForDependencyContext(options.DependencyContext);
        }

        assemblyFinder.GetAssemblies();
        return assemblyFinder;
    }
}
