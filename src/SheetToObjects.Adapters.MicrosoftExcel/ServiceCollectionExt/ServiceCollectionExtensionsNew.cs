using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SheetToObjects.Adapters.MicrosoftExcel.Configurations;
using System;

namespace SheetToObjects.Adapters.MicrosoftExcel.ServiceCollectionExt;
public static class ServiceCollectionExtensionsNew
{
    public static IServiceCollection AddMicrosoftExcelSheetMapConfigs(
        this IServiceCollection services,
        IConfiguration configuration,
        Action<MicrosoftExcelConfigurationReaderOptions>? configure = null)
    {
        MicrosoftExcelConfigurationReaderOptions options = new();
        configure?.Invoke(options);

        MicrosoftExcelSheetMapConfigBuilder builder = new(configuration, options);
        var parentSection = configuration.GetSection(options.SectionName);
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
}
