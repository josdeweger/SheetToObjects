using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace SheetToObjects.Configuration;
public static class ServiceCollectionExtensions
{
    private static IServiceCollection AddConfig(this IServiceCollection services)
    {
        IConfiguration config = null;
        services.Configure<string>("name", e => { });

        return services;
    }

    private static void ConfigureSections(IConfiguration configuration)
    {
        var configurationSections = configuration.GetSection("SheetToObjects")
            .GetChildren();

        foreach (IConfigurationSection section in configurationSections)
        {
            string serviceName = section.Key;
            var sectionValue = section.GetSection(section.Key);


        }


    }

    public static IServiceCollection CreateConfiguration(this IServiceCollection services, string sectionPath)
    {
        throw new Exception();
    }

    //public static IServiceCollection CreateConfiguration<T>(this IServiceCollection services, IConfigurationSection section)
    //{
    //    ConfigurationReader<T> configurationReader = new(section, null, null);
    //    configurationReader.Configure()
    //}
}
