using Microsoft.Extensions.Configuration;

namespace SheetToObjects.Configuration.Settings.Configuration;

public static class ConfigurationSectionExtensions
{
    public static T GetRequired<T>(this IConfigurationSection section)
    {
        T? value = section.Get<T>() ?? throw new InvalidOperationException($"");
        return value;
    }
}
