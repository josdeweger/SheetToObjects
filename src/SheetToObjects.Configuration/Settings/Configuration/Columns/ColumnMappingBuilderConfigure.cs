using SheetToObjects.Lib.FluentConfiguration;
using System.Reflection;

namespace SheetToObjects.Configuration.Settings.Configuration.Columns;

internal class ColumnMappingBuilderConfigure<T>
{
    private readonly MappingConfigBuilder<T> _configBuilder;

    public ColumnMappingBuilderConfigure(MappingConfigBuilder<T> configBuilder)
    {
        _configBuilder = configBuilder;
    }

    public void Configure(ColumnMappingConfig mappingConfig)
    {
        Func<ColumnMappingBuilder<T>, ColumnMapping> columnMappingBuilderFunc = (builder) =>
        {
            if (mappingConfig.WithHeader is not null)
            {
                builder.WithHeader(mappingConfig.WithHeader);
            }

            if (mappingConfig.WithColumnIndex is not null)
            {
                builder.WithColumnIndex(mappingConfig.WithColumnIndex.Value);
            }

            if (mappingConfig.WithColumnLetter is not null)
            {
                builder.WithColumnLetter(mappingConfig.WithColumnLetter);
            }

            if (mappingConfig.UsingFormat is not null)
            {
                builder.UsingFormat(mappingConfig.UsingFormat);
            }

            if (mappingConfig.IsRequired)
            {
                builder.IsRequired();
            }

            Type type = typeof(T);
            PropertyInfo propertyInfo = type.GetProperty(mappingConfig.MapTo)
                ?? throw new InvalidOperationException($"Property {mappingConfig.MapTo} not found for column mapping");

            ColumnMapping columnMapping = builder.MapTo(propertyInfo);
            return columnMapping;
        };

        _configBuilder.MapColumn(columnMappingBuilderFunc);
    }

    //private static Expression<Func<T, TProperty>> CreateLambdaExpression(string mapToProperty)
    //{
    //    Type type = typeof(T);

    //    ParameterExpression parameterExpr = Expression.Parameter(type, "o");
    //    PropertyInfo propertyInfo = type.GetProperty(mapToProperty)
    //        ?? throw new InvalidOperationException($"Property {mapToProperty} not found for column mapping");

    //    MemberExpression memberExpression = Expression.Property(parameterExpr, propertyInfo);
    //    LambdaExpression lambdaExpression = Expression.Lambda(memberExpression, parameterExpr);

    //}
}


