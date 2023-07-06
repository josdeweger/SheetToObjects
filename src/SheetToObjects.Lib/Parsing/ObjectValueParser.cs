using CSharpFunctionalExtensions;
using System;
using System.Globalization;

namespace SheetToObjects.Lib.Parsing
{
    internal class ObjectValueParser : IValueParsingStrategy
    {
        private readonly Type _type;

        public ObjectValueParser(Type type)
        {
            _type = Nullable.GetUnderlyingType(type) ?? type;
        }

        public Result<object, string> Parse(string value)
        {
            try
            {
                object? parsedValue = null;

                if (_type == typeof(double) || _type == typeof(decimal) || _type == typeof(double?) || _type == typeof(decimal?))
                {
                    value = value.Replace(',', '.');
                    parsedValue = Convert.ChangeType(value, _type, CultureInfo.InvariantCulture);
                }
                else
                {
                    parsedValue = Convert.ChangeType(value, _type);
                }

                return Result.Success<object, string>(parsedValue);
            }
            catch (Exception)
            {
                return Result.Failure<object, string>($"Cannot parse value '{value}' to type '{_type.Name}'");
            }
        }
    }
}