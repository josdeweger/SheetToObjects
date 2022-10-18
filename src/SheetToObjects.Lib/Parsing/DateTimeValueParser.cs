using System;
using System.Globalization;
using CSharpFunctionalExtensions;
using SheetToObjects.Core;

namespace SheetToObjects.Lib.Parsing
{
    internal class DateTimeValueParser : IValueParsingStrategy
    {
        private readonly string _format;

        public DateTimeValueParser(string format)
        {
            _format = format;
        }

        public Result<object, string> Parse(string value)
        {
            var errorMessage = $"Cannot parse value '{value}' to DateTime using format '{_format}'";

            try
            {
                var parsedDateTime = DateTime.ParseExact(value, _format, CultureInfo.InvariantCulture);
                return Result.Success<object, string>(parsedDateTime);
            }
            catch (Exception)
            {
                return Result.Failure<object, string>(errorMessage);
            }
        }
    }
    
    internal class GuidValueParser : IValueParsingStrategy
    {
        private readonly bool _isNullable;

        public GuidValueParser(Type type)
        {
            _isNullable = type.IsNullable();
        }
        
        public Result<object, string> Parse(string value)
        {
            try
            {
                if (_isNullable && string.IsNullOrEmpty(value))
                    return Result.Success<object, string>(default!);
                
                return Result.Success<object, string>(Guid.Parse(value));
            }
            catch (Exception)
            {
                return Result.Failure<object, string>($"Could not parse value {value} to Guid");
            }
        }
    }
}