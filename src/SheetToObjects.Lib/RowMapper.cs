using CSharpFunctionalExtensions;
using SheetToObjects.Core;
using SheetToObjects.Lib.FluentConfiguration;
using SheetToObjects.Lib.Validation;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Lib
{
    internal class RowMapper : IMapRow
    {
        private readonly IMapValue _valueMapper;

        public RowMapper(IMapValue valueMapper)
        {
            _valueMapper = valueMapper;
        }

        public Result<ParsedModel<T>, List<IValidationError>> Map<T>(Row row, MappingConfig mappingConfig)
            where T : new()
        {
            var rowValidationErrors = new List<IValidationError>();
            var obj = new T();
            var properties = obj.GetType().GetProperties().ToList();

            properties.ForEach(property =>
            {
                rowValidationErrors.AddRange(MapRow(row, mappingConfig, property, obj));
            });

            if (rowValidationErrors.Any())
                return Result.Failure<ParsedModel<T>, List<IValidationError>>(rowValidationErrors);

            return Result.Success<ParsedModel<T>, List<IValidationError>>(new ParsedModel<T>(obj, row.RowIndex));
        }

        private IEnumerable<IValidationError> MapRow<TModel>(
            Row row,
            MappingConfig mappingConfig,
            PropertyInfo property,
            TModel obj) where TModel : new()
        {
            ColumnMapping? columnMapping = mappingConfig.GetColumnMappingByPropertyName(property.Name);

            if (columnMapping is null)
                return new List<IValidationError>();

            Cell? cell = row.GetCellByColumnIndex(columnMapping.ColumnIndex);

            if (cell is null)
            {
                return HandleEmptyCell(columnMapping, row.RowIndex, property.Name)
                    .OnValue(error => new List<IValidationError> { error })
                    .OnEmpty(() => new List<IValidationError>())
                    .GetValueOrDefault();
            }

            var validationErrors = new List<IValidationError>();

            _valueMapper
                .Map(cell.Value.ToString(), property.PropertyType, row.RowIndex, columnMapping)
                .Tap(value =>
                {
                    if (value.ToString().IsNotNullOrEmpty())
                        property.SetValue(obj, value);
                })
                .TapError(validationErrors.Add);

            return validationErrors;
        }

        private static Maybe<IValidationError> HandleEmptyCell(ColumnMapping columnMapping, int rowIndex, string propertyName)
        {
            if (!columnMapping.IsRequired)
                return Maybe<IValidationError>.None;

            var parsingValidationError = ParsingValidationError.CellNotFound(
                columnMapping.ColumnIndex,
                rowIndex,
                columnMapping.DisplayName,
                propertyName);

            return Maybe<IValidationError>.From(parsingValidationError);
        }
    }
}