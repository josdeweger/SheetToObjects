using CSharpFunctionalExtensions;
using SheetToObjects.Core;
using SheetToObjects.Lib.FluentConfiguration;
using SheetToObjects.Lib.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SheetToObjects.Lib.SetValues
{
    public class RowSetter : IRowSetter
    {
        private readonly IRawExcelService _rawExcelService;

        public RowSetter(IRawExcelService rawExcelService)
        {
            _rawExcelService = rawExcelService;
        }

        public Maybe<List<IValidationError>> SetProperties<T>(Row row, T previousObj, T obj, MappingConfig mappingConfig)
            where T : new()
        {
            var rowValidationErrors = new List<IValidationError>();
            var properties = obj.GetType().GetProperties().ToList();

            properties.ForEach(property =>
            {
                object? previousPropertyValue = property.GetValue(previousObj);
                object? currentPropertyValue = property.GetValue(obj);

                if (previousPropertyValue == null && currentPropertyValue == null)
                    return;

                bool equals = previousPropertyValue?.Equals(currentPropertyValue) ?? false;
                if (equals)
                    return;

                rowValidationErrors.AddRange(SetPropertyValue(row, mappingConfig, obj, property));
            });

            if (rowValidationErrors.Any())
                return Maybe<List<IValidationError>>.From(rowValidationErrors);

            return Maybe<List<IValidationError>>.None;
        }

        private IEnumerable<IValidationError> SetPropertyValue(
            Row row,
            MappingConfig mappingConfig,
            object obj,
            PropertyInfo property)
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

            object? propertyValue = property.GetValue(obj);

            _rawExcelService.SetPropertyValue(cell, propertyValue);
            return Array.Empty<IValidationError>();
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