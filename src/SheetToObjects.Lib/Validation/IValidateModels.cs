using SheetToObjects.Lib.FluentConfiguration;
using System.Collections.Generic;

namespace SheetToObjects.Lib.Validation
{
    internal interface IValidateModels
    {
        ValidationResult<ParsedModel<TModel>> Validate<TModel>(
            List<ParsedModel<TModel>> parsedModels,
            List<ColumnMapping> columnMappings)
            where TModel : new();
        ValidationResult<ParsedModel<TModel>> ValidateRow<TModel>(
            ParsedModel<TModel> parsedModels,
            ColumnMapping columnMappings)
            where TModel : new();
    }
}