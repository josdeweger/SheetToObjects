using System.Collections.Generic;
using System.Linq;
using SheetToObjects.Lib.Validation;

namespace SheetToObjects.Lib
{
    public class MappingResult<T>
        where T : new()
    {
        public List<ParsedModel<T>> ParsedModels { get; set; }
        public List<IValidationError> ValidationErrors { get; set; }

        public bool IsFailure => ValidationErrors.Any();
        public bool IsSuccess => !IsFailure;

        private MappingResult(List<ParsedModel<T>> parsedModels, List<IValidationError> validationErrors)
        {
            ParsedModels = parsedModels;
            ValidationErrors = validationErrors;
        }

        public static MappingResult<T> Create(List<ParsedModel<T>> parsedModels,List<IValidationError> validationErrors)
        {
            return new MappingResult<T>(parsedModels, validationErrors);
        }
    }
}