using SheetToObjects.Lib;
using SheetToObjects.Lib.Validation;
using System.Collections.Generic;

public record MappingRowResult<T>
    where T : new()
{
    public ParsedModel<T>? ParsedModel { get; set; }
    public List<IValidationError> ValidationError { get; set; }

    public bool IsFailure => ValidationError is not null;
    public bool IsSuccess => !IsFailure;

    private MappingRowResult(ParsedModel<T>? parsedModel, List<IValidationError> validationError)
    {
        ParsedModel = parsedModel;
        ValidationError = validationError;
    }

    public static MappingRowResult<T> Create(ParsedModel<T>? parsedModel, List<IValidationError> validationError)
    {
        return new MappingRowResult<T>(parsedModel, validationError);
    }
}