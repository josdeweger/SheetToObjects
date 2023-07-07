using CSharpFunctionalExtensions;
using SheetToObjects.Lib.FluentConfiguration;
using SheetToObjects.Lib.Validation;
using System.Collections.Generic;

namespace SheetToObjects.Lib.SetValues;
public interface IRowSetter
{
    Maybe<List<IValidationError>> SetProperties<T>(Row row, T previousObj, T obj, MappingConfig mappingConfig)
        where T : new();
}