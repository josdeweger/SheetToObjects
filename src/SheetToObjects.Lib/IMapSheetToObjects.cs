using SheetToObjects.Lib.FluentConfiguration;
using SheetToObjects.Lib.Validation;
using System;
using System.Collections.Generic;

namespace SheetToObjects.Lib
{
    public interface IMapSheetToObjects
    {
        SheetMapper AddConfigFor<T>(Func<MappingConfigBuilder<T>, MappingConfigBuilder<T>> mappingConfigFunc)
            where T : new();

        List<IValidationError>? MapHeadersToIndex<T>(Row firstRow);
        MappingResult<T> Map<T>(Sheet sheet) where T : new();
        MappingRowResult<T> MapRow<T>(Row row) where T : new();
    }
}