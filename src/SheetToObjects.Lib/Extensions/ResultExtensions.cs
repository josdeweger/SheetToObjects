using System;
using CSharpFunctionalExtensions;

namespace SheetToObjects.Lib.Extensions
{
    internal static class ResultExtensions
    {
        public static Result<TResult, TError> OnValidationSuccess<TResult, TError>(this Result<TResult, TError> result, Func<TResult, TResult> func)
            where TError : class
        {
            if (result.IsSuccess)
                return Result.Success<TResult, TError>(func(result.Value));

            return Result.Failure<TResult, TError>(result.Error);
        }

        public static Result<TResult, TError> OnValidationFailure<TResult, TError>(this Result<TResult, string> result, Func<string, TError> func) 
            where TError : class
        {
            if (result.IsFailure)
                return Result.Failure<TResult, TError>(func(result.Error));

            return Result.Success<TResult, TError>(result.Value);
        }
    }
}
