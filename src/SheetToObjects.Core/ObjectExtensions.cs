﻿namespace SheetToObjects.Core
{
    internal static class ObjectExtensions
    {
        public static bool IsNull(this object? obj)
        {
            return obj == null;
        }

        public static bool IsNotNull(this object obj)
        {
            return obj != null;
        }
    }
}
