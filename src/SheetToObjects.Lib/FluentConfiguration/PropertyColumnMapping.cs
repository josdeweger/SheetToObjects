﻿using System.Collections.Generic;
using SheetToObjects.Lib.Validation;

namespace SheetToObjects.Lib.FluentConfiguration
{
    internal class PropertyColumnMapping : ColumnMapping, IUseHeaderRow
    {
        public string ColumnName { get; }

        public PropertyColumnMapping(string propertyName, string format, List<IParsingRule> parsingRules, List<IRule> rules) 
            : base(propertyName, format, parsingRules, rules)
        {
            ColumnName = propertyName;
            ColumnIndex = -1;
        }
        public void SetColumnIndex(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        public override string DisplayName => ColumnName;

    }
}