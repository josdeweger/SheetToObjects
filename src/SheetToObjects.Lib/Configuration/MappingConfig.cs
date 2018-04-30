﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace SheetToObjects.Lib.Configuration
{
    public class MappingConfig
    {
        public bool DataHasHeaders { get; set; }

        public List<ColumnMapping> ColumnMappings = new List<ColumnMapping>();
        
        public ColumnMapping GetColumnMappingByPropertyName(string propertyName)
        {
            return ColumnMappings.FirstOrDefault(m =>
                m.PropertyName.Equals(propertyName, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}