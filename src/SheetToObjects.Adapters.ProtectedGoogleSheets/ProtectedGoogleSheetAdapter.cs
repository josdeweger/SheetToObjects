﻿using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4.Data;
using SheetToObjects.Adapters.GoogleSheets.Shared.Extensions;
using SheetToObjects.Core;
using SheetToObjects.Lib;
using Sheet = SheetToObjects.Lib.Sheet;

namespace SheetToObjects.Adapters.ProtectedGoogleSheets
{
    internal class ProtectedGoogleSheetAdapter : IConvertResponseToSheet<ValueRange>
    {
        public Sheet Convert(ValueRange googleSheetData)
        {
            if (googleSheetData.IsNull())
                throw new ArgumentException(nameof(googleSheetData));

            if (googleSheetData.Values.IsNull() || !googleSheetData.Values.Any())
                return new Sheet(new List<Row>());

            var cells = googleSheetData.Values.ToRows();

            return new Sheet(cells);
        }
    }
}