using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetExport.Helpers
{
    internal class SpreadsheetDocumentCreatorHelper
    {
        private static void SetCellValue(CellType cell, object value, OpenXmlElement sharedStringTable, string type)
        {
            if (IsNumeric(value))
            {
                AddNumericValue(cell, value);
                return;
            }

            DateTime temp;
            if (IsDate(value, type, out temp))
            {
                AddDateValue(cell, temp);
                return;
            }

            AddString(cell, value, sharedStringTable);
        }

        private static bool IsNumeric(object value)
        {
            return value is int || value is long || value is decimal || value is float;
        }

        private static bool IsDate(object value, string type, out DateTime temp)
        {
            var parsed = DateTime.MinValue;
            var isValid = type.Contains("DateTime") &&
                   DateTime.TryParseExact(value.ToString(), "ddd dd/MMM/yyyy", CultureInfo.CurrentCulture,
                       DateTimeStyles.None, out parsed);
            temp = parsed;
            return isValid;
        }

        private static void AddDateValue(CellType cell, DateTime temp)
        {
            cell.StyleIndex = 3U;
            cell.CellValue = new CellValue(temp.ToOADate().ToString(CultureInfo.InvariantCulture));
        }

        private static void AddString(CellType cell, object value, OpenXmlElement sharedStringTable)
        {
            cell.DataType = CellValues.SharedString;
            var sharedString =
                sharedStringTable.Descendants<SharedStringItem>().FirstOrDefault(si => si.Text.Text == value.ToString());
            if (sharedString != null)
            {
                var existingStringvalue = sharedStringTable.ToList().IndexOf(sharedString);
                cell.CellValue = new CellValue(existingStringvalue.ToString(CultureInfo.InvariantCulture));
                return;
            }
            cell.CellValue = new CellValue(AddSharedString(sharedStringTable, value.ToString()));
        }

        private static void AddNumericValue(CellType cell, object value)
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(value.ToString());
        }

        private static string AddSharedString(OpenXmlElement stringTable, string text)
        {
            var sharedStringItem1 = new SharedStringItem();
            var text1 = new Text { Text = text };
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            sharedStringItem1.Append(text1);
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            stringTable.Append(sharedStringItem1);
            return stringTable.Descendants<SharedStringItem>()
                .ToList()
                .IndexOf(sharedStringItem1)
                .ToString(CultureInfo.InvariantCulture);
        }

        public static Cell CreateCell(object value, SharedStringTable sharedStringTable, string nextCell, string type, uint styleIndex)
        {
            var cell = new Cell
            {
                CellReference = nextCell,
                StyleIndex = styleIndex
            };

            SetCellValue(cell, value, sharedStringTable, type);
            return cell;
        }
    }
}