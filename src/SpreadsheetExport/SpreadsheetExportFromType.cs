using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using AtrumSoft.SpreadsheetExport.Attributes;
using AtrumSoft.SpreadsheetExport.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace AtrumSoft.SpreadsheetExport
{
    public class SpreadsheetExportFromType<T>
    {
        private readonly IEnumerable<T> _source;
        private readonly List<PropertyInfo> _propertyList;
        private readonly SpreadsheetDocument _spreadSheet;
        private WorksheetPart _workSheetPart;
        private readonly string _sheetName;
        private SharedStringTablePart _sharedStringsPart;
        private SheetData _sheetData;

        public SpreadsheetExportFromType(IEnumerable<T> source, Stream stream)
        {
            _source = source;
            _propertyList = ExtractProperties();
            _spreadSheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            InitDocument();
        }

        public SpreadsheetExportFromType(IEnumerable<T> source, string filename)
        {
            _source = source;
            _propertyList = ExtractProperties();
            _spreadSheet = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
            InitDocument();
        }

        public SpreadsheetExportFromType(IEnumerable<T> source, Stream stream, string sheetName)
        {
            _source = source;
            _propertyList = ExtractProperties();
            _spreadSheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            _sheetName = sheetName;
            InitDocument();
        }


        public void Generate()
        {
            CreateDocumentContent();
        }

        private void CreateDocumentContent()
        {
            FillRows(_source);
            Save();
        }

        private void InitDocument()
        {
            _spreadSheet.AddWorkbookPart();
            _spreadSheet.WorkbookPart.Workbook = new Workbook();
            _workSheetPart = _spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
            _workSheetPart.Worksheet = new Worksheet();
            _sharedStringsPart = _spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            CreateColumnsAndHeaders();
        }

        private void CreateColumnsAndHeaders()
        {
            var propertyDict = GetMaxLenghtPropertyDictionary(_propertyList, _source);
            var columns1 = ConfigureColumns(propertyDict);
            _workSheetPart.Worksheet.Append(columns1);
            _spreadSheet.WorkbookPart.CreateStyleSheet();
            var sharedStringTable = new SharedStringTable();
            _sheetData = new SheetData();
            AppendHeaders(_sheetData, sharedStringTable, _propertyList);
            SetSheetata();
            _sharedStringsPart.SharedStringTable = sharedStringTable;
        }

        private void Save()
        {
            Debug.Assert(_sheetData != null, "_sheetData != null");
            _workSheetPart.Worksheet.Append(_sheetData);
            _spreadSheet.WorkbookPart.Workbook.Save();
            _spreadSheet.Dispose();
        }

        private void SetSheetata()
        {
            _spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();
            _spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());
            _spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>()
                .AppendChild(new Sheet
                {
                    Id = WorkSheetPartId,
                    SheetId = 1,
                    Name = GetSheetTitle()
                });
        }

        private string WorkSheetPartId => _spreadSheet.WorkbookPart.GetIdOfPart(_workSheetPart);

        private void FillRows(IEnumerable<T> source)
        {
            UInt32Value counter = 2;
            foreach (var element in source)
            {
                AppendRow(_sheetData, _sharedStringsPart.SharedStringTable, element, _propertyList, counter);
                counter++;
            }
        }

        private static Dictionary<PropertyInfo, double> GetMaxLenghtPropertyDictionary(IEnumerable<PropertyInfo> props, IEnumerable<T> source)
        {
            const double columnRelativeWidth = 1.3;
            return props.ToDictionary(p => p,
                p =>
                {
                    var maxlenght = Maxlenght(source.ToArray(), p);
                    var header = GetHeader(p);
                    if (header.Length > maxlenght)
                        return header.Length* columnRelativeWidth;
                    return maxlenght * columnRelativeWidth;
                });
        }

        private static int Maxlenght(T[] source, PropertyInfo p)
        {
            if (!source.Any())
                return 0;
            return source.Max(e =>
            {
                var value = p.GetValue(e, null);
                return value == null ? 0 : Regex.Replace(value.ToString(), @"\s+", "").Length;
            });
        }

        private static List<PropertyInfo> ExtractProperties()
        {
            return typeof(T).GetProperties()
                .Where(p => !p.GetCustomAttributes(true).OfType<SpreadsheetIgnoreAttribute>().Any())
                .ToList();
        }

        private string GetSheetTitle()
        {
            if (!string.IsNullOrWhiteSpace(_sheetName)) return _sheetName;
            var attr = typeof(T).GetCustomAttributes(true).OfType<SpreadsheetInfoAttribute>().FirstOrDefault();
            return attr != null ? attr.SheetTitle : typeof(T).Name;
        }

        private static void AppendRow(OpenXmlElement sheetData, SharedStringTable sharedStringTable, T element, IEnumerable<PropertyInfo> propertyList, UInt32Value counter)
        {
            var firstChar = 65;
            var row = new Row { RowIndex = counter };
            foreach (var prop in propertyList)
            {
                var attr = prop.GetCustomAttributes(true).OfType<SpreadsheetColumnAttribute>().FirstOrDefault();
                var nextCell = $"{Convert.ToChar(firstChar)}{counter}";
                var value = prop.GetValue(element, null);
                var type = prop.PropertyType.FullName;
                value = SetFormat(value, attr, type);
                row.Append(CreateCell(value, sharedStringTable, nextCell, type, true));
                firstChar++;
            }
            sheetData.Append(row);
        }

        private static string Modify(SpreadsheetColumnAttribute.TextModifierEnum modifier, string value)
        {
            var textInfo = CultureInfo.CurrentCulture.TextInfo;
            switch (modifier)
            {
                case SpreadsheetColumnAttribute.TextModifierEnum.Upper:
                    return value.ToUpper();
                case SpreadsheetColumnAttribute.TextModifierEnum.Lower:
                    return value.ToLower();
                case SpreadsheetColumnAttribute.TextModifierEnum.TitleCase:
                    return textInfo.ToTitleCase(value);
                default:
                    return value;
            }
        }

        private static object SetFormat(object value, SpreadsheetColumnAttribute attr, string type)
        {
            if (value == null)
                return string.Empty;
            if (attr == null)
                return value;
            if (!string.IsNullOrEmpty(attr.Format))
            {
                var formato = "{0:" + attr.Format + "}";
                value = string.Format(formato, value);
            }
            if (!type.Contains("DateTime"))
                value = Modify(attr.Modifier, value.ToString());

            return value;
        }

        private static void AppendHeaders(OpenXmlElement sheetData, SharedStringTable sharedStringTable, IEnumerable<PropertyInfo> propertyList)
        {
            var firstChar = 65;
            var row = new Row { RowIndex = 1 };
            foreach (var prop in propertyList)
            {
                var nextCell = $"{Convert.ToChar(firstChar)}{1}";

                var value = GetHeader(prop);
                var type = prop.PropertyType.FullName;
                row.Append(CreateCell(value, sharedStringTable, nextCell, type));
                firstChar++;
            }
            sheetData.Append(row);
        }

        private static string GetHeader(PropertyInfo prop)
        {
            var valueAttr = prop.GetCustomAttributes(true).OfType<SpreadsheetColumnAttribute>().FirstOrDefault();
            var displayNameAttr = prop.GetCustomAttributes(true).OfType<DisplayNameAttribute>().FirstOrDefault();
            string value;
            if (valueAttr != null)
                value = valueAttr.ColumnHeader;
            else if (displayNameAttr != null)
                value = displayNameAttr.DisplayName;
            else
                value = prop.Name;
            return value;
        }

        private static Columns ConfigureColumns(IEnumerable<KeyValuePair<PropertyInfo, double>> propertyList)
        {
            var columns1 = new Columns();
            var i = 1;
            foreach (var keyValuePair in propertyList)
            {
                var column1 = new Column { Min = (uint)i, Max = (uint)i, Width = keyValuePair.Value, BestFit = true, CustomWidth = true };
                columns1.Append(column1);
                i++;
            }
            return columns1;
        }

        private static Cell CreateCell(object value, SharedStringTable sharedStringTable, string nextCell, string type, bool isHeader = false)
        {
            return SpreadsheetDocumentCreatorHelper.CreateCell(value, sharedStringTable, nextCell, type, isHeader ? 0U : 1U);
        }
    }
}