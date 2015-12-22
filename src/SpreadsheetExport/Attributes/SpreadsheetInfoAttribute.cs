using System;

namespace AtrumSoft.SpreadsheetExport.Attributes
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
    public class SpreadsheetInfoAttribute : Attribute
    {
        public string SheetTitle { get; private set; }

        public SpreadsheetInfoAttribute(string sheetTitle)
        {
            SheetTitle = sheetTitle;
        }
    }
}