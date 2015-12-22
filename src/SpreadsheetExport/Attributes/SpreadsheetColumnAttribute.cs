using System;

namespace SpreadsheetExport.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class SpreadsheetColumnAttribute : Attribute
    {
        public enum TextModifierEnum
        {
            None = 0,
            Upper = 1,
            Lower = 2,
            TitleCase = 3
        }

        public TextModifierEnum Modifier { get; set; }

        public string ColumnHeader { get; private set; }

        public SpreadsheetColumnAttribute(string columnHeader)
        {
            ColumnHeader = columnHeader;
        }

        public string Format { get; set; }
    }
}