using System;
using System.ComponentModel;
using SpreadsheetExport.Attributes;

namespace SpreadsheetExport.Tests.Entities
{
    [SpreadsheetInfo("My sheet")]
    internal class Dummy
    {
        [SpreadsheetColumn("This is a string")]
        public string AString { get; set; }

        [DisplayName("This is a string with DisplayName")]
        public string AnotherString { get; set; }

        [SpreadsheetColumn("This is a title string",Modifier = SpreadsheetColumnAttribute.TextModifierEnum.TitleCase)]
        public string ATitleString { get; set; }

        [SpreadsheetColumn("This is an upper string", Modifier = SpreadsheetColumnAttribute.TextModifierEnum.Upper)]
        public string AnUpperString { get; set; }

        [SpreadsheetIgnore]
        public string IgnoreMe { get; set; }

        [SpreadsheetColumn("A date",Format = "ddd dd/MMM/yyyy")]
        public DateTime Date { get; set; }

        [SpreadsheetColumn("An integer")]
        public int AnInteger { get; set; }
    }
}
