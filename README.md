# SpreadsheetExport

Quickly turns an IEnumerable into an excel spreadsheet

How to use it:

Download the NuGet package from Gallery

add the namespace to your code

    using AtrumSoft.SpreadsheetExport.Extensions;

then just export your IEnummerable to a file

    dummyList.ToSpreadsheet("myExcel.xlsx");

or you can get the byte array 

    var array = dummyList.ToSpreadsheet();


But wait!!, i dont need all the properties and the column names aren't very friendly, well you can use the custom attributes, or use your already declared display names

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

Thats it.

If you need just the template for your class you can just do this

    SpreadsheetExtensions.TemplateFor<Dummy>("myExcel.xlsx");

Or get it as a byte array

    var array = SpreadsheetExtensions.TemplateFor<Dummy>();

