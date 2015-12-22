using System;

namespace SpreadsheetExport.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class SpreadsheetIgnoreAttribute : Attribute
    {
    }
}