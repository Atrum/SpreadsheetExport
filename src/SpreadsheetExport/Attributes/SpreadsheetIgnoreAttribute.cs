using System;

namespace AtrumSoft.SpreadsheetExport.Attributes
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class SpreadsheetIgnoreAttribute : Attribute
    {
    }
}