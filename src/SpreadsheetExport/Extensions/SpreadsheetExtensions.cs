using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetExport.Extensions
{
    public static class SpreadsheetExtensions
    {
        public static void ToSpreadsheet<T>(this IEnumerable<T> source,string filename)
        {
            var exportTool = new SpreadsheetExportFromType<T>(source,filename);
            exportTool.Generate();
        }

        public static byte[] ToSpreadsheet<T>(this IEnumerable<T> source)
        {
            var stream = new MemoryStream();
            var exportTool = new SpreadsheetExportFromType<T>(source, stream);
            exportTool.Generate();
            return stream.ToArray();
        }
    }
}
