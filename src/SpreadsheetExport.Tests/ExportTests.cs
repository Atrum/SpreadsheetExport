using System.Collections.Generic;
using System.Diagnostics;
using FizzWare.NBuilder;
using NUnit.Framework;
using SpreadsheetExport.Extensions;
using SpreadsheetExport.Tests.Entities;

//using NUnit.Framework;

namespace SpreadsheetExport.Tests
{
    [TestFixture]
    public class ExportTests
    {
        private IList<Dummy> _dummyList;

        [SetUp]
        public void CreateEnumerable()
        {
            _dummyList = Builder<Dummy>.CreateListOfSize(10)
                .Build();
        }

        [Test]
        public void DummyExport()
        {
            const string thesheet = "sheet.xlsx";
            _dummyList.ToSpreadsheet(thesheet);
            Process.Start(thesheet);
        }

        [Test]
        public void DummyTemplate()
        {
            const string thesheet = "sheet2.xlsx";
            SpreadsheetExtensions.TemplateFor<Dummy>(thesheet);
            Process.Start(thesheet);
        }
    }
}