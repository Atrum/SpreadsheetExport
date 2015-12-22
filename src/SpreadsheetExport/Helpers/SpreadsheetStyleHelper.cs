using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossiblyMistakenUseOfParamsMethod

namespace AtrumSoft.SpreadsheetExport.Helpers
{
    internal static class SpreadsheetStyleHelper
    {
        public static void CreateStyleSheet(this WorkbookPart part)
        {
            var stylePart = part.AddNewPart<WorkbookStylesPart>();
            var stylesheet1 = new Stylesheet {MCAttributes = new MarkupCompatibilityAttributes {Ignorable = "x14ac"}};
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            GenerateNumbering(stylesheet1);
            GenerateFonts(stylesheet1);
            GenerateFills(stylesheet1);
            GenerateBorders(stylesheet1);
            GenerateCellStyleFormats(stylesheet1);
            GenerateCellFormats(stylesheet1);
            GenerateCellStyles(stylesheet1);
            stylePart.Stylesheet = stylesheet1;
        }

        private static void GenerateNumbering(OpenXmlElement stylesheet)
        {
            var numberingFormats1 = new NumberingFormats {Count = 1U};
            var numberingFormat1 = new NumberingFormat {NumberFormatId = 164U, FormatCode = "ddd dd/MMM/yyyy"};
            numberingFormats1.Append(numberingFormat1);
            stylesheet.Append(numberingFormats1);
        }

        private static void GenerateFills(OpenXmlElement stylesheet)
        {
            var fills1 = new Fills {Count = 2U};
            var fill1 = new Fill();
            var patternFill1 = new PatternFill {PatternType = PatternValues.None};
            fill1.Append(patternFill1);
            var fill2 = new Fill();
            var patternFill2 = new PatternFill {PatternType = PatternValues.Gray125};
            fill2.Append(patternFill2);
            fills1.Append(fill1);
            fills1.Append(fill2);
            stylesheet.Append(fills1);
        }

        private static void GenerateFonts(OpenXmlElement stylesheet)
        {
            var fonts1 = new Fonts {Count = 2U, KnownFonts = true};

            var font1 = new Font();
            var fontSize1 = new FontSize {Val = 11D};
            var color1 = new Color {Theme = 1U};
            var fontName1 = new FontName {Val = "Calibri"};
            var fontFamilyNumbering1 = new FontFamilyNumbering {Val = 2};
            var fontScheme1 = new FontScheme {Val = FontSchemeValues.Minor};

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            var font2 = new Font();
            var bold1 = new Bold();
            var fontSize2 = new FontSize {Val = 11D};
            var color2 = new Color {Theme = 1U};
            var fontName2 = new FontName {Val = "Calibri"};
            var fontFamilyNumbering2 = new FontFamilyNumbering {Val = 2};
            var fontScheme2 = new FontScheme {Val = FontSchemeValues.Minor};

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);
            stylesheet.Append(fonts1);
        }

        private static void GenerateBorders(OpenXmlElement stylesheet)
        {
            var borders1 = new Borders {Count = 2U};

            var border1 = new Border();
            var leftBorder1 = new LeftBorder();
            var rightBorder1 = new RightBorder();
            var topBorder1 = new TopBorder();
            var bottomBorder1 = new BottomBorder();
            var diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            var border2 = new Border();
            var leftBorder2 = new LeftBorder();
            var rightBorder2 = new RightBorder();
            var topBorder2 = new TopBorder();

            var bottomBorder2 = new BottomBorder {Style = BorderStyleValues.Thin};
            var color1 = new Color {Indexed = 64U};

            bottomBorder2.Append(color1);
            var diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);
            stylesheet.Append(borders1);
        }

        private static void GenerateCellStyleFormats(OpenXmlElement stylesheet)
        {
            var cellStyleFormats1 = new CellStyleFormats {Count = 1U};
            var cellFormat1 = new CellFormat {NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U};
            cellStyleFormats1.Append(cellFormat1);
            stylesheet.Append(cellStyleFormats1);
        }

        private static void GenerateCellFormats(OpenXmlElement stylesheet)
        {
            var cellFormats1 = new CellFormats {Count = 4U};
            var cellFormat1 = new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U
            };
            var cellFormat2 = new CellFormat
            {
                NumberFormatId = 164U,
                FontId = 1U,
                FillId = 0U,
                BorderId = 1U,
                FormatId = 0U,
                ApplyNumberFormat = true,
                ApplyFont = true,
                ApplyBorder = true
            };
            var cellFormat3 = new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 1U,
                FillId = 0U,
                BorderId = 1U,
                FormatId = 0U,
                ApplyFont = true,
                ApplyBorder = true
            };
            var cellFormat4 = new CellFormat
            {
                NumberFormatId = 164U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyNumberFormat = true
            };
            var cellFormat5 = new CellFormat
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyProtection = true,
                Protection = new Protection {Locked = false}
            };
            var alignment1 = new Alignment {Horizontal = HorizontalAlignmentValues.Left};
            cellFormat4.Append(alignment1);

            cellFormats1.Append(cellFormat1);
            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            stylesheet.Append(cellFormats1);
        }

        private static void GenerateCellStyles(OpenXmlElement stylesheet)
        {
            var cellStyles1 = new CellStyles {Count = 1U};
            var cellStyle1 = new CellStyle {Name = "Normal", FormatId = 0U, BuiltinId = 0U};
            cellStyles1.Append(cellStyle1);
            stylesheet.Append(cellStyles1);
        }
    }
}