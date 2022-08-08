using DocumentFormat.OpenXml.Spreadsheet;

namespace SdiExcelExport;

public sealed class ExcelExportStyles: ExcelStyles
{
    public override Stylesheet GenerateStyleSheet()
    {
        Fonts fonts = GetFonts();
        Fills fills = GetFills();
        Borders borders = GetBorders();
        
        CellFormats cellFormats = new();
        cellFormats.InsertAt(new CellFormat(), 0);
        cellFormats.InsertAt(new CellFormat {FontId = 1, FillId = 2, BorderId = 0, ApplyFill = true, NumberFormatId = 0}, 1);
        cellFormats.InsertAt(new CellFormat {FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true, NumberFormatId = 0}, 2);
        cellFormats.InsertAt(new CellFormat {FontId = 0, FillId = 3, BorderId = 0, ApplyFill = true, NumberFormatId = 0}, 3);
        Stylesheet styleSheet = new(fonts, fills, borders, cellFormats);
        return styleSheet;
    }
}