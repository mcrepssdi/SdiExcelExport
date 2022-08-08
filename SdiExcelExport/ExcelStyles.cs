using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SdiExcelExport;

public abstract class ExcelStyles
{

    public abstract Stylesheet GenerateStyleSheet();
    
    protected static Fonts GetFonts()
    {
        Fonts fonts = new (
            new Font(new FontSize { Val = 10 }, new FontName{Val = "Calibri"}),
            new Font (new FontSize { Val = 11 }, new Bold(), new FontName{Val = "Segoe UI"})
        );
        return fonts;
    }

    protected static Fills GetFills()
    {
        Fills fills = new (
            new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
            new Fill(new PatternFill ( new ForegroundColor { Rgb = new HexBinaryValue { Value = "D9D9D9" } }) { PatternType = PatternValues.Solid }), // Index 1 - default
            new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue { Value = "D9D9D9" } }) { PatternType = PatternValues.Solid }), // Index 2 - header
            new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue { Value = "B0C4DE" } }) { PatternType = PatternValues.Solid })
        );

        return fills;
    }

    protected static Borders GetBorders()
    {
        Borders borders = new (
            new Border(), // index 0 default
            new Border( // index 1 black border
                new LeftBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thick },
                new BottomBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder()
            ),
            new Border( // index 2 thick black top border
                new TopBorder(new Color { Auto = true }) { Style = BorderStyleValues.Thick }
            )
        );

        return borders;
    }

    protected static NumberingFormat GetCurrenyFormat()
    {
        NumberingFormat currencyFormat = new ()
        {
            NumberFormatId = UInt32Value.FromUInt32(3453),
            FormatCode = StringValue.FromString("[$$-en-US] #,##0.00")
        };
        return currencyFormat;
    }
}