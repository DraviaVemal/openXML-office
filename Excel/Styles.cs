using LiteDB;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel
{
    public class Styles
    {
        private readonly Stylesheet Stylesheet;
        private static readonly LiteDatabase LiteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
        private readonly ILiteCollection<FontStyle> FontStyleCollection;
        private readonly ILiteCollection<FillStyle> FillStyleCollection;
        private readonly ILiteCollection<BorderStyle> BorderStyleCollection;
        internal Styles(Stylesheet Stylesheet)
        {
            this.Stylesheet = Stylesheet;
            FontStyleCollection = LiteDatabase.GetCollection<FontStyle>("FontStyle");
            FillStyleCollection = LiteDatabase.GetCollection<FillStyle>("FillStyle");
            BorderStyleCollection = LiteDatabase.GetCollection<BorderStyle>("BorderStyle");
            Initialise();
        }

        private void Initialise()
        {
            Stylesheet.Fonts ??= new(
                new Font(
                    new FontSize() { Val = 11 },
                    new Color() { Theme = 1 },
                    new FontName() { Val = "Calibri" },
                    new FontFamily() { Val = 2 },
                    new FontScheme() { Val = FontSchemeValues.Minor }
                ))
            { Count = 1 };
            Stylesheet.Fills ??= new(
                new Fill(
                    new PatternFill() { PatternType = PatternValues.None }
                ),
                new Fill(
                    new PatternFill() { PatternType = PatternValues.DarkGray }
                ))
            { Count = 2 };
            Stylesheet.Borders ??= new(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            )
            { Count = 1 };
            Stylesheet.CellStyleFormats ??= new(
                new CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 })
            { Count = 1 };//cellStyleXfs
            Stylesheet.CellFormats ??= new(
                new CellFormat() { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 })
            { Count = 1 };//cellXfs
            Stylesheet.CellStyles ??= new(
                new CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 })
            { Count = 1 };//cellStyles
            Stylesheet.DifferentialFormats ??= new() { Count = 0 };//dxfs
        }

        public int GetCellStyleId(CellStyleSetting CellStyleSetting)
        {
            return 0;
        }

        private void AddUniqueFont(CellStyleSetting CellStyleSetting)
        {
            Font Font = new();
        }
    }
}