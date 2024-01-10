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
            Stylesheet.Fonts ??= new() { Count = 0 };
            Stylesheet.Fills ??= new() { Count = 0 };
            Stylesheet.Borders ??= new() { Count = 0 };
            Stylesheet.CellFormats ??= new() { Count = 0 };
            Stylesheet.CellStyleFormats ??= new() { Count = 0 };
            Stylesheet.CellStyles ??= new() { Count = 0 };
            Stylesheet.DifferentialFormats ??= new() { Count = 0 };
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