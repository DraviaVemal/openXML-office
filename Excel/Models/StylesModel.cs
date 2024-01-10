namespace OpenXMLOffice.Excel
{
    public class CellStyleSetting
    {
        public enum VerticalAlignmentValues
        {
            TOP,
            MIDDLE,
            BOTTOM
        }
        public enum HorizontalAlignmentValues
        {
            LEFT,
            CENTER,
            RIGHT
        }
        public string FontFamily = "Calibri";
        public int FontSize = 11;
        public bool IsBold = false;
        public bool IsItalic = false;
        public bool IsStrick = false;
        public bool IsDoubleStrick = false;
        public bool IsWrapText = false;
        public string TextColor = "000000";
        public string? BackgroundColor;
        public VerticalAlignmentValues VerticalAlignment = VerticalAlignmentValues.BOTTOM;
        public HorizontalAlignmentValues HorizontalAlignment = HorizontalAlignmentValues.LEFT;
        public string NumberFormat = "General";
    }
    public class BorderBase
    {
        public enum StyleValues
        {
            THIN
        }
        public StyleValues Style = StyleValues.THIN;
        public string Color = "64";
    }
    public class FontStyle
    {
        public int Id { get; set; }
        public string Size = "11";
        public string Color = "accent1";
        public string Family = "2";
        public string name = "Calibri";
    }

    public class FillStyle
    {
        public int Id { get; set; }
    }

    public class BorderStyle
    {
        public class LeftBorder : BorderBase
        {

        }
        public class RightBorder : BorderBase
        {

        }
        public class TopBorder : BorderBase
        {

        }
        public class BottomBorder : BorderBase
        {

        }
        public int Id { get; set; }
        public LeftBorder Left = new();
        public RightBorder Right = new();
        public TopBorder Top = new();
        public BottomBorder Bottom = new();
    }
}