namespace OpenXMLOffice.Presentation
{
    public class TableCell
    {
        public bool LeftBorder = false;
        public bool TopBorder = false;
        public bool RightBorder = false;
        public bool BottomBorder = false;
        public string CellBackground = "FFFFFF";
        public string TextBackground = "FFFFFF";
        public string TextColor = "000000";
        public string FontFamily = "Calibri (Body)";
        public int FontSize = 16;
        public bool IsBold = false;
        public bool IsItalic = false;
        public bool IsUnderline = false;
        public string? Value;
    }

    public class TableRow
    {
        public int Height = 370840;
        public List<TableCell> TableCells = new();
    }

    public class TableSetting
    {
        public string Name = "Table 1";
        public enum eWidthType
        {
            AUTO,
            EMU,
            PIXEL,
            PERCENTAGE,
            RATIO
        }
        public eWidthType WidthType = eWidthType.AUTO;
        /// <summary>
        /// AUTO - Ingnore User Width value and space the colum equally
        /// EMU - (English Metric Units) Direct PPT standard Sizing 1 Inch * 914400 EMU's
        /// PIXEL - Based on Target DPI the pixel is converted to EMU and used when running
        /// PERCENTAGE - 0-100 Width percentage split for each column
        /// RATIO - 0-10 Width ratio of each column 
        /// </summary>
        public float Width = 0;
        public List<float> TableColumnwidth = new();
    }
}