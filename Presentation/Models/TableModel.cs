namespace OpenXMLOffice.Presentation
{
    public class TableCell
    {
        #region Public Fields

        public bool BottomBorder = false;
        public string? CellBackground;
        public string FontFamily = "Calibri (Body)";
        public int FontSize = 16;
        public bool IsBold = false;
        public bool IsItalic = false;
        public bool IsUnderline = false;
        public bool LeftBorder = false;
        public bool RightBorder = false;
        public string? TextBackground;
        public string TextColor = "000000";
        public bool TopBorder = false;
        public string? Value;

        #endregion Public Fields
    }

    public class TableRow
    {
        #region Public Fields

        public int Height = 370840;
        public List<TableCell> TableCells = new();

        #endregion Public Fields
    }

    public class TableSetting
    {
        #region Public Fields

        public uint Height = 741680;
        public string Name = "Table 1";
        public List<float> TableColumnwidth = new();
        public uint Width = 8128000;

        /// <summary>
        /// AUTO - Ingnore User Width value and space the colum equally EMU - (English Metric Units)
        /// Direct PPT standard Sizing 1 Inch * 914400 EMU's PIXEL - Based on Target DPI the pixel
        /// is converted to EMU and used when running PERCENTAGE - 0-100 Width percentage split for
        /// each column RATIO - 0-10 Width ratio of each column
        /// </summary>
        public eWidthType WidthType = eWidthType.AUTO;

        public uint X = 0;
        public uint Y = 0;

        #endregion Public Fields

        #region Public Enums

        public enum eWidthType
        {
            AUTO,
            EMU,
            PIXEL,
            PERCENTAGE,
            RATIO
        }

        #endregion Public Enums
    }
}