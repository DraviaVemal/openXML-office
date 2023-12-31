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
        public string? Value;
    }

    public class TableRow
    {
        public int Height = 370840;
        public List<TableCell> TableCells = new();
    }

    public class TableColumnSetting
    {
        public int Width = -1;
    }

    public class TableSetting
    {
        public string Name = "Table 1";

        public List<TableColumnSetting> TableColumnSettings = new();
    }
}