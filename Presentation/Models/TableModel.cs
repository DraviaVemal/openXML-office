namespace OpenXMLOffice.Presentation
{

    public class TableCell
    {
        public string? Value;
    }

    public class TableRow
    {
        public string Background = "FFFFFF";
        public string TextColor = "000000";
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