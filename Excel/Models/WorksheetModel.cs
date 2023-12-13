namespace OpenXMLOffice.Excel;

public enum CellDataType
{
    DATE,
    NUMBER,
    STRING
}

public class DataCell
{
    public string? CellValue;
    public CellDataType DataType;
    public string? numberFormatting;
    public int? styleId;
}

public class RowProperties
{
    public double? height;
    public bool Hidden;
}

public class ColumnProperties
{
    public double? Width;
    public bool Hidden;
    public bool BestFit;
}

