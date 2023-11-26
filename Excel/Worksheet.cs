using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel;

public class Worksheet
{
    private readonly S.Worksheet openXMLworksheet;
    public Sheet sheet;
    public Worksheet(S.Worksheet worksheet, Sheet _sheet)
    {
        openXMLworksheet = worksheet;
        sheet = _sheet;
    }

    private CellValues GetCellValues(CellDataType cellDataType)
    {
        switch (cellDataType)
        {
            case CellDataType.DATE:
                return CellValues.Date;
            case CellDataType.NUMBER:
                return CellValues.Number;
            default:
                return CellValues.String;
        }
    }

    public int GetSheetId()
    {
        return int.Parse(sheet.Id!.Value!);
    }

    public string GetSheetName()
    {
        return sheet.Name!;
    }

    public void SetRow(int row, int col, DataCell[] dataCells, RowProperties? rowProperties = null)
    {
        SetRow(ConverterUtils.ConvertToExcelCellReference(row, col), dataCells, rowProperties);
    }

    public void SetRow(string cellId, DataCell[] dataCells, RowProperties? rowProperties = null)
    {
        SheetData sheetData = openXMLworksheet.Elements<SheetData>().First();
        (int rowIndex, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cellId);
        Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        if (row == null)
        {
            row = new Row
            {
                RowIndex = new UInt32Value((uint)rowIndex)
            };
            sheetData.AppendChild(row);
        }
        if (rowProperties != null)
        {
            if (rowProperties.height != null)
            {
                row.Height = rowProperties.height;
                row.CustomHeight = true;
            }
            row.Hidden = rowProperties.Hidden;
        }
        foreach (DataCell dataCell in dataCells)
        {
            string currentCellId = ConverterUtils.ConvertToExcelCellReference(rowIndex, colIndex);
            colIndex++;
            Cell? cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == currentCellId);
            if (string.IsNullOrEmpty(dataCell.CellValue))
            {
                cell?.Remove();
            }
            else
            {
                if (cell == null)
                {
                    cell = new Cell
                    {
                        CellReference = currentCellId
                    };
                    row.AppendChild(cell);
                }
                cell.DataType = GetCellValues(dataCell.DataType);
                cell.CellValue = new CellValue(dataCell.CellValue);
            }
        }
        openXMLworksheet.Save();
    }

    public void SetColumn(string cellId, ColumnProperties? columnProperties = null)
    {
        (int _, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cellId);
        SetColumn(colIndex, columnProperties);
    }

    public void SetColumn(int col, ColumnProperties? columnProperties = null)
    {
        Columns? columns = openXMLworksheet.GetFirstChild<Columns>();
        if (columns == null)
        {
            columns = new Columns();
            openXMLworksheet.InsertBefore(columns, openXMLworksheet.GetFirstChild<SheetData>());
        }
        Column? existingColumn = columns.Elements<Column>().FirstOrDefault(c => c.Max?.Value == col && c.Min?.Value == col);
        if (existingColumn != null)
        {
            existingColumn.CustomWidth = true;
            if (columnProperties != null)
            {
                if (columnProperties.Width != null && !columnProperties.BestFit)
                    existingColumn.Width = columnProperties.Width;
                existingColumn.Hidden = columnProperties.Hidden;
                existingColumn.BestFit = BooleanValue.FromBoolean(columnProperties.BestFit);
            }
        }
        else
        {
            Column newColumn = new()
            {
                Min = (uint)col,
                Max = (uint)col,
            };
            if (columnProperties != null)
            {
                if (columnProperties.Width != null && !columnProperties.BestFit)
                {
                    newColumn.Width = columnProperties.Width;
                    newColumn.CustomWidth = true;
                }
                newColumn.Hidden = columnProperties.Hidden;
                newColumn.BestFit = columnProperties.BestFit;
            }
            columns.Append(newColumn);
        }
    }
}