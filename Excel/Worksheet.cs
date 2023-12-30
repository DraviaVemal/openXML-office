using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel
{
    public class Worksheet
    {
        #region Public Fields

        public Sheet sheet;

        #endregion Public Fields

        #region Private Fields

        private readonly S.Worksheet openXMLworksheet;

        #endregion Private Fields

        #region Public Constructors

        public Worksheet(S.Worksheet worksheet, Sheet _sheet)
        {
            openXMLworksheet = worksheet;
            sheet = _sheet;
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Return Sheet ID of current Worksheet
        /// </summary>
        /// <returns>
        /// </returns>
        public int GetSheetId()
        {
            return int.Parse(sheet.Id!.Value!);
        }

        /// <summary>
        /// Returb Sheet Name of Current Worksheet
        /// </summary>
        /// <returns>
        /// </returns>
        public string GetSheetName()
        {
            return sheet.Name!;
        }

        /// <summary>
        /// Sets the properties for a column based on a starting cell ID in a worksheet.
        /// </summary>
        /// <param name="cellId">
        /// The cell ID (e.g., "A1") in the desired column.
        /// </param>
        /// <param name="columnProperties">
        /// Optional column properties to be applied (e.g., width, hidden).
        /// </param>
        public void SetColumn(string cellId, ColumnProperties? columnProperties = null)
        {
            (int _, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cellId);
            SetColumn(colIndex, columnProperties);
        }

        /// <summary>
        /// Sets the properties for a column at the specified column index in a worksheet.
        /// </summary>
        /// <param name="col">
        /// The zero-based column index where properties will be applied.
        /// </param>
        /// <param name="columnProperties">
        /// Optional column properties to be applied (e.g., width, hidden).
        /// </param>
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

        /// <summary>
        /// Sets the data and properties for a specific row and its cells in a worksheet.
        /// </summary>
        /// <param name="row">
        /// The row index (non zero-based) where the data and properties will be applied.
        /// </param>
        /// <param name="col">
        /// The starting column index (non zero-based) for adding data cells.
        /// </param>
        /// <param name="dataCells">
        /// An array of data cells to be added to the row.
        /// </param>
        /// <param name="rowProperties">
        /// Optional row properties to be applied to the row (e.g., height, custom formatting).
        /// </param>
        public void SetRow(int row, int col, DataCell[] dataCells, RowProperties? rowProperties = null)
        {
            SetRow(ConverterUtils.ConvertToExcelCellReference(row, col), dataCells, rowProperties);
        }

        /// <summary>
        /// Sets the data and properties for a row based on a starting cell ID and its data cells in
        /// a worksheet.
        /// </summary>
        /// <param name="cellId">
        /// The cell ID (e.g., "A1") from which the row will be determined.
        /// </param>
        /// <param name="dataCells">
        /// An array of data cells to be added to the row.
        /// </param>
        /// <param name="rowProperties">
        /// Optional row properties to be applied to the row (e.g., height, custom formatting).
        /// </param>
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
                if (string.IsNullOrEmpty(dataCell?.CellValue))
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

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Gets the CellValues enumeration corresponding to the specified cell data type.
        /// </summary>
        /// <param name="cellDataType">
        /// The data type of the cell.
        /// </param>
        /// <returns>
        /// The CellValues enumeration representing the cell data type.
        /// </returns>
        private CellValues GetCellValues(CellDataType cellDataType)
        {
            return cellDataType switch
            {
                CellDataType.DATE => CellValues.Date,
                CellDataType.NUMBER => CellValues.Number,
                _ => CellValues.String,
            };
        }

        #endregion Private Methods
    }
}