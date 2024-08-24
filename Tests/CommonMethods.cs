// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using X = OpenXMLOffice.Spreadsheet_2007;
namespace OpenXMLOffice.Tests
{

    /// <summary>
    /// 
    /// </summary>
    public class CommonMethod
    {
        /// <summary>
        /// 
        /// </summary>
        public static X.ColumnCell[][] CreateDataCellPayload(int colSize = 5, int rowSize = 5, bool IsValueAxis = false)
        {
            Random random = new();
            X.ColumnCell[][] data = new X.ColumnCell[rowSize][];
            data[0] = new X.ColumnCell[colSize];
            for (int col = 1; col < colSize; col++)
            {
                data[0][col] = new X.ColumnCell
                {
                    cellValue = $"Series {col}",
                    dataType = X.CellDataType.STRING
                };
            }
            for (int row = 1; row < rowSize; row++)
            {
                data[row] = new X.ColumnCell[colSize];
                data[row][0] = new X.ColumnCell
                {
                    cellValue = $"Category {row}",
                    dataType = X.CellDataType.STRING
                };
                for (int col = IsValueAxis ? 0 : 1; col < colSize; col++)
                {
                    data[row][col] = new X.ColumnCell
                    {
                        cellValue = (row % 2 == 0) ? random.Next(1, 10).ToString() : random.Next(11, 100).ToString(),
                        dataType = X.CellDataType.NUMBER,
                        styleSetting = new()
                        {
                            numberFormat = "0.00",
                            fontSize = 20
                        }
                    };
                }
            }
            return data;
        }
    }

}