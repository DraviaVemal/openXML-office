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
        public static X.DataCell[][] CreateDataCellPayload(int colSize = 5, int rowSize = 5, bool IsValueAxis = false)
        {
            Random random = new();
            X.DataCell[][] data = new X.DataCell[rowSize][];
            data[0] = new X.DataCell[colSize];
            for (int col = 1; col < colSize; col++)
            {
                data[0][col] = new X.DataCell
                {
                    cellValue = $"Series {col}",
                    dataType = X.CellDataType.STRING
                };
            }
            for (int row = 1; row < rowSize; row++)
            {
                data[row] = new X.DataCell[colSize];
                data[row][0] = new X.DataCell
                {
                    cellValue = $"Category {row}",
                    dataType = X.CellDataType.STRING
                };
                for (int col = IsValueAxis ? 0 : 1; col < colSize; col++)
                {
                    data[row][col] = new X.DataCell
                    {
                        cellValue = random.Next(1, 100).ToString(),
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