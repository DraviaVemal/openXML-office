namespace OpenXMLOffice.Excel;

// Define a simple model class
public static class ConverterUtils
{
    /// <summary>
    /// This function converts a pair of row and column indices (non zero-based) into an Excel-style cell reference (e.g., "A1" for row 1, column 1)
    /// </summary>
    /// <param name="row"></param>
    /// <param name="column"></param>
    /// <returns>Excel Cell Name</returns>
    /// <exception cref="ArgumentException"></exception>
    public static string ConvertToExcelCellReference(int row, int column)
    {
        if (row < 1 || column < 1)
            throw new ArgumentException("Row and column indices must be positive integers.");
        int dividend = column;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName + row;
    }
    /// <summary>
    /// This function converts an Excel-style cell reference (e.g., "A1") into row and column indices (non zero-based) to identify the corresponding cell within a worksheet
    /// </summary>
    /// <param name="cellReference"></param>
    /// <returns>(int,int) row, col</returns>
    /// <exception cref="ArgumentException"></exception>
    public static (int, int) ConvertFromExcelCellReference(string cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
            throw new ArgumentException("Cell reference cannot be empty.");
        string columnName = string.Empty;
        int rowIndex = 0;
        int columnIndex = 0;
        foreach (char c in cellReference)
        {
            if (char.IsLetter(c))
            {
                columnName += c;
            }
            else if (char.IsDigit(c))
            {
                rowIndex = rowIndex * 10 + (c - '0');
            }
            else
            {
                throw new ArgumentException("Invalid character in cell reference.");
            }
        }
        for (int i = 0; i < columnName.Length; i++)
        {
            columnIndex = columnIndex * 26 + columnName[i] - 'A' + 1;
        }
        if (rowIndex < 1 || columnIndex < 1)
        {
            throw new ArgumentException("Invalid row or column index in cell reference.");
        }
        return (rowIndex, columnIndex);
    }
}