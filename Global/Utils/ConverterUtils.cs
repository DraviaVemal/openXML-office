// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Text;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Converter Utils
	/// </summary>
	public static class ConverterUtils
	{


		/// <summary>
		/// This function converts an Excel-style cell reference (e.g., "A1") into row and column
		/// indices (non zero-based) to identify the corresponding cell within a worksheet
		/// </summary>
		/// <returns>row,col</returns>
		public static (int, int) ConvertFromExcelCellReference(string cellReference)
		{
			if (string.IsNullOrEmpty(cellReference)) { throw new ArgumentException("Cell reference cannot be empty."); }
			StringBuilder columnName = new();
			int rowIndex = 0;
			int columnIndex = 0;
			foreach (char c in cellReference)
			{
				if (char.IsLetter(c))
				{
					columnName.Append(c);
				}
				else if (char.IsDigit(c))
				{
					rowIndex = (rowIndex * 10) + (c - '0');
				}
				else
				{
					throw new ArgumentException("Invalid character in cell reference.");
				}
			}
			for (int i = 0; i < columnName.Length; i++)
			{
				columnIndex = (columnIndex * 26) + columnName[i] - 'A' + 1;
			}
			if (rowIndex < 1 || columnIndex < 1)
			{
				throw new ArgumentException("Invalid row or column index in cell reference.");
			}
			return (rowIndex, columnIndex);
		}

		/// <summary>
		/// Converts an integer representing an Excel column index to its corresponding column name.
		/// </summary>
		public static string ConvertIntToColumnName(int column)
		{
			if (column < 1) { throw new ArgumentException("Column indices must be positive integers."); }
			int dividend = column;
			string columnName = string.Empty;
			while (dividend > 0)
			{
				int modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo) + columnName;
				dividend = (dividend - modulo) / 26;
			}
			return columnName;
		}

		/// <summary>
		/// This function converts a pair of row and column indices (non zero-based) into an
		/// Excel-style cell reference (e.g., "A1" for row 1, column 1)
		/// </summary>
		public static string ConvertToExcelCellReference(int row, int column)
		{
			if (row < 1 || column < 1) { throw new ArgumentException("Row and column indices must be positive integers."); }
			return ConvertIntToColumnName(column) + row;
		}

		/// <summary>
		/// Convert Emu to Pixels
		/// </summary>
		public static int EmuToPixels(long emuValue)
		{
			return (int)Math.Round((double)emuValue / 914400 * GlobalConstants.defaultDPI);
		}

		/// <summary>
		/// Convert Pixels to Emu
		/// </summary>
		public static long PixelsToEmu(int pixels)
		{
			return (long)Math.Round((double)pixels / GlobalConstants.defaultDPI * 914400);
		}

		/// <summary>
		/// Convert Point to Emu
		/// </summary>
		public static long PointToEmu(double point)
		{
			return (long)Math.Round(point * GlobalConstants.defaultPointToEmu);
		}

		/// <summary>
		/// Convert Emu to Point
		/// </summary>
		public static int EmuToPoint(long emu)
		{
			return (int)Math.Round((double)emu / GlobalConstants.defaultPointToEmu);
		}

		/// <summary>
		/// TODO: Make Meaningful name
		/// </summary>
		public static int FontSizeToFontSize(float fontSize)
		{
			return (int)fontSize * 100;


		}
	}
}
