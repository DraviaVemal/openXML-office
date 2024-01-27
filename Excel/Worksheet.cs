// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLOffice.Global_2013;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLOffice.Excel_2013
{
	/// <summary>
	/// Represents a worksheet in an Excel workbook.
	/// </summary>
	public class Worksheet
	{
		private readonly S.Worksheet openXMLworksheet;
		private readonly Sheet sheet;

		/// <summary>
		/// Initializes a new instance of the <see cref="Worksheet"/> class.
		/// </summary>
		public Worksheet(S.Worksheet worksheet, Sheet _sheet)
		{
			openXMLworksheet = worksheet;
			sheet = _sheet;
		}


		/// <summary>
		/// Returns the sheet ID of the current worksheet.
		/// </summary>
		public int GetSheetId()
		{
			return int.Parse(sheet.Id!.Value!);
		}

		/// <summary>
		/// Returns the sheet name of the current worksheet.
		/// </summary>
		public string GetSheetName()
		{
			return sheet.Name!;
		}

		/// <summary>
		/// Sets the properties for a column based on a starting cell ID in a worksheet.
		/// </summary>
		public void SetColumn(string cellId, ColumnProperties ColumnProperties)
		{
			(int _, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cellId);
			SetColumn(colIndex, ColumnProperties);
		}

		/// <summary>
		/// Sets the properties for a column at the specified column index in a worksheet.
		/// </summary>
		public void SetColumn(int col, ColumnProperties ColumnProperties)
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
				if (ColumnProperties != null)
				{
					if (ColumnProperties.width != null && !ColumnProperties.bestFit) { existingColumn.Width = ColumnProperties.width; }
					existingColumn.Hidden = ColumnProperties.hidden;
					existingColumn.BestFit = BooleanValue.FromBoolean(ColumnProperties.bestFit);
				}
			}
			else
			{
				Column newColumn = new()
				{
					Min = (uint)col,
					Max = (uint)col,
				};
				if (ColumnProperties != null)
				{
					if (ColumnProperties.width != null && !ColumnProperties.bestFit)
					{
						newColumn.Width = ColumnProperties.width;
						newColumn.CustomWidth = true;
					}
					newColumn.Hidden = ColumnProperties.hidden;
					newColumn.BestFit = ColumnProperties.bestFit;
				}
				columns.Append(newColumn);
			}
		}

		/// <summary>
		/// Sets the data and properties for a specific row and its cells in a worksheet.
		/// </summary>
		public void SetRow(int row, int col, DataCell[] dataCells, RowProperties RowProperties)
		{
			SetRow(ConverterUtils.ConvertToExcelCellReference(row, col), dataCells, RowProperties);
		}

		/// <summary>
		/// Sets the data and properties for a row based on a starting cell ID and its data cells in
		/// a worksheet.
		/// </summary>
		public void SetRow(string cellId, DataCell[] DataCells, RowProperties RowProperties)
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
			if (RowProperties != null)
			{
				if (RowProperties.height != null)
				{
					row.Height = RowProperties.height;
					row.CustomHeight = true;
				}
				row.Hidden = RowProperties.hidden;
			}
			foreach (DataCell DataCell in DataCells)
			{
				string currentCellId = ConverterUtils.ConvertToExcelCellReference(rowIndex, colIndex);
				colIndex++;
				Cell? cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == currentCellId);
				if (string.IsNullOrEmpty(DataCell?.cellValue))
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
					CellValues DataType = GetCellValues(DataCell.dataType);
					cell.StyleIndex = DataCell.styleId ?? Styles.Instance.GetCellStyleId(DataCell.styleSetting ?? new());
					if (DataType == CellValues.String)
					{
						cell.DataType = CellValues.SharedString;
						cell.CellValue = new CellValue(ShareString.Instance.InsertUnique(DataCell.cellValue));
					}
					else
					{
						cell.DataType = DataType;
						cell.CellValue = new CellValue(DataCell.cellValue);
					}
				}
			}
			openXMLworksheet.Save();
		}

		/// <summary>
		/// Gets the CellValues enumeration corresponding to the specified cell data type.
		/// </summary>
		/// <param name="cellDataType">
		/// The data type of the cell.
		/// </param>
		/// <returns>
		/// The CellValues enumeration representing the cell data type.
		/// </returns>
		private static CellValues GetCellValues(CellDataType cellDataType)
		{
			return cellDataType switch
			{
				CellDataType.DATE => CellValues.Date,
				CellDataType.NUMBER => CellValues.Number,
				_ => CellValues.String,
			};
		}


	}
}
