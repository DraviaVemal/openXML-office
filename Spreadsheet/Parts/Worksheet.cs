// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2013;
using X = DocumentFormat.OpenXml.Spreadsheet;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	/// Represents a worksheet in an Excel workbook.
	/// </summary>
	public class Worksheet : Drawing
	{
		private readonly Excel excel;
		private readonly X.Worksheet openXMLworksheet;
		private readonly X.Sheet sheet;

		/// <summary>
		/// Initializes a new instance of the <see cref="Worksheet"/> class.
		/// </summary>
		internal Worksheet(Excel excel, X.Worksheet worksheet, X.Sheet _sheet)
		{
			this.excel = excel;
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
		///
		/// </summary>
		internal X.Worksheet GetWorksheet()
		{
			return openXMLworksheet;
		}

		internal X.SheetData GetWorkSheetData()
		{
			return openXMLworksheet.Elements<X.SheetData>().First();
		}

		internal WorksheetPart GetWorksheetPart()
		{
			return openXMLworksheet.WorksheetPart!;
		}

		internal string GetNextSheetPartRelationId()
		{
			return string.Format("rId{0}", GetWorksheetPart().Parts.Count() + 1);
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
		public void SetColumn(string cellId, ColumnProperties columnProperties)
		{
			(int _, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cellId);
			SetColumn(colIndex, columnProperties);
		}

		/// <summary>
		/// Sets the properties for a column at the specified column index in a worksheet.
		/// </summary>
		public void SetColumn(int col, ColumnProperties columnProperties)
		{
			X.Columns? columns = openXMLworksheet.GetFirstChild<X.Columns>();
			if (columns == null)
			{
				columns = new X.Columns();
				openXMLworksheet.InsertBefore(columns, openXMLworksheet.GetFirstChild<X.SheetData>());
			}
			X.Column? existingColumn = columns.Elements<X.Column>().FirstOrDefault(c => c.Max?.Value == col && c.Min?.Value == col);
			if (existingColumn != null)
			{
				existingColumn.CustomWidth = true;
				if (columnProperties != null)
				{
					if (columnProperties.width != null && !columnProperties.bestFit) { existingColumn.Width = columnProperties.width; }
					existingColumn.Hidden = columnProperties.hidden;
					existingColumn.BestFit = BooleanValue.FromBoolean(columnProperties.bestFit);
				}
			}
			else
			{
				X.Column newColumn = new()
				{
					Min = (uint)col,
					Max = (uint)col,
				};
				if (columnProperties != null)
				{
					if (columnProperties.width != null && !columnProperties.bestFit)
					{
						newColumn.Width = columnProperties.width;
						newColumn.CustomWidth = true;
					}
					newColumn.Hidden = columnProperties.hidden;
					newColumn.BestFit = columnProperties.bestFit;
				}
				columns.Append(newColumn);
			}
		}

		/// <summary>
		/// Sets the data and properties for a specific row and its cells in a worksheet.
		/// </summary>
		public void SetRow(int row, int col, DataCell[] dataCells, RowProperties rowProperties)
		{
			SetRow(ConverterUtils.ConvertToExcelCellReference(row, col), dataCells, rowProperties);
		}

		/// <summary>
		/// Sets the data and properties for a row based on a starting cell ID and its data cells in
		/// a worksheet.
		/// </summary>
		public void SetRow(string cellId, DataCell[] dataCells, RowProperties rowProperties)
		{
			(int rowIndex, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cellId);
			X.Row? row = GetWorkSheetData().Elements<X.Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
			if (row == null)
			{
				row = new X.Row
				{
					RowIndex = new UInt32Value((uint)rowIndex)
				};
				GetWorkSheetData().AppendChild(row);
			}
			if (rowProperties != null)
			{
				if (rowProperties.height != null)
				{
					row.Height = rowProperties.height;
					row.CustomHeight = true;
				}
				row.Hidden = rowProperties.hidden;
			}
			foreach (DataCell DataCell in dataCells)
			{
				string currentCellId = ConverterUtils.ConvertToExcelCellReference(rowIndex, colIndex);
				colIndex++;
				X.Cell? cell = row.Elements<X.Cell>().FirstOrDefault(c => c.CellReference?.Value == currentCellId);
				if (string.IsNullOrEmpty(DataCell?.cellValue))
				{
					cell?.Remove();
				}
				else
				{
					if (cell == null)
					{
						cell = new X.Cell
						{
							CellReference = currentCellId
						};
						row.AppendChild(cell);
					}
					X.CellValues dataType = GetCellValueType(DataCell.dataType);
					cell.StyleIndex = DataCell.styleId ?? Styles.Instance.GetCellStyleId(DataCell.styleSetting ?? new());
					if (dataType == X.CellValues.String)
					{
						cell.DataType = X.CellValues.SharedString;
						cell.CellValue = new X.CellValue(ShareString.Instance.InsertUnique(DataCell.cellValue));
					}
					else
					{
						cell.DataType = dataType;
						cell.CellValue = new X.CellValue(DataCell.cellValue);
					}
				}
			}
			openXMLworksheet.Save();
		}

		/// <summary>
		/// Gets the CellValues enumeration corresponding to the specified cell data type.
		/// </summary>
		private static X.CellValues GetCellValueType(CellDataType cellDataType)
		{
			return cellDataType switch
			{
				CellDataType.DATE => X.CellValues.Date,
				CellDataType.NUMBER => X.CellValues.Number,
				_ => X.CellValues.String,
			};
		}

		private static DataType GetCellDataType(EnumValue<X.CellValues>? cellValueType)
		{
			if (cellValueType == null)
			{
				return DataType.STRING;
			}
			return cellValueType.ToString() switch
			{
				"d" => DataType.DATE,
				"n" => DataType.NUMBER,
				_ => DataType.STRING,
			};
		}

		/// <summary>
		///
		/// </summary>
		public void AddPicture(string filePath, ExcelPictureSetting pictureSetting)
		{
			AddPicture(new FileStream(filePath, FileMode.Open, FileAccess.Read), pictureSetting);
		}

		/// <summary>
		///
		/// </summary>
		public Picture? AddPicture(Stream stream, ExcelPictureSetting pictureSetting)
		{
			if (pictureSetting.fromCol < pictureSetting.toCol || pictureSetting.fromRow < pictureSetting.toRow)
			{
				return new Picture(this, stream, new()
				{
					fromCol = pictureSetting.fromCol,
					fromRow = pictureSetting.fromRow,
					toCol = pictureSetting.toCol,
					toRow = pictureSetting.toRow,
				});
			}
			throw new ArgumentException("At least one cell range must be covered by the picture.");
		}

		internal DrawingsPart GetDrawingsPart()
		{
			return GetDrawingsPart(this);
		}

		internal XDR.WorksheetDrawing GetDrawing()
		{
			return GetDrawing(this);
		}

		/// <summary>
		/// 
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName ??= GetSheetName();
			return new(this, chartDatas, dataRange, areaChartSetting);
		}

		/// <summary>
		/// 
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName ??= GetSheetName();
			return new(this, chartDatas, dataRange, barChartSetting);
		}

		/// <summary>
		/// 
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName ??= GetSheetName();
			return new(this, chartDatas, dataRange, columnChartSetting);
		}

		/// <summary>
		/// 
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName ??= GetSheetName();
			return new(this, chartDatas, dataRange, lineChartSetting);
		}

		/// <summary>
		/// 
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName ??= GetSheetName();
			return new(this, chartDatas, dataRange, pieChartSetting);
		}

		/// <summary>
		/// 
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName ??= GetSheetName();
			return new(this, chartDatas, dataRange, scatterChartSetting);
		}

		private ChartData[][] PrepareCacheData(DataRange dataRange)
		{
			Worksheet? worksheet = excel.GetWorksheet(dataRange.sheetName ?? GetSheetName()) ?? throw new ArgumentException("Data Range Sheet not founds");
			(int rowStart, int colStart) = ConverterUtils.ConvertFromExcelCellReference(dataRange.cellIdStart);
			(int rowEnd, int colEnd) = ConverterUtils.ConvertFromExcelCellReference(dataRange.cellIdEnd);
			List<X.Row> dataRows = GetWorkSheetData().Elements<X.Row>().Where(row => row.RowIndex?.Value >= rowStart && row.RowIndex?.Value <= rowEnd).ToList();
			ChartData[][] chartDatas = new ChartData[rowEnd - rowStart + 1][];
			dataRows.ForEach(row =>
			{
				chartDatas[(int)(row.RowIndex?.Value - rowStart)!] = new ChartData[colEnd - colStart + 1];
				List<string> cellIds = new();
				for (int col = colStart; col <= colEnd; col++)
				{
					cellIds.Add(ConverterUtils.ConvertToExcelCellReference((int)(row.RowIndex?.Value)!, col));
				}
				List<X.Cell> dataCells = row.Elements<X.Cell>().Where(c => cellIds.Contains(c.CellReference?.Value!)).ToList();
				dataCells.ForEach(cell =>
				{
					(int _, int colIndex) = ConverterUtils.ConvertFromExcelCellReference(cell.CellReference?.Value!);
					// TODO : Cell Value is bit confusing for value types and formula do furter research for extending the functionality
					DataType cellDataType = GetCellDataType(cell.DataType);
					string? cellValue = cellDataType switch
					{
						_ => cell.CellValue!.Text
					};
					if (cell.DataType?.ToString() == "s")
					{
						cellValue = ShareString.Instance.GetValue(int.Parse(cellValue));
					}
					chartDatas[(int)(row.RowIndex?.Value - rowStart)!][colIndex - colStart] = new()
					{
						dataType = cellDataType,
						numberFormat = Styles.Instance.GetStyleForId(cell.StyleIndex!).numberFormat,
						value = cellValue ?? ""
					};
				});
			});
			return chartDatas;
		}
	}
}
