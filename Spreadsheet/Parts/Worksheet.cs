// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2007;
using X = DocumentFormat.OpenXml.Spreadsheet;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
namespace OpenXMLOffice.Spreadsheet_2007
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
			return int.Parse(sheet.Id.Value);
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
			return openXMLworksheet.WorksheetPart;
		}
		internal string GetNextSheetPartRelationId()
		{
			return string.Format("rId{0}", GetWorksheetPart().Parts.Count() + 1);
		}
		internal string GetNextDrawingPartRelationId()
		{
			return string.Format("rId{0}", GetDrawingsPart().Parts.Count() + 1);
		}
		/// <summary>
		/// Returns the sheet name of the current worksheet.
		/// </summary>
		public string GetSheetName()
		{
			return sheet.Name;
		}
		/// <summary>
		/// Sets the properties for a column based on a starting cell ID in a worksheet.
		/// </summary>
		public void SetColumn(string cellId, ColumnProperties columnProperties)
		{
			Tuple<int, int> result = ConverterUtils.ConvertFromExcelCellReference(cellId);
			SetColumn(result.Item2, columnProperties);
		}
		/// <summary>
		/// Sets the properties for a column at the specified column index in a worksheet.
		/// </summary>
		public void SetColumn(int col, ColumnProperties columnProperties)
		{
			X.Columns columns = openXMLworksheet.GetFirstChild<X.Columns>();
			if (columns == null)
			{
				columns = new X.Columns();
				openXMLworksheet.InsertBefore(columns, openXMLworksheet.GetFirstChild<X.SheetData>());
			}
			X.Column existingColumn = columns.Elements<X.Column>().FirstOrDefault(c => c.Max.Value == col && c.Min.Value == col);
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
				X.Column newColumn = new X.Column()
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
			Tuple<int, int> result = ConverterUtils.ConvertFromExcelCellReference(cellId);
			int rowIndex = result.Item1;
			int columnIndex = result.Item2;
			X.Row row = GetWorkSheetData().Elements<X.Row>().FirstOrDefault(r => r.RowIndex.Value == (uint)rowIndex);
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
				if (DataCell != null)
				{
					string currentCellId = ConverterUtils.ConvertToExcelCellReference(rowIndex, columnIndex);
					columnIndex++;
					X.Cell cell = row.Elements<X.Cell>().FirstOrDefault(c => c.CellReference.Value == currentCellId);
					if (string.IsNullOrEmpty(DataCell.cellValue))
					{
						cell.Remove();
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
						cell.StyleIndex = DataCell.styleId ?? excel.GetStyleService().GetCellStyleId(DataCell.styleSetting ?? new CellStyleSetting());
						if (dataType == X.CellValues.String)
						{
							cell.DataType = X.CellValues.SharedString;
							cell.CellValue = new X.CellValue(excel.GetShareStringService().InsertUnique(DataCell.cellValue));
						}
						else
						{
							cell.DataType = dataType;
							cell.CellValue = new X.CellValue(DataCell.cellValue);
						}
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
			switch (cellDataType)
			{
				case CellDataType.DATE:
					return X.CellValues.Date;
				case CellDataType.NUMBER:
					return X.CellValues.Number;
				default:
					return X.CellValues.String;
			}
		}

		private static DataType GetCellDataType(EnumValue<X.CellValues> cellValueType)
		{
			if (cellValueType == null)
			{
				return DataType.STRING;
			}
			else
			{
				string valueType = cellValueType.ToString();
				switch (valueType)
				{
					case "d":
						return DataType.DATE;
					case "n":
						return DataType.NUMBER;
					default:
						return DataType.STRING;
				}
			}
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
		public Picture AddPicture(Stream stream, ExcelPictureSetting pictureSetting)
		{
			if (pictureSetting.fromCol < pictureSetting.toCol || pictureSetting.fromRow < pictureSetting.toRow)
			{
				return new Picture(this, stream, new ExcelPictureSetting()
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
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, areaChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, barChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, columnChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, lineChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, pieChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting) where ApplicationSpecificSetting : ExcelSetting
		{
			ChartData[][] chartDatas = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartDatas, dataRange, comboChartSetting);
		}
		private ChartData[][] PrepareCacheData(DataRange dataRange)
		{
			string sheetName = dataRange.sheetName ?? GetSheetName();
			Worksheet worksheet = excel.GetWorksheet(sheetName);
			if (worksheet == null)
			{
				throw new ArgumentException("Data Range Sheet not found");
			}
			Tuple<int, int> result = ConverterUtils.ConvertFromExcelCellReference(dataRange.cellIdStart);
			int rowStart = result.Item1;
			int colStart = result.Item2;
			result = ConverterUtils.ConvertFromExcelCellReference(dataRange.cellIdEnd);
			int rowEnd = result.Item1;
			int colEnd = result.Item2;
			List<X.Row> dataRows = GetWorkSheetData().Elements<X.Row>().Where(row => row.RowIndex.Value >= rowStart && row.RowIndex.Value <= rowEnd).ToList();
			ChartData[][] chartDatas = new ChartData[rowEnd - rowStart + 1][];
			dataRows.ForEach(row =>
			{
				chartDatas[(int)(row.RowIndex.Value - rowStart)] = new ChartData[colEnd - colStart + 1];
				List<string> cellIds = new List<string>();
				for (int col = colStart; col <= colEnd; col++)
				{
					cellIds.Add(ConverterUtils.ConvertToExcelCellReference((int)row.RowIndex.Value, col));
				}
				List<X.Cell> dataCells = row.Elements<X.Cell>().Where(c => cellIds.Contains(c.CellReference.Value)).ToList();
				dataCells.ForEach(cell =>
				{
					result = ConverterUtils.ConvertFromExcelCellReference(cell.CellReference.Value);
					int colIndex = result.Item2;
					// TODO : Cell Value is bit confusing for value types and formula do furter research for extending the functionality
					DataType cellDataType = GetCellDataType(cell.DataType);
					string cellValue;
					switch (cellDataType)
					{
						default:
							cellValue = cell.CellValue.Text;
							break;
					}
					if (cell.DataType.ToString() == "s")
					{
						cellValue = excel.GetShareStringService().GetValue(int.Parse(cellValue));
					}
					chartDatas[(int)(row.RowIndex.Value - rowStart)][colIndex - colStart] = new ChartData()
					{
						dataType = cellDataType,
						// TODO : Do Performance Update
						numberFormat = excel.GetStyleService().GetStyleForId(cell.StyleIndex).numberFormat,
						value = cellValue ?? ""
					};
				});
			});
			return chartDatas;
		}
	}
}
