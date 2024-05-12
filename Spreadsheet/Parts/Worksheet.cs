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
		public string GetSheetId()
		{
			return sheet.Id.Value;
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
			X.SheetData SheetData = openXMLworksheet.Elements<X.SheetData>().FirstOrDefault();
			if (SheetData == null)
			{
				return openXMLworksheet.AppendChild(new X.SheetData());
			}
			return SheetData;
		}
		internal X.Hyperlinks GetWorkSheetHyperlinks()
		{
			X.Hyperlinks hyperlinks = openXMLworksheet.Elements<X.Hyperlinks>().FirstOrDefault();
			return hyperlinks;
		}
		internal X.MergeCells GetWorkSheetMergeCell()
		{
			X.MergeCells mergeCells = openXMLworksheet.Elements<X.MergeCells>().FirstOrDefault();
			return mergeCells;
		}
		internal void AddHyperlink(string relationshipId, string toolTip)
		{
			AddHyperlink(relationshipId, toolTip, null);
		}
		internal void AddHyperlink(string relationshipId, string toolTip, string location)
		{
			X.Hyperlinks hyperlinks = GetWorkSheetHyperlinks() ?? openXMLworksheet.AppendChild(new X.Hyperlinks());
			if (location != null)
			{
				hyperlinks.AppendChild(new X.Hyperlink()
				{
					Reference = relationshipId,
					Tooltip = toolTip,
					Location = location,
				});
			}
			else
			{
				hyperlinks.AppendChild(new X.Hyperlink()
				{
					Reference = relationshipId,
					Tooltip = toolTip,
				});
			}
		}
		internal WorksheetPart GetWorksheetPart()
		{
			return openXMLworksheet.WorksheetPart;
		}
		internal string GetNextSheetPartRelationId()
		{
			return string.Format("rId{0}", GetWorksheetPart().Parts.Count() + GetWorksheetPart().ExternalRelationships.Count() + GetWorksheetPart().HyperlinkRelationships.Count() +
			GetWorksheetPart().DataPartReferenceRelationships.Count() + GetWorksheetPart().Model3DReferenceRelationshipParts.Count() + 1);
		}
		internal string GetNextDrawingPartRelationId()
		{
			return string.Format("rId{0}", GetDrawingsPart().Parts.Count() + GetWorksheetPart().ExternalRelationships.Count() + GetWorksheetPart().HyperlinkRelationships.Count() +
			GetWorksheetPart().DataPartReferenceRelationships.Count() + GetWorksheetPart().Model3DReferenceRelationshipParts.Count() + 1);
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
			foreach (DataCell dataCell in dataCells)
			{
				string currentCellId = ConverterUtils.ConvertToExcelCellReference(rowIndex, columnIndex);
				columnIndex++;
				if (dataCell != null)
				{
					X.Cell cell = row.Elements<X.Cell>().FirstOrDefault(c => c.CellReference.Value == currentCellId);
					if (cell != null && string.IsNullOrEmpty(dataCell.cellValue))
					{
						cell.Remove();
						X.Hyperlinks hyperlinks = GetWorkSheetHyperlinks();
						if (hyperlinks != null)
						{
							X.Hyperlink hyperlink = hyperlinks.Elements<X.Hyperlink>().FirstOrDefault(h => h.Reference == cellId);
							if (hyperlink != null)
							{
								hyperlink.Remove();
								// TODO : Remove the relationship
							}
						}
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
						X.CellValues dataType = GetCellValueType(dataCell.dataType);
						cell.StyleIndex = dataCell.styleId ?? excel.GetStyleService().GetCellStyleId(dataCell.styleSetting ?? new CellStyleSetting());
						if (dataType == X.CellValues.String)
						{
							cell.DataType = X.CellValues.SharedString;
							cell.CellValue = new X.CellValue(excel.GetShareStringService().InsertUnique(dataCell.cellValue));
						}
						else
						{
							cell.DataType = dataType;
							cell.CellValue = new X.CellValue(dataCell.cellValue);
						}
						if (dataCell.hyperlinkProperties != null)
						{
							string relationshipId = GetNextSheetPartRelationId();
							switch (dataCell.hyperlinkProperties.hyperlinkPropertyType)
							{
								case HyperlinkPropertyTypeValues.EXISTING_FILE:
									dataCell.hyperlinkProperties.relationId = relationshipId;
									dataCell.hyperlinkProperties.action = "ppaction://hlinkfile";
									GetWorksheetPart().AddHyperlinkRelationship(new Uri(dataCell.hyperlinkProperties.value), true, relationshipId);
									break;
								case HyperlinkPropertyTypeValues.TARGET_SHEET: // Target use location Do nothing in relation
									break;
								case HyperlinkPropertyTypeValues.TARGET_SLIDE:
								case HyperlinkPropertyTypeValues.FIRST_SLIDE:
								case HyperlinkPropertyTypeValues.LAST_SLIDE:
								case HyperlinkPropertyTypeValues.NEXT_SLIDE:
								case HyperlinkPropertyTypeValues.PREVIOUS_SLIDE:
									throw new ArgumentException("This Option is valid only for Powerpoint Files");
								default:// Web URL
									dataCell.hyperlinkProperties.relationId = relationshipId;
									GetWorksheetPart().AddHyperlinkRelationship(new Uri(dataCell.hyperlinkProperties.value), true, relationshipId);
									break;
							}
							if (dataCell.hyperlinkProperties.hyperlinkPropertyType == HyperlinkPropertyTypeValues.TARGET_SHEET)
							{
								AddHyperlink(relationshipId, dataCell.hyperlinkProperties.toolTip, dataCell.hyperlinkProperties.value);
							}
							else
							{
								AddHyperlink(relationshipId, dataCell.hyperlinkProperties.toolTip);
							}
						}
					}
				}
			}
			openXMLworksheet.Save();
		}
		/// <summary>
		/// Gets the CellValues enumeration corresponding to the specified cell data type.
		/// </summary>
		private X.CellValues GetCellValueType(CellDataType cellDataType)
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
		private DataType GetCellDataType(EnumValue<X.CellValues> cellValueType)
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
		public Picture AddPicture(string filePath, ExcelPictureSetting pictureSetting)
		{
			return AddPicture(new FileStream(filePath, FileMode.Open, FileAccess.Read), pictureSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Picture AddPicture(Stream stream, ExcelPictureSetting pictureSetting)
		{
			if (pictureSetting.from.column < pictureSetting.to.column || pictureSetting.from.row < pictureSetting.to.row)
			{
				return new Picture(this, stream, pictureSetting);
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
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, areaChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, barChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, columnChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, lineChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, pieChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart<ApplicationSpecificSetting> AddChart<ApplicationSpecificSetting>(DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting) where ApplicationSpecificSetting : ExcelSetting, new()
		{
			ChartData[][] chartData = PrepareCacheData(dataRange);
			dataRange.sheetName = dataRange.sheetName ?? GetSheetName();
			return new Chart<ApplicationSpecificSetting>(this, chartData, dataRange, comboChartSetting);
		}
		/// <summary>
		/// Add Merge cell range to the current sheet if it's not overlapping with existing range
		/// </summary>
		/// <returns>Return true is inserted and false if failed</returns>
		public bool SetMergeCell(MergeCellRange mergeCellRange)
		{
			List<MergeCellRange> existingMergeCellRange = GetMergeCellList();
			Tuple<int, int> newTopLeft = ConverterUtils.ConvertFromExcelCellReference(mergeCellRange.topLeftCell);
			Tuple<int, int> newBottomRight = ConverterUtils.ConvertFromExcelCellReference(mergeCellRange.bottomRightCell);
			if (existingMergeCellRange.Any(range =>
			{
				Tuple<int, int> topLeft = ConverterUtils.ConvertFromExcelCellReference(range.topLeftCell);
				Tuple<int, int> bottomRight = ConverterUtils.ConvertFromExcelCellReference(range.bottomRightCell);
				return IsWithinRange(newTopLeft.Item1, newTopLeft.Item2, topLeft.Item1, topLeft.Item2, bottomRight.Item1, bottomRight.Item2) ||
				IsWithinRange(newBottomRight.Item1, newBottomRight.Item2, topLeft.Item1, topLeft.Item2, bottomRight.Item1, bottomRight.Item2);
			}))
			{
				return false;
			}
			else
			{
				X.MergeCells mergeCells = GetWorkSheetMergeCell();
				if (mergeCells == null)
				{
					mergeCells = new X.MergeCells();
					openXMLworksheet.Append(mergeCells);
				}
				if (mergeCells != null && mergeCells.Count > 0)
				{
					mergeCells.Append(new X.MergeCell()
					{
						Reference = string.Format("{0}:{1}", mergeCellRange.topLeftCell, mergeCellRange.bottomRightCell)
					});
				}
				return true;
			}
		}
		/// <summary>
		/// Return every merged cell range in current sheet
		/// </summary>
		public List<MergeCellRange> GetMergeCellList()
		{
			List<MergeCellRange> mergeCellRanges = new List<MergeCellRange>();
			X.MergeCells mergeCells = GetWorkSheetMergeCell();
			if (mergeCells != null && mergeCells.Count > 0)
			{
				mergeCells.Descendants<X.MergeCell>().ToList().ForEach(mergeCell =>
				{
					mergeCellRanges.Add(new MergeCellRange()
					{
						topLeftCell = mergeCell.Reference.ToString().Split(':')[0],
						bottomRightCell = mergeCell.Reference.ToString().Split(':')[1]
					});
				});
			}
			return mergeCellRanges;
		}
		private static bool IsWithinRange(int x, int y, int topLeftX, int topLeftY, int bottomRightX, int bottomRightY)
		{
			return x >= topLeftX && x <= bottomRightX && y >= topLeftY && y <= bottomRightY;
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
			ChartData[][] chartData = new ChartData[rowEnd - rowStart + 1][];
			dataRows.ForEach(row =>
			{
				chartData[(int)(row.RowIndex.Value - rowStart)] = new ChartData[colEnd - colStart + 1];
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
					// TODO : Cell Value is bit confusing for value types and formula do further research for extending the functionality
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
					Console.WriteLine(cellValue ?? "");
					chartData[(int)(row.RowIndex.Value - rowStart)][colIndex - colStart] = new ChartData()
					{
						dataType = cellDataType,
						// TODO : Do Performance Update
						numberFormat = excel.GetStyleService().GetStyleForId(cell.StyleIndex).numberFormat,
						value = cellValue ?? ""
					};
				});
			});
			return chartData;
		}
	}
}
