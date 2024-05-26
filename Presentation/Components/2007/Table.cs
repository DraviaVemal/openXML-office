// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using G = OpenXMLOffice.Global_2007;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Represents Table Class
	/// </summary>
	public class Table : G.CommonProperties
	{
		private class MergeRange
		{
			public int topLeftX;
			public int topLeftY;
			public int bottomRightX;
			public int bottomRightY;
		}
		private List<MergeRange> mergeRanges = new List<MergeRange>();
		private readonly P.GraphicFrame graphicFrame = new P.GraphicFrame();
		private readonly TableSetting tableSetting;
		/// <summary>
		/// Create Table with provided settings
		/// </summary>
		public Table(TableRow[] TableRows, TableSetting TableSetting)
		{
			tableSetting = TableSetting;
			CreateTableGraphicFrame(TableRows);
		}
		/// <summary>
		/// </summary>
		/// <returns>
		/// X,Y
		/// </returns>
		public Tuple<uint, uint> GetPosition()
		{
			return Tuple.Create(tableSetting.x, tableSetting.y);
		}
		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		public Tuple<uint, uint> GetSize()
		{
			return Tuple.Create(tableSetting.width, tableSetting.height);
		}
		/// <summary>
		/// Get Table Graphic Frame
		/// </summary>
		/// <returns>
		/// </returns>
		public P.GraphicFrame GetTableGraphicFrame()
		{
			return graphicFrame;
		}
		/// <summary>
		/// Update Table Position
		/// </summary>
		public void UpdatePosition(uint X, uint Y)
		{
			tableSetting.x = (uint)G.ConverterUtils.PixelsToEmu((int)X);
			tableSetting.y = (uint)G.ConverterUtils.PixelsToEmu((int)Y);
			if (graphicFrame != null)
			{
				graphicFrame.Transform = new P.Transform
				{
					Offset = new A.Offset { X = tableSetting.x, Y = tableSetting.y },
					Extents = new A.Extents { Cx = tableSetting.width, Cy = tableSetting.height }
				};
			}
		}
		/// <summary>
		/// Update Table Size
		/// </summary>
		public void UpdateSize(uint Width, uint Height)
		{
			ReCalculateColumnWidth();
			tableSetting.width = (uint)G.ConverterUtils.PixelsToEmu((int)Width);
			tableSetting.height = (uint)G.ConverterUtils.PixelsToEmu((int)Height);
			if (graphicFrame != null)
			{
				graphicFrame.Transform = new P.Transform
				{
					Offset = new A.Offset { X = tableSetting.x, Y = tableSetting.y },
					Extents = new A.Extents { Cx = tableSetting.width, Cy = tableSetting.height }
				};
			}
		}
		private long CalculateColumnWidth(TableSetting.WidthOptionValues widthType, float InputWidth)
		{
			int calculatedWidth;
			switch (widthType)
			{
				case TableSetting.WidthOptionValues.PIXEL:
					calculatedWidth = (int)G.ConverterUtils.PixelsToEmu(Convert.ToInt32(InputWidth));
					break;
				case TableSetting.WidthOptionValues.PERCENTAGE:
					calculatedWidth = Convert.ToInt32(tableSetting.width / 100 * InputWidth);
					break;
				case TableSetting.WidthOptionValues.RATIO:
					calculatedWidth = Convert.ToInt32(tableSetting.width / 100 * (InputWidth * 10));
					break;
				default:
					calculatedWidth = Convert.ToInt32(InputWidth);
					break;
			}
			return calculatedWidth;
		}
		private void CreateTableGraphicFrame(TableRow[] TableRows)
		{
			A.GraphicData GraphicData = new A.GraphicData(CreateTable(TableRows))
			{
				Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
			};
			graphicFrame.NonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties(
			   new P.NonVisualDrawingProperties()
			   {
				   Id = 1,
				   Name = "Table 1"
			   },
			   new P.NonVisualGraphicFrameDrawingProperties(
				new A.GraphicFrameLocks()
				{
					NoGrouping = true
				}),
			   new P.ApplicationNonVisualDrawingProperties());
			graphicFrame.Graphic = new A.Graphic()
			{
				GraphicData = GraphicData
			};
			graphicFrame.Transform = new P.Transform()
			{
				Offset = new A.Offset()
				{
					X = tableSetting.x,
					Y = tableSetting.y
				},
				Extents = new A.Extents()
				{
					Cx = tableSetting.width,
					Cy = tableSetting.height
				}
			};
		}
		private A.TableGrid CreateTableGrid(int ColumnCount)
		{
			A.TableGrid TableGrid = new A.TableGrid();
			if (tableSetting.widthType == TableSetting.WidthOptionValues.AUTO)
			{
				for (int i = 0; i < ColumnCount; i++)
				{
					TableGrid.Append(new A.GridColumn() { Width = tableSetting.width / ColumnCount });
				}
			}
			else
			{
				for (int i = 0; i < ColumnCount; i++)
				{
					TableGrid.Append(new A.GridColumn() { Width = CalculateColumnWidth(tableSetting.widthType, tableSetting.tableColumnWidth[i]) });
				}
			}
			return TableGrid;
		}
		private void AddMergeRange(int topLeftX, int topLeftY, int bottomRightX, int bottomRightY)
		{
			if (mergeRanges.Any(
				range =>
					G.Validation.IsWithinRange(topLeftX, topLeftY, range.topLeftX, range.topLeftY, range.bottomRightX, range.bottomRightY) ||
					G.Validation.IsWithinRange(bottomRightX, bottomRightY, range.topLeftX, range.topLeftY, range.bottomRightX, range.bottomRightY)))
			{
				throw new ArgumentException("Table Merge Range Conflict");
			}
			mergeRanges.Add(new MergeRange()
			{
				topLeftX = topLeftX,
				topLeftY = topLeftY,
				bottomRightX = bottomRightX,
				bottomRightY = bottomRightY,
			});
		}
		private bool CheckIsColumnMerged(int col, int row)
		{
			return mergeRanges.Any(range =>
			{
				return range.topLeftY == range.bottomRightY && G.Validation.IsWithinRange(col, row, range.topLeftX, range.topLeftY, range.bottomRightX, range.bottomRightY);
			});
		}
		private bool CheckIsRowMerged(int col, int row)
		{
			return mergeRanges.Any(range =>
			{
				return range.topLeftX == range.bottomRightX && G.Validation.IsWithinRange(col, row, range.topLeftX, range.topLeftY, range.bottomRightX, range.bottomRightY);
			});
		}
		private bool CheckIsRangeMerged(int col, int row)
		{
			return mergeRanges.Any(range =>
			{
				return G.Validation.IsWithinRange(col, row, range.topLeftX, range.topLeftY, range.bottomRightX, range.bottomRightY);
			});
		}
		private A.Table CreateTable(TableRow[] TableRows)
		{
			int columnSize = TableRows.Max(item => item.tableCells.Count);
			if (columnSize < 1)
			{
				throw new DataException("No Table Data Provided");
			}
			A.Table Table = new A.Table()
			{
				TableProperties = new A.TableProperties()
				{
					FirstRow = true,
					BandRow = true
				},
				TableGrid = CreateTableGrid(columnSize)
			};
			int rowIndex = 0;
			// Add Table Data Row
			foreach (TableRow row in TableRows)
			{
				Table.Append(CreateTableRow(row, rowIndex, columnSize));
				++rowIndex;
			}
			return Table;
		}
		private A.TableRow CreateTableRow(TableRow row, int rowIndex, int columnSize)
		{
			A.TableRow TableRow = new A.TableRow()
			{
				Height = row.height
			};
			int columnIndex = 0;
			foreach (TableCell cell in row.tableCells)
			{
				TableRow.Append(CreateTableCell(cell, row, rowIndex, columnIndex));
				++columnIndex;
			}
			for (int i = columnIndex; i < columnSize; i++)
			{
				TableRow.Append(CreateTableCell(new TableCell(), row, rowIndex, columnIndex));
			}
			return TableRow;
		}

		private A.TableCell CreateTableCell(TableCell cell, TableRow row, int rowIndex, int columnIndex)
		{
			if (cell.rowSpan > 1 || cell.columnSpan > 1)
			{
				AddMergeRange(rowIndex, columnIndex, (int)(cell.rowSpan > 1 ? (rowIndex + cell.rowSpan) : rowIndex), (int)(cell.columnSpan > 1 ? (columnIndex + cell.columnSpan) : columnIndex));
			}
			A.Paragraph paragraph = new A.Paragraph();
			if (cell.horizontalAlignment != null)
			{
				A.TextAlignmentTypeValues alignment;
				switch (cell.horizontalAlignment)
				{
					case G.HorizontalAlignmentValues.CENTER:
						alignment = A.TextAlignmentTypeValues.Center;
						break;
					case G.HorizontalAlignmentValues.LEFT:
						alignment = A.TextAlignmentTypeValues.Left;
						break;
					case G.HorizontalAlignmentValues.JUSTIFY:
						alignment = A.TextAlignmentTypeValues.Justified;
						break;
					case G.HorizontalAlignmentValues.RIGHT:
						alignment = A.TextAlignmentTypeValues.Right;
						break;
					default:
						alignment = A.TextAlignmentTypeValues.Left;
						break;
				}
				paragraph.Append(new A.ParagraphProperties()
				{
					Alignment = alignment
				});
			}
			if (cell.textValue == null)
			{
				paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
			}
			else
			{
				G.SolidFillModel solidFillModel = new G.SolidFillModel()
				{
					schemeColorModel = new G.SchemeColorModel()
					{
						themeColorValues = G.ThemeColorValues.TEXT_1
					}
				};
				if (cell.textColor != null)
				{
					solidFillModel.hexColor = cell.textColor;
					solidFillModel.schemeColorModel = null;
				}
				paragraph.Append(CreateDrawingRun(new List<G.DrawingRunModel>()
				{
					new G.DrawingRunModel(){
						text = cell.textValue,
						textHighlight = cell.textBackground,
						drawingRunProperties = new G.DrawingRunPropertiesModel()
						{
							solidFill = solidFillModel,
							fontFamily = cell.fontFamily,
							fontSize = cell.fontSize,
							isBold = cell.isBold,
							isItalic = cell.isItalic,
							underLineValues = cell.underLineValues,
						}
					}
				}.ToArray()));
			}
			A.TableCell tableCellXML = new A.TableCell(new A.TextBody(
				new A.BodyProperties(),
				new A.ListStyle(),
				paragraph
			));
			if (cell.columnSpan > 1)
			{
				tableCellXML.GridSpan = (int)cell.columnSpan;
			}
			if (cell.rowSpan > 1)
			{
				tableCellXML.RowSpan = (int)cell.rowSpan;
			}
			tableCellXML.HorizontalMerge = CheckIsColumnMerged(columnIndex, rowIndex);
			tableCellXML.VerticalMerge = CheckIsRowMerged(columnIndex, rowIndex);
			if (!(tableCellXML.HorizontalMerge || tableCellXML.VerticalMerge) && CheckIsRangeMerged(columnIndex, rowIndex))
			{
				tableCellXML.HorizontalMerge = true;
				tableCellXML.VerticalMerge = true;
			}
			A.TextAnchoringTypeValues anchor;
			switch (cell.verticalAlignment)
			{
				case G.VerticalAlignmentValues.TOP:
					anchor = A.TextAnchoringTypeValues.Top;
					break;
				case G.VerticalAlignmentValues.MIDDLE:
					anchor = A.TextAnchoringTypeValues.Center;
					break;
				case G.VerticalAlignmentValues.BOTTOM:
					anchor = A.TextAnchoringTypeValues.Bottom;
					break;
				default:
					anchor = A.TextAnchoringTypeValues.Top;
					break;
			}
			A.TableCellProperties tableCellProperties = new A.TableCellProperties()
			{
				Anchor = anchor
			};
			if (cell.borderSettings.leftBorder.showBorder)
			{
				tableCellProperties.Append(new A.LeftBorderLineProperties(
					CreateSolidFill(new G.SolidFillModel() { hexColor = cell.borderSettings.leftBorder.borderColor }),
					new A.PresetDash() { Val = GetDashStyleValue(cell.borderSettings.leftBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.leftBorder.width),
					CompoundLineType = GetBorderStyleValue(cell.borderSettings.leftBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.LeftBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.rightBorder.showBorder)
			{
				tableCellProperties.Append(new A.RightBorderLineProperties(
					  CreateSolidFill(new G.SolidFillModel() { hexColor = cell.borderSettings.rightBorder.borderColor }),
					new A.PresetDash() { Val = GetDashStyleValue(cell.borderSettings.rightBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.rightBorder.width),
					CompoundLineType = GetBorderStyleValue(cell.borderSettings.rightBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.RightBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.topBorder.showBorder)
			{
				tableCellProperties.Append(new A.TopBorderLineProperties(
					 CreateSolidFill(new G.SolidFillModel() { hexColor = cell.borderSettings.topBorder.borderColor }),
					new A.PresetDash() { Val = GetDashStyleValue(cell.borderSettings.topBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.topBorder.width),
					CompoundLineType = GetBorderStyleValue(cell.borderSettings.topBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.TopBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.bottomBorder.showBorder)
			{
				tableCellProperties.Append(new A.BottomBorderLineProperties(
					CreateSolidFill(new G.SolidFillModel() { hexColor = cell.borderSettings.bottomBorder.borderColor }),
					new A.PresetDash() { Val = GetDashStyleValue(cell.borderSettings.bottomBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.bottomBorder.width),
					CompoundLineType = GetBorderStyleValue(cell.borderSettings.bottomBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.BottomBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.topLeftToBottomRightBorder.showBorder)
			{
				tableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(
					CreateSolidFill(new G.SolidFillModel() { hexColor = cell.borderSettings.topLeftToBottomRightBorder.borderColor }),
					new A.PresetDash() { Val = GetDashStyleValue(cell.borderSettings.topLeftToBottomRightBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.topLeftToBottomRightBorder.width),
					CompoundLineType = GetBorderStyleValue(cell.borderSettings.topLeftToBottomRightBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.bottomLeftToTopRightBorder.showBorder)
			{
				tableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(
					CreateSolidFill(new G.SolidFillModel() { hexColor = cell.borderSettings.bottomLeftToTopRightBorder.borderColor }),
					new A.PresetDash() { Val = GetDashStyleValue(cell.borderSettings.bottomLeftToTopRightBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.bottomLeftToTopRightBorder.width),
					CompoundLineType = GetBorderStyleValue(cell.borderSettings.bottomLeftToTopRightBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(new A.NoFill()));
			}
			if (cell.cellBackground != null || row.rowBackground != null)
			{
				tableCellProperties.Append(CreateSolidFill(new G.SolidFillModel() { hexColor = cell.cellBackground ?? row.rowBackground }));
			}
			else
			{
				tableCellProperties.Append(new A.NoFill());
			}
			tableCellXML.Append(tableCellProperties);
			return tableCellXML;
		}
		private void ReCalculateColumnWidth()
		{
			A.Table Table = graphicFrame.Graphic.GraphicData.GetFirstChild<A.Table>();
			if (Table != null)
			{
				List<A.GridColumn> GridColumn = Table.TableGrid.Elements<A.GridColumn>().ToList();
				if (tableSetting.widthType == TableSetting.WidthOptionValues.AUTO)
				{
					GridColumn.ForEach(Column => Column.Width = tableSetting.width / GridColumn.Count);
				}
				else
				{
					GridColumn.Select((item, index) => Tuple.Create(item, index)).ToList().ForEach(result =>
						result.Item1.Width = CalculateColumnWidth(tableSetting.widthType, tableSetting.tableColumnWidth[result.Item2]));
				}
			}
		}
	}
}
