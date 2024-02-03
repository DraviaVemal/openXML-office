// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Data;
using A = DocumentFormat.OpenXml.Drawing;
using G = OpenXMLOffice.Global_2013;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	/// <summary>
	/// Represents Table Class
	/// </summary>
	public class Table : G.CommonProperties
	{
		private readonly P.GraphicFrame graphicFrame = new();
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
		public (uint, uint) GetPosition()
		{
			return (tableSetting.x, tableSetting.y);
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		public (uint, uint) GetSize()
		{
			return (tableSetting.width, tableSetting.height);
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
			tableSetting.x = X;
			tableSetting.y = Y;
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
			tableSetting.width = Width;
			tableSetting.height = Height;
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
			return widthType switch
			{
				TableSetting.WidthOptionValues.PIXEL => G.ConverterUtils.PixelsToEmu(Convert.ToInt32(InputWidth)),
				TableSetting.WidthOptionValues.PERCENTAGE => Convert.ToInt32(tableSetting.width / 100 * InputWidth),
				TableSetting.WidthOptionValues.RATIO => Convert.ToInt32(tableSetting.width / 100 * (InputWidth * 10)),
				_ => Convert.ToInt32(InputWidth)
			};
		}

		private A.Table CreateTable(TableRow[] TableRows)
		{
			if (TableRows.Length < 1 || TableRows[0].tableCells.Count < 1)
			{
				throw new DataException("No Table Data Provided");
			}
			if (tableSetting.widthType != TableSetting.WidthOptionValues.AUTO && tableSetting.tableColumnWidth.Count != TableRows[0].tableCells.Count)
			{
				throw new ArgumentException("Column With Setting Does Not Match Data");
			}
			A.Table Table = new()
			{
				TableProperties = new A.TableProperties()
				{
					FirstRow = true,
					BandRow = true
				},
				TableGrid = CreateTableGrid(TableRows[0].tableCells.Count)
			};
			// Add Table Data Row
			foreach (TableRow row in TableRows)
			{
				Table.Append(CreateTableRow(row));
			}
			return Table;
		}

		private static A.TableCell CreateTableCell(TableCell cell, TableRow row)
		{
			A.Paragraph paragraph = new();
			if (cell.horizontalAlignment != null)
			{
				paragraph.Append(new A.ParagraphProperties()
				{
					Alignment = cell.horizontalAlignment switch
					{
						G.HorizontalAlignmentValues.CENTER => A.TextAlignmentTypeValues.Center,
						G.HorizontalAlignmentValues.LEFT => A.TextAlignmentTypeValues.Left,
						G.HorizontalAlignmentValues.JUSTIFY => A.TextAlignmentTypeValues.Justified,
						G.HorizontalAlignmentValues.RIGHT => A.TextAlignmentTypeValues.Right,
						_ => A.TextAlignmentTypeValues.Left
					},
				});
			}
			if (cell.value == null)
			{
				paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
			}
			else
			{
				G.SolidFillModel solidFillModel = new()
				{
					schemeColorModel = new()
					{
						themeColorValues = G.ThemeColorValues.TEXT_1
					}
				};
				if (cell.textColor != null)
				{
					solidFillModel.hexColor = cell.textColor;
					solidFillModel.schemeColorModel = null;
				}
				paragraph.Append(CreateDrawingRun(new()
				{
					text = cell.value,
					textBackground = cell.textBackground,
					drawingRunProperties = new()
					{
						solidFill = solidFillModel,
						fontFamily = cell.fontFamily,
						fontSize = cell.fontSize,
						isBold = cell.isBold,
						isItalic = cell.isItalic,
						underline = cell.isUnderline ? G.UnderLineValues.SINGLE : G.UnderLineValues.NONE,
					}
				}));
			}
			A.TableCell tableCellXML = new(new A.TextBody(
				new A.BodyProperties(),
				new A.ListStyle(),
				paragraph
			));
			A.TableCellProperties tableCellProperties = new()
			{
				Anchor = cell.verticalAlignment switch
				{
					G.VerticalAlignmentValues.TOP => A.TextAnchoringTypeValues.Top,
					G.VerticalAlignmentValues.MIDDLE => A.TextAnchoringTypeValues.Center,
					G.VerticalAlignmentValues.BOTTOM => A.TextAnchoringTypeValues.Bottom,
					_ => A.TextAnchoringTypeValues.Top
				}
			};
			if (cell.borderSettings.leftBorder.showBorder)
			{
				tableCellProperties.Append(new A.LeftBorderLineProperties(
					CreateSolidFill(new() { hexColor = cell.borderSettings.leftBorder.borderColor }),
					new A.PresetDash() { Val = TableBorderSetting.GetDashStyleValue(cell.borderSettings.leftBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.leftBorder.width),
					CompoundLineType = TableBorderSetting.GetBorderStyleValue(cell.borderSettings.leftBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.LeftBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.rightBorder.showBorder)
			{
				tableCellProperties.Append(new A.RightBorderLineProperties(
					  CreateSolidFill(new() { hexColor = cell.borderSettings.rightBorder.borderColor }),
					new A.PresetDash() { Val = TableBorderSetting.GetDashStyleValue(cell.borderSettings.rightBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.rightBorder.width),
					CompoundLineType = TableBorderSetting.GetBorderStyleValue(cell.borderSettings.rightBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.RightBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.topBorder.showBorder)
			{
				tableCellProperties.Append(new A.TopBorderLineProperties(
					 CreateSolidFill(new() { hexColor = cell.borderSettings.topBorder.borderColor }),
					new A.PresetDash() { Val = TableBorderSetting.GetDashStyleValue(cell.borderSettings.topBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.topBorder.width),
					CompoundLineType = TableBorderSetting.GetBorderStyleValue(cell.borderSettings.topBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.TopBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.bottomBorder.showBorder)
			{
				tableCellProperties.Append(new A.BottomBorderLineProperties(
					CreateSolidFill(new() { hexColor = cell.borderSettings.bottomBorder.borderColor }),
					new A.PresetDash() { Val = TableBorderSetting.GetDashStyleValue(cell.borderSettings.bottomBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.bottomBorder.width),
					CompoundLineType = TableBorderSetting.GetBorderStyleValue(cell.borderSettings.bottomBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.BottomBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.topLeftToBottomRightBorder.showBorder)
			{
				tableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(
					CreateSolidFill(new() { hexColor = cell.borderSettings.topLeftToBottomRightBorder.borderColor }),
					new A.PresetDash() { Val = TableBorderSetting.GetDashStyleValue(cell.borderSettings.topLeftToBottomRightBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.topLeftToBottomRightBorder.width),
					CompoundLineType = TableBorderSetting.GetBorderStyleValue(cell.borderSettings.topLeftToBottomRightBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(new A.NoFill()));
			}
			if (cell.borderSettings.bottomLeftToTopRightBorder.showBorder)
			{
				tableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(
					CreateSolidFill(new() { hexColor = cell.borderSettings.bottomLeftToTopRightBorder.borderColor }),
					new A.PresetDash() { Val = TableBorderSetting.GetDashStyleValue(cell.borderSettings.bottomLeftToTopRightBorder.dashStyle) }
				)
				{
					Width = (DocumentFormat.OpenXml.Int32Value)G.ConverterUtils.PixelsToEmu((int)cell.borderSettings.bottomLeftToTopRightBorder.width),
					CompoundLineType = TableBorderSetting.GetBorderStyleValue(cell.borderSettings.bottomLeftToTopRightBorder.borderStyle)
				});
			}
			else
			{
				tableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(new A.NoFill()));
			}
			tableCellProperties.Append((cell.cellBackground != null || row.rowBackground != null) ? G.CommonProperties.CreateSolidFill(new() { hexColor = cell.cellBackground ?? row.rowBackground! }) : new A.NoFill());
			tableCellXML.Append(tableCellProperties);
			return tableCellXML;
		}

		private void CreateTableGraphicFrame(TableRow[] TableRows)
		{
			A.GraphicData GraphicData = new(CreateTable(TableRows))
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
			A.TableGrid TableGrid = new();
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

		private A.TableRow CreateTableRow(TableRow Row)
		{
			A.TableRow TableRow = new()
			{
				Height = Row.height
			};
			foreach (TableCell cell in Row.tableCells)
			{
				TableRow.Append(CreateTableCell(cell, Row));
			}
			return TableRow;
		}

		private void ReCalculateColumnWidth()
		{
			A.Table? Table = graphicFrame!.Graphic!.GraphicData!.GetFirstChild<A.Table>();
			if (Table != null)
			{
				List<A.GridColumn> GridColumn = Table.TableGrid!.Elements<A.GridColumn>().ToList();
				if (tableSetting.widthType == TableSetting.WidthOptionValues.AUTO)
				{
					GridColumn.ForEach(Column => Column.Width = tableSetting.width / GridColumn.Count);
				}
				else
				{
					GridColumn.Select((item, index) => (item, index)).ToList().ForEach(Column =>
						Column.item.Width = CalculateColumnWidth(tableSetting.widthType, tableSetting.tableColumnWidth[Column.index]));
				}
			}
		}


	}
}
