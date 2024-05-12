// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Chart Base Class Common to all charts. Class is only intended to get created by inherited classes
	/// </summary>
	public class ChartBase<ApplicationSpecificSetting> : CommonProperties where ApplicationSpecificSetting : class, ISizeAndPosition, new()
	{
		internal const int AccentColurCount = 6;
		internal uint CategoryAxisId = 1362418656;
		internal uint ValueAxisId = 1358349936;
		internal const int SecondaryCategoryAxisId = 1615085760;
		internal const int SecondaryValueAxisId = 1474633616;
		/// <summary>
		/// Chart Data Groupings
		/// </summary>
		internal List<ChartDataGrouping> chartDataGroupings = new List<ChartDataGrouping>();
		/// <summary>
		/// Core chart settings common for every possible chart
		/// </summary>
		internal ChartSetting<ApplicationSpecificSetting> chartSetting;
		private readonly C.Chart chart;
		private readonly C.ChartSpace openXMLChartSpace;
		/// <summary>
		/// Chartbase class constructor restricted only for inheritance use
		/// </summary>
		/// <param name="chartSetting">
		/// </param>
		internal ChartBase(ChartSetting<ApplicationSpecificSetting> chartSetting)
		{
			CategoryAxisId = chartSetting.categoryAxisId ?? CategoryAxisId;
			ValueAxisId = chartSetting.valueAxisId ?? ValueAxisId;
			this.chartSetting = chartSetting;
			openXMLChartSpace = CreateChartSpace();
			chart = CreateChart();
			GetChartSpace().Append(chart);
		}
		/// <summary>
		/// Get OpenXML ChartSpace
		/// </summary>
		public virtual C.ChartSpace GetChartSpace()
		{
			return openXMLChartSpace;
		}
		/// <summary>
		/// Create Bubble Size Axis for the chart
		/// </summary>
		internal static C.BubbleSize CreateBubbleSizeAxisData(string formula, ChartData[] cells)
		{
			if (cells.All(v => v.dataType != DataType.NUMBER))
			{
				Console.WriteLine(string.Format("Object Details Value : {0} is not numeric", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).value));
				Console.WriteLine(string.Format("Object Details Number Format : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).numberFormat));
				Console.WriteLine(string.Format("Object Details Data Type : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).dataType));
				LogUtils.ShowWarning("Bubble Size Data Prefered in numaric.");
			}
			return new C.BubbleSize(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
		}
		/// <summary>
		/// Create Category Axis for the chart
		/// </summary>
		/// <param name="categoryAxisSetting">
		/// </param>
		/// <returns>
		/// </returns>
		internal C.CategoryAxis CreateCategoryAxis(CategoryAxisSetting categoryAxisSetting)
		{
			C.AxisPositionValues axisPositionValue;
			switch (categoryAxisSetting.axisPosition)
			{
				case AxisPosition.LEFT:
					axisPositionValue = C.AxisPositionValues.Left;
					break;
				case AxisPosition.RIGHT:
					axisPositionValue = C.AxisPositionValues.Right;
					break;
				case AxisPosition.TOP:
					axisPositionValue = C.AxisPositionValues.Top;
					break;
				default:
					axisPositionValue = C.AxisPositionValues.Bottom;
					break;
			}
			C.CategoryAxis CategoryAxis = new C.CategoryAxis(
				new C.AxisId { Val = categoryAxisSetting.id },
				new C.Scaling(new C.Orientation { Val = categoryAxisSetting.invertOrder ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax }),
				new C.Delete { Val = !categoryAxisSetting.isVisible },
				new C.AxisPosition { Val = axisPositionValue },
				new C.MajorTickMark { Val = C.TickMarkValues.None },
				new C.MinorTickMark { Val = C.TickMarkValues.None },
				new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo }
			);
			if (categoryAxisSetting.isVisible)
			{
				if (chartSetting.chartGridLinesOptions.isMajorCategoryLinesEnabled)
				{
					CategoryAxis.Append(CreateMajorGridLine());
				}
				if (chartSetting.chartGridLinesOptions.isMinorCategoryLinesEnabled)
				{
					CategoryAxis.Append(CreateMinorGridLine());
				}
				CategoryAxis.Append(CreateChartShapeProperties());
				SolidFillModel solidFillModel = new SolidFillModel()
				{
					schemeColorModel = new SchemeColorModel()
					{
						themeColorValues = ThemeColorValues.TEXT_1,
						luminanceModulation = 65000,
						luminanceOffset = 35000
					}
				};
				if (categoryAxisSetting.fontColor != null)
				{
					solidFillModel.hexColor = categoryAxisSetting.fontColor;
					solidFillModel.schemeColorModel = null;
				}
				CategoryAxis.Append(CreateChartTextProperties(new ChartTextPropertiesModel()
				{
					drawingBodyProperties = new DrawingBodyPropertiesModel(),
					drawingParagraph = new DrawingParagraphModel()
					{
						paragraphPropertiesModel = new ParagraphPropertiesModel()
						{
							defaultRunProperties = new DefaultRunPropertiesModel()
							{
								solidFill = solidFillModel,
								fontSize = ConverterUtils.FontSizeToFontSize(categoryAxisSetting.fontSize),
								isBold = categoryAxisSetting.isBold,
								isItalic = categoryAxisSetting.isItalic,
								underline = categoryAxisSetting.underLineValues,
								strike = categoryAxisSetting.strikeValues,
								baseline = 0,
							}
						}
					}
				}));
			}
			CategoryAxis.Append(
				new C.CrossingAxis { Val = categoryAxisSetting.crossAxisId },
				new C.Crosses { Val = C.CrossesValues.AutoZero },
				new C.AutoLabeled { Val = true },
				new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
				new C.LabelOffset { Val = 100 },
				new C.NoMultiLevelLabels { Val = false });
			return CategoryAxis;
		}
		/// <summary>
		/// 
		/// </summary>
		internal C.TrendlineLabel CreateTrendLineLabel()
		{
			C.TrendlineLabel trendlineLabel = new C.TrendlineLabel
			{
				NumberingFormat = new C.NumberingFormat() { FormatCode = "General", SourceLinked = false },
			};
			trendlineLabel.Append(CreateChartShapeProperties(new ShapePropertiesModel()));
			trendlineLabel.Append(CreateChartTextProperties(new ChartTextPropertiesModel()
			{
				drawingBodyProperties = new DrawingBodyPropertiesModel()
				{
					rotation = 0,
					useParagraphSpacing = true,
					verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
					vertical = TextVerticalAlignmentValues.HORIZONTAL,
					wrap = TextWrappingValues.SQUARE,
					anchor = TextAnchoringValues.CENTER,
					anchorCenter = true,
				},
				drawingParagraph = new DrawingParagraphModel()
				{
					paragraphPropertiesModel = new ParagraphPropertiesModel()
					{
						defaultRunProperties = new DefaultRunPropertiesModel()
						{
							fontSize = 1197,
							isBold = false,
							isItalic = false,
							underline = UnderLineValues.NONE,
							strike = StrikeValues.NO_STRIKE,
							kerning = 1200,
							baseline = 0,
							solidFill = new SolidFillModel()
							{
								schemeColorModel = new SchemeColorModel()
								{
									themeColorValues = ThemeColorValues.TEXT_1,
									luminanceModulation = 65000,
									luminanceOffset = 35000,
								},
							},
							complexScriptFont = "+mn-cs",
							eastAsianFont = "+mn-ea",
							latinFont = "+mn-lt",
						}
					},
				}
			}));
			return trendlineLabel;
		}
		/// <summary>
		/// Create Chart Shape Properties for the chart
		/// </summary>
		internal static C.Layout CreateLayout(LayoutModel layoutModel = null)
		{
			if (layoutModel == null)
			{
				return new C.Layout();
			}
			double x = layoutModel.x;
			double y = layoutModel.y;
			double width = layoutModel.width;
			double height = layoutModel.height;
			if (x < 0 || x > 1 || width < 0 || width > 1 || x + width < 0 || x + width > 1)
			{
				throw new ArgumentException("Layout value is not within acceptable range. X and Width values should be between 0 and 1, and their sum should be between 0 and 1.");
			}
			if (y < 0 || y > 1 || height < 0 || height > 1 || y + height < 0 || y + height > 1)
			{
				throw new ArgumentException("Layout value is not within acceptable range. Y and Height values should be between 0 and 1, and their sum should be between 0 and 1.");
			}
			return new C.Layout(
				new C.ManualLayout(
					new C.LayoutTarget { Val = C.LayoutTargetValues.Inner },
					new C.LeftMode { Val = C.LayoutModeValues.Edge },
					new C.TopMode { Val = C.LayoutModeValues.Edge },
					new C.Left { Val = x },
					new C.Top { Val = y },
					new C.Width { Val = width },
					new C.Height { Val = height }
				));
		}
		/// <summary>
		/// Create Category Axis Data for the chart
		/// </summary>
		/// <param name="formula">
		/// </param>
		/// <param name="cells">
		/// </param>
		/// <returns>
		/// </returns>
		internal static C.CategoryAxisData CreateCategoryAxisData(string formula, ChartData[] cells)
		{
			if (cells.All(v => v.dataType == DataType.NUMBER))
			{
				return new C.CategoryAxisData(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
			}
			else
			{
				return new C.CategoryAxisData(new C.StringReference(new C.Formula(formula), AddStringCacheValue(cells)));
			}
		}
		/// <summary>
		/// Create Data Series for the chart
		/// </summary>
		internal List<ChartDataGrouping> CreateDataSeries(ChartDataSetting chartDataSetting, ChartData[][] dataCols, DataRange dataRange)
		{
			List<uint> seriesColumns = new List<uint>();
			for (uint col = chartDataSetting.chartDataColumnStart + 1; col <= (chartDataSetting.chartDataColumnEnd == 0 ? dataCols.Length - 1 : (int)chartDataSetting.chartDataColumnEnd); col++)
			{
				seriesColumns.Add(col);
			}
			if ((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - chartDataSetting.chartDataRowStart < 1 || (chartDataSetting.chartDataColumnEnd == 0 ? dataCols.Length : (int)chartDataSetting.chartDataColumnEnd) - chartDataSetting.chartDataColumnStart < 1)
			{
				throw new ArgumentException("Data Series Invalid Range");
			}
			for (int i = 0; i < seriesColumns.Count; i++)
			{
				uint column = seriesColumns[i];
				string sheetName = dataRange != null ? dataRange.sheetName : "Sheet1";
				string columnName = ConverterUtils.ConvertIntToColumnName((int)column + 1);
				string startColumnName = ConverterUtils.ConvertIntToColumnName((int)chartDataSetting.chartDataColumnStart + 1);
				string endColumnName = ConverterUtils.ConvertIntToColumnName((int)chartDataSetting.chartDataColumnStart + 1);
				uint rowNumber = chartDataSetting.chartDataRowStart + 1;
				uint startRowNumber = chartDataSetting.chartDataRowStart + 2;
				List<ChartData> xAxisCells = ((ChartData[])dataCols[chartDataSetting.chartDataColumnStart].Clone()).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
				List<ChartData> yAxisCells = ((ChartData[])dataCols[column].Clone()).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
				long endRowNumberX = chartDataSetting.chartDataRowStart + xAxisCells.Count + 1;
				long endRowNumberY = chartDataSetting.chartDataRowStart + yAxisCells.Count + 1;
				ChartDataGrouping chartDataGrouping = new ChartDataGrouping()
				{
					id = i,
					seriesHeaderFormula = string.Format("'{0}'!${1}${2}", sheetName, columnName, rowNumber),
					seriesHeaderCells = ((ChartData[])dataCols[column].Clone())[chartDataSetting.chartDataRowStart],
					xAxisFormula = string.Format("'{0}'!${1}${2}:${3}${4}", sheetName, startColumnName, startRowNumber, endColumnName, endRowNumberX),
					xAxisCells = xAxisCells.ToArray(),
					yAxisFormula = string.Format("'{0}'!${1}${2}:${3}${4}", sheetName, columnName, startRowNumber, columnName, endRowNumberY),
					yAxisCells = yAxisCells.ToArray(),
				};
				if (chartDataSetting.is3dData && seriesColumns.Count > i + 1)
				{
					i++;
					column = seriesColumns[i];
					List<ChartData> zAxisCells = ((ChartData[])dataCols[column].Clone()).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
					long endRowNumberZ = chartDataSetting.chartDataRowStart + zAxisCells.Count + 1;
					chartDataGrouping.zAxisFormula = string.Format("'{0}'!${1}${2}:${3}${4}", sheetName, columnName, startRowNumber, columnName, endRowNumberZ);
					chartDataGrouping.zAxisCells = zAxisCells.ToArray();
				}
				// TODO: Reorganise to Move to 2013 Namespace extension
				uint DataValueColumn;
				if (chartDataSetting.advancedDataLabel.valueFromColumn.TryGetValue(column, out DataValueColumn))
				{
					columnName = ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1);
					List<ChartData> dataLabelCells = ((ChartData[])dataCols[DataValueColumn].Clone()).Skip((int)chartDataSetting.chartDataRowStart).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
					long endRowNumberD = chartDataSetting.chartDataRowStart + dataLabelCells.Count;
					chartDataGrouping.dataLabelFormula = string.Format("'{0}'!${1}${2}:${3}${4}", sheetName, columnName, startRowNumber, columnName, endRowNumberD);
					chartDataGrouping.dataLabelCells = dataLabelCells.ToArray();
				}
				chartDataGroupings.Add(chartDataGrouping);
			}
			return chartDataGroupings;
		}
		/// <summary>
		/// Create Series Text for the chart
		/// </summary>
		internal static C.SeriesText CreateSeriesText(string formula, ChartData[] cells)
		{
			return new C.SeriesText(new C.StringReference(new C.Formula(formula), AddStringCacheValue(cells)));
		}
		/// <summary>
		/// Create Value Axis for the chart
		/// </summary>
		internal C.ValueAxis CreateValueAxis(ValueAxisSetting valueAxisSetting)
		{
			C.AxisPositionValues axisPositionValue;
			switch (valueAxisSetting.axisPosition)
			{
				case AxisPosition.LEFT:
					axisPositionValue = C.AxisPositionValues.Left;
					break;
				case AxisPosition.RIGHT:
					axisPositionValue = C.AxisPositionValues.Right;
					break;
				case AxisPosition.TOP:
					axisPositionValue = C.AxisPositionValues.Top;
					break;
				default:
					axisPositionValue = C.AxisPositionValues.Bottom;
					break;
			}
			C.ValueAxis valueAxis = new C.ValueAxis(
				new C.AxisId { Val = valueAxisSetting.id },
				new C.Scaling(new C.Orientation { Val = valueAxisSetting.invertOrder ? C.OrientationValues.MaxMin : C.OrientationValues.MinMax }),
				new C.Delete { Val = !valueAxisSetting.isVisible },
				new C.AxisPosition { Val = axisPositionValue });
			if (chartSetting.chartGridLinesOptions.isMajorValueLinesEnabled)
			{
				valueAxis.Append(CreateMajorGridLine());
			}
			if (chartSetting.chartGridLinesOptions.isMinorValueLinesEnabled)
			{
				valueAxis.Append(CreateMinorGridLine());
			}
			valueAxis.Append(
				new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
				new C.MajorTickMark { Val = valueAxisSetting.majorTickMark },
				new C.MinorTickMark { Val = valueAxisSetting.minorTickMark },
				new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
			valueAxis.Append(CreateChartShapeProperties());
			SolidFillModel solidFillModel = new SolidFillModel()
			{
				schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.TEXT_1,
					luminanceModulation = 65000,
					luminanceOffset = 35000
				}
			};
			if (valueAxisSetting.fontColor != null)
			{
				solidFillModel.hexColor = valueAxisSetting.fontColor;
				solidFillModel.schemeColorModel = null;
			}
			valueAxis.Append(CreateChartTextProperties(new ChartTextPropertiesModel()
			{
				drawingBodyProperties = new DrawingBodyPropertiesModel(),
				drawingParagraph = new DrawingParagraphModel()
				{
					paragraphPropertiesModel = new ParagraphPropertiesModel()
					{
						defaultRunProperties = new DefaultRunPropertiesModel()
						{
							solidFill = solidFillModel,
							fontSize = ConverterUtils.FontSizeToFontSize(valueAxisSetting.fontSize),
							isBold = valueAxisSetting.isBold,
							isItalic = valueAxisSetting.isItalic,
							underline = valueAxisSetting.underLineValues,
							strike = valueAxisSetting.strikeValues,
							baseline = 0
						}
					}
				}
			}));
			valueAxis.Append(
				new C.CrossingAxis { Val = valueAxisSetting.crossAxisId },
				new C.Crosses { Val = valueAxisSetting.crosses },
				new C.CrossBetween { Val = C.CrossBetweenValues.Between });
			return valueAxis;
		}
		/// <summary>
		/// Create Value Axis Data for the chart
		/// </summary>
		internal static C.Values CreateValueAxisData(string formula, ChartData[] cells)
		{
			if (cells.All(v => v.dataType != DataType.NUMBER))
			{
				Console.WriteLine(string.Format("Object Details Value : {0} is not numaric", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).value));
				Console.WriteLine(string.Format("Object Details Number Format : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).numberFormat));
				Console.WriteLine(string.Format("Object Details Data Type : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).dataType));
				LogUtils.ShowWarning("Not Every Values is numeric some assumptions are made in chart construction");
			}
			return new C.Values(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
		}
		/// <summary>
		/// Create X Axis Data for the chart
		/// </summary>
		internal static C.XValues CreateXValueAxisData(string formula, ChartData[] cells)
		{
			if (cells.All(v => v.dataType != DataType.NUMBER))
			{
				Console.WriteLine(string.Format("Object Details Value : {0} is not numaric", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).value));
				Console.WriteLine(string.Format("Object Details Number Format : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).numberFormat));
				Console.WriteLine(string.Format("Object Details Data Type : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).dataType));
				LogUtils.ShowWarning("Not Every Values is numeric some assumptions are made in chart construction");
			}
			return new C.XValues(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
		}
		/// <summary>
		/// Create Y Axis Data for the chart
		/// </summary>
		internal static C.YValues CreateYValueAxisData(string formula, ChartData[] cells)
		{
			if (cells.All(v => v.dataType != DataType.NUMBER))
			{
				Console.WriteLine(string.Format("Object Details Value : {0} is not numaric", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).value));
				Console.WriteLine(string.Format("Object Details Number Format : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).numberFormat));
				Console.WriteLine(string.Format("Object Details Data Type : {0}", cells.FirstOrDefault(v => v.dataType != DataType.NUMBER).dataType));
				LogUtils.ShowWarning("Not Every Values is numeric some assumptions are made in chart construction");
			}
			return new C.YValues(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
		}
		/// <summary>
		/// Set chart plot area
		/// </summary>
		internal void SetChartPlotArea(C.PlotArea plotArea)
		{
			chart.PlotArea = plotArea;
		}
		/// <summary>
		///
		/// </summary>
		internal C.PlotArea GetPlotArea()
		{
			return chart.PlotArea;
		}
		private static C.NumberingCache AddNumberCacheValue(ChartData[] cells)
		{
			try
			{
				C.NumberingCache numberingCache = new C.NumberingCache()
				{
					PointCount = new C.PointCount()
					{
						Val = (UInt32Value)(uint)cells.Length
					},
				};
				int count = 0;
				foreach (ChartData Cell in cells)
				{
					C.NumericPoint numericPoint = new C.NumericPoint()
					{
						Index = (UInt32Value)(uint)count,
						FormatCode = Cell.numberFormat,
					};
					if (Cell.dataType == DataType.NUMBER)
					{
						numericPoint.AppendChild(new C.NumericValue(Cell.value ?? "0"));
					}
					else
					{
						numericPoint.AppendChild(new C.NumericValue("0"));
					}
					numberingCache.AppendChild(numericPoint);
					++count;
				}
				return numberingCache;
			}
			catch
			{
				throw new Exception("Chart. Numeric Ref Error");
			}
		}
		private static C.StringCache AddStringCacheValue(ChartData[] cells)
		{
			try
			{
				C.StringCache stringCache = new C.StringCache()
				{
					PointCount = new C.PointCount()
					{
						Val = (UInt32Value)(uint)cells.Length
					},
				};
				int count = 0;
				foreach (ChartData Cell in cells)
				{
					C.StringPoint stringPoint = new C.StringPoint()
					{
						Index = (UInt32Value)(uint)count
					};
					stringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
					stringCache.AppendChild(stringPoint);
					++count;
				}
				return stringCache;
			}
			catch
			{
				throw new Exception("Chart. String Ref Error");
			}
		}
		private C.Chart CreateChart()
		{
			C.Chart chart = new C.Chart()
			{
				PlotVisibleOnly = new C.PlotVisibleOnly()
				{
					Val = true
				},
				AutoTitleDeleted = new C.AutoTitleDeleted()
				{
					Val = true
				},
				DisplayBlanksAs = new C.DisplayBlanksAs()
				{
					Val = C.DisplayBlanksAsValues.Gap
				},
				ShowDataLabelsOverMaximum = new C.ShowDataLabelsOverMaximum()
				{
					Val = false
				}
			};
			if (chartSetting.chartLegendOptions.isEnableLegend)
			{
				chart.Legend = CreateChartLegend(chartSetting.chartLegendOptions);
			}
			if (chartSetting.titleOptions != null)
			{
				chart.Title = CreateTitle(chartSetting.titleOptions);
			}
			return chart;
		}
		/// <summary>
		///
		/// </summary>
		protected void Add3Dcontrol()
		{
			chart.View3D = CreateView3D();
			chart.Floor = CreateFloor();
			chart.SideWall = CreateSideWall();
			chart.BackWall = CreateBackWall();
		}
		private C.BackWall CreateBackWall()
		{
			return new C.BackWall()
			{
				Thickness = new C.Thickness() { Val = 0 },
				ShapeProperties = CreateChartShapeProperties(new ShapePropertiesModel()
				{
					shapeProperty3D = new ShapeProperty3D()
				}),
			};
		}
		private C.SideWall CreateSideWall()
		{
			return new C.SideWall()
			{
				Thickness = new C.Thickness() { Val = 0 },
				ShapeProperties = CreateChartShapeProperties(new ShapePropertiesModel()
				{
					shapeProperty3D = new ShapeProperty3D()
				}),
			};
		}
		private C.Floor CreateFloor()
		{
			return new C.Floor()
			{
				Thickness = new C.Thickness() { Val = 0 },
				ShapeProperties = CreateChartShapeProperties(new ShapePropertiesModel()
				{
					shapeProperty3D = new ShapeProperty3D()
				}),
			};
		}
		private C.View3D CreateView3D()
		{
			return new C.View3D()
			{
				RotateX = new C.RotateX() { Val = 15 },
				RotateY = new C.RotateY() { Val = 15 },
				DepthPercent = new C.DepthPercent() { Val = 100 },
				RightAngleAxes = new C.RightAngleAxes() { Val = true },
			};
		}
		private C.Legend CreateChartLegend(ChartLegendOptions chartLegendOptions)
		{
			C.Legend legend = new C.Legend();
			if (chartLegendOptions.manualLayout != null)
			{
				legend.Append(CreateLayout(chartLegendOptions.manualLayout));
			}
			C.LegendPositionValues legendPositionValue;
			switch (chartLegendOptions.legendPosition)
			{
				case ChartLegendOptions.LegendPositionValues.TOP_RIGHT:
					legendPositionValue = C.LegendPositionValues.TopRight;
					break;
				case ChartLegendOptions.LegendPositionValues.TOP:
					legendPositionValue = C.LegendPositionValues.Top;
					break;
				case ChartLegendOptions.LegendPositionValues.BOTTOM:
					legendPositionValue = C.LegendPositionValues.Bottom;
					break;
				case ChartLegendOptions.LegendPositionValues.LEFT:
					legendPositionValue = C.LegendPositionValues.Left;
					break;
				case ChartLegendOptions.LegendPositionValues.RIGHT:
					legendPositionValue = C.LegendPositionValues.Right;
					break;
				default:
					legendPositionValue = C.LegendPositionValues.Bottom;
					break;
			}
			legend.Append(new C.LegendPosition() { Val = legendPositionValue });
			legend.Append(new C.Overlay { Val = chartLegendOptions.isLegendChartOverLap });
			legend.Append(CreateChartShapeProperties());
			SolidFillModel solidFillModel = new SolidFillModel()
			{
				schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.TEXT_1,
					luminanceModulation = 65000,
					luminanceOffset = 35000
				}
			};
			if (chartLegendOptions.fontColor != null)
			{
				solidFillModel.hexColor = chartLegendOptions.fontColor;
				solidFillModel.schemeColorModel = null;
			}
			legend.Append(CreateChartTextProperties(new ChartTextPropertiesModel()
			{
				drawingBodyProperties = new DrawingBodyPropertiesModel()
				{
					rotation = 0,
					useParagraphSpacing = true,
					verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
					vertical = TextVerticalAlignmentValues.HORIZONTAL,
					wrap = TextWrappingValues.SQUARE,
					anchor = TextAnchoringValues.CENTER,
					anchorCenter = true,
				},
				drawingParagraph = new DrawingParagraphModel()
				{
					paragraphPropertiesModel = new ParagraphPropertiesModel()
					{
						defaultRunProperties = new DefaultRunPropertiesModel()
						{
							solidFill = solidFillModel,
							complexScriptFont = "+mn-cs",
							eastAsianFont = "+mn-ea",
							latinFont = "+mn-lt",
							fontSize = ConverterUtils.FontSizeToFontSize(chartLegendOptions.fontSize),
							isBold = chartLegendOptions.isBold,
							isItalic = chartLegendOptions.isItalic,
							underline = chartLegendOptions.underLineValues,
							strike = chartLegendOptions.strikeValues,
							kerning = 1200,
							baseline = 0,
						}
					}
				}
			}));
			return legend;
		}
		private static C.ChartSpace CreateChartSpace()
		{
			C.ChartSpace chartSpace = new C.ChartSpace();
			chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			chartSpace.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
			chartSpace.RoundedCorners = new C.RoundedCorners()
			{
				Val = false
			};
			chartSpace.Date1904 = new C.Date1904()
			{
				Val = false
			};
			chartSpace.EditingLanguage = new C.EditingLanguage()
			{
				Val = "en-IN"
			};
			return chartSpace;
		}
		/// <summary>
		///
		/// </summary>
		protected static A.Field CreateField(string type, string text)
		{
			return new A.Field(
				new A.RunProperties() { Language = "en-IN" },
				new A.ParagraphProperties(),
				new A.Text()
				{
					Text = text
				}
			)
			{ Type = type, Id = GeneratorUtils.GenerateNewGUID() };
		}
		private C.MajorGridlines CreateMajorGridLine()
		{
			return new C.MajorGridlines(CreateChartShapeProperties(new ShapePropertiesModel()
			{
				outline = new OutlineModel()
				{
					solidFill = new SolidFillModel()
					{
						schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.TEXT_1,
							luminanceModulation = 15000,
							luminanceOffset = 85000
						}
					},
					width = 9525,
					outlineCapTypeValues = OutlineCapTypeValues.FLAT,
					outlineLineTypeValues = OutlineLineTypeValues.SINGEL,
					outlineAlignmentValues = OutlineAlignmentValues.CENTER
				}
			}));
		}
		private C.MinorGridlines CreateMinorGridLine()
		{
			return new C.MinorGridlines(CreateChartShapeProperties(new ShapePropertiesModel()
			{
				outline = new OutlineModel()
				{
					solidFill = new SolidFillModel()
					{
						schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.TEXT_1,
							luminanceModulation = 5000,
							luminanceOffset = 95000
						}
					},
					width = 9525,
					outlineCapTypeValues = OutlineCapTypeValues.FLAT,
					outlineLineTypeValues = OutlineLineTypeValues.SINGEL,
					outlineAlignmentValues = OutlineAlignmentValues.CENTER
				}
			}));
		}
		private C.Title CreateTitle(ChartTitleModel titleModel)
		{
			SolidFillModel solidFillModel = new SolidFillModel()
			{
				schemeColorModel = new SchemeColorModel()
				{
					themeColorValues = ThemeColorValues.TEXT_1
				}
			};
			if (titleModel.fontColor != null)
			{
				solidFillModel.hexColor = titleModel.fontColor;
				solidFillModel.schemeColorModel = null;
			}
			C.Title title = new C.Title(new C.ChartText(CreateChartRichText(new ChartTextPropertiesModel()
			{
				drawingBodyProperties = new DrawingBodyPropertiesModel()
				{
					anchor = TextAnchoringValues.CENTER,
					anchorCenter = true,
					useParagraphSpacing = true,
					vertical = TextVerticalAlignmentValues.HORIZONTAL,
					verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
					wrap = TextWrappingValues.SQUARE,
					rotation = 0,
				},
				drawingParagraph = new DrawingParagraphModel()
				{
					paragraphPropertiesModel = new ParagraphPropertiesModel(),
					drawingRuns = new List<DrawingRunModel>()
					{
						new DrawingRunModel(){
						text = titleModel.title,
						drawingRunProperties = new DrawingRunPropertiesModel()
						{
							solidFill = solidFillModel,
							fontSize = titleModel.fontSize,
							isBold = titleModel.isBold,
							isItalic = titleModel.isItalic,
							underline = titleModel.underLineValues,
						}
						}
					}.ToArray()
				}
			})));
			title.Append(new C.Overlay { Val = false });
			title.Append(CreateChartShapeProperties());
			return title;
		}
		/// <summary>
		///
		/// </summary>
		internal static C.Marker CreateMarker(MarkerModel marketModel)
		{
			C.Marker marker = new C.Marker()
			{
				Symbol = new C.Symbol() { Val = MarkerModel.GetMarkerStyleValues(marketModel.markerShapeType) },
			};
			if (marketModel.markerShapeType != MarkerShapeTypes.NONE)
			{
				marker.Size = new C.Size() { Val = (ByteValue)marketModel.size };
				marker.Append(CreateChartShapeProperties(marketModel.shapeProperties));
			}
			return marker;
		}
		internal C.Trendline CreateTrendLine(TrendLineModel trendLineModel)
		{
			C.Trendline trendLine = new C.Trendline()
			{
				TrendlineName = new C.TrendlineName(trendLineModel.trendLineName),
				TrendlineType = new C.TrendlineType() { Val = TrendLineModel.GetTrendlineValues(trendLineModel.trendLineType) },
				Forward = new C.Forward() { Val = trendLineModel.forecastForward },
				DisplayEquation = new C.DisplayEquation() { Val = trendLineModel.showEquation },
				DisplayRSquaredValue = new C.DisplayRSquaredValue() { Val = trendLineModel.showRSquareValue }
			};
			trendLine.Append(CreateChartShapeProperties(new ShapePropertiesModel()
			{
				outline = new OutlineModel()
				{
					width = 19050,
					outlineCapTypeValues = OutlineCapTypeValues.ROUND,
					solidFill = new SolidFillModel()
					{
						schemeColorModel = new SchemeColorModel()
						{
							themeColorValues = ThemeColorValues.ACCENT_1
						}
					},
					dashType = DrawingPresetLineDashValues.SYSTEM_DOT,
				},
				effectList = new EffectListModel()
			}));
			if (trendLineModel.trendLineType == TrendLineTypes.POLYNOMIAL)
			{
				trendLine.PolynomialOrder = new C.PolynomialOrder() { Val = (ByteValue)trendLineModel.secondaryValue };
			}
			if (trendLineModel.trendLineType == TrendLineTypes.MOVING_AVERAGE)
			{
				trendLine.Period = new C.Period { Val = (UInt32Value)(uint)trendLineModel.secondaryValue };
			}
			if (trendLineModel.setIntercept)
			{
				trendLine.Intercept = new C.Intercept() { Val = trendLineModel.interceptValue };
			}
			if (trendLineModel.setIntercept || trendLineModel.showEquation || trendLineModel.showRSquareValue)
			{
				trendLine.Append(CreateTrendLineLabel());
			}
			return trendLine;
		}
	}
}
