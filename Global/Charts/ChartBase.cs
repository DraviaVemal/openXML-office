// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;

/// <summary>
/// Chart Base Class Common to all charts. Class is only intended to get created by inherited classes
/// </summary>
public class ChartBase : CommonProperties
{
    #region Protected Fields

    /// <summary>
    /// Chart Data Groupings
    /// </summary>
    protected List<ChartDataGrouping> chartDataGroupings = new();

    /// <summary>
    /// Core chart settings common for every possible chart
    /// </summary>
    protected ChartSetting chartSetting;

    #endregion Protected Fields

    #region Private Fields

    private readonly C.Chart chart;

    private readonly C.ChartSpace openXMLChartSpace;

    #endregion Private Fields

    #region Protected Constructors

    /// <summary>
    /// Chartbase class constructor restricted only for inheritance use
    /// </summary>
    /// <param name="chartSetting">
    /// </param>
    protected ChartBase(ChartSetting chartSetting)
    {
        this.chartSetting = chartSetting;
        openXMLChartSpace = CreateChartSpace();
        chart = CreateChart();
        GetChartSpace().Append(chart);
        GetChartSpace().Append(new C.ExternalData(
            new C.AutoUpdate() { Val = false })
        { Id = "rId1" });
    }

    #endregion Protected Constructors

    #region Public Methods

    /// <summary>
    /// Get OpenXML ChartSpace
    /// </summary>
    /// <returns>
    /// </returns>
    public C.ChartSpace GetChartSpace()
    {
        return openXMLChartSpace;
    }

    #endregion Public Methods

    #region Protected Methods

    /// <summary>
    /// Create Bubble Size Axis for the chart
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    /// <exception cref="ArgumentException">
    /// </exception>
    protected C.BubbleSize CreateBubbleSizeAxisData(string formula, ChartData[] cells)
    {
        if (cells.All(v => v.dataType != DataType.NUMBER))
        {
            throw new ArgumentException("Bubble Size Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
    }

    /// <summary>
    /// Create Category Axis for the chart
    /// </summary>
    /// <param name="categoryAxisSetting">
    /// </param>
    /// <returns>
    /// </returns>
    protected C.CategoryAxis CreateCategoryAxis(CategoryAxisSetting categoryAxisSetting)
    {
        C.CategoryAxis CategoryAxis = new(
            new C.AxisId { Val = categoryAxisSetting.id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition
            {
                Val = categoryAxisSetting.axisPosition switch
                {
                    AxisPosition.LEFT => C.AxisPositionValues.Left,
                    AxisPosition.RIGHT => C.AxisPositionValues.Right,
                    AxisPosition.TOP => C.AxisPositionValues.Top,
                    _ => C.AxisPositionValues.Bottom
                }
            },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
        if (chartSetting.chartGridLinesOptions.isMajorCategoryLinesEnabled)
        {
            CategoryAxis.Append(CreateMajorGridLine());
        }
        if (chartSetting.chartGridLinesOptions.isMinorCategoryLinesEnabled)
        {
            CategoryAxis.Append(CreateMinorGridLine());
        }
        C.TextProperties TextProperties = new(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.ParagraphProperties(
                    CreateDefaultRunProperties(new()
                    {
                        fontSize = (int)categoryAxisSetting.fontSize * 100,
                        bold = categoryAxisSetting.isBold,
                        italic = categoryAxisSetting.isItalic,
                        baseline = 0
                    })
                ),
                new A.EndParagraphRunProperties { Language = "en-US" }
            )
        );
        CategoryAxis.Append(CreateChartShapeProperties());
        CategoryAxis.Append(TextProperties);
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
    /// Create Category Axis Data for the chart
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    protected C.CategoryAxisData CreateCategoryAxisData(string formula, ChartData[] cells)
    {
        if (cells.All(v => v.dataType == DataType.NUMBER))
        {
            return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
        }
        else
        {
            return new(new C.StringReference(new C.Formula(formula), AddStringCacheValue(cells)));
        }
    }

    /// <summary>
    /// Create Chart Styles for the chart
    /// </summary>
    /// <returns>
    /// </returns>
    protected static CS.ChartStyle CreateChartStyles()
    {
        ChartStyle ChartStyle = new();
        return ChartStyle.CreateChartStyles();
    }

    /// <summary>
    /// Create Color Styles for the chart
    /// </summary>
    /// <returns>
    /// </returns>
    protected static CS.ColorStyle CreateColorStyles()
    {
        ChartColor ChartColor = new();
        return Global.ChartColor.CreateColorStyles();
    }

    /// <summary>
    /// Create Data Labels for the chart
    /// </summary>
    /// <param name="chartDataLabel">
    /// </param>
    /// <param name="dataLabelCount">
    /// </param>
    /// <returns>
    /// </returns>
    protected C.DataLabels CreateDataLabels(ChartDataLabel chartDataLabel, int? dataLabelCount = 0)
    {
        C.DataLabels DataLabels = new();
        for (int i = 0; i < dataLabelCount; i++)
        {
            A.Paragraph Paragraph = new(CreateField("CELLRANGE", "[CELLRANGE]"));
            if (chartDataLabel.showSeriesName)
            {
                Paragraph.Append(new TextBoxBase(
                    new TextBoxSetting()
                    {
                        text = chartDataLabel.separator
                    }).GetTextBoxBaseRun());
                Paragraph.Append(CreateField("SERIESNAME", "[SERIES NAME]"));
            }
            if (chartDataLabel.showCategoryName)
            {
                Paragraph.Append(new TextBoxBase(
                    new TextBoxSetting()
                    {
                        text = chartDataLabel.separator
                    }).GetTextBoxBaseRun());
                Paragraph.Append(CreateField("CATEGORYNAME", "[CATEGORY NAME]"));
            }
            if (chartDataLabel.showValue)
            {
                Paragraph.Append(new TextBoxBase(
                    new TextBoxSetting()
                    {
                        text = chartDataLabel.separator
                    }).GetTextBoxBaseRun());
                Paragraph.Append(CreateField("VALUE", "[VALUE]"));
            }
            Paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });
            DataLabels.Append(new C.DataLabel(
                new C.Index() { Val = (uint)i },
                new C.SeriesText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        Paragraph
                    )
                ),
                new C.ShowLegendKey { Val = chartDataLabel.showLegendKey },
                new C.ShowValue { Val = chartDataLabel.showValue },
                new C.ShowCategoryName { Val = chartDataLabel.showCategoryName },
                new C.ShowSeriesName { Val = chartDataLabel.showSeriesName },
                new C.ShowPercent() { Val = true },
                new C.ShowBubbleSize() { Val = true },
                new C.Separator(chartDataLabel.separator)
            ));
        }
        DataLabels.Append(new C.ShowLegendKey { Val = chartDataLabel.showLegendKey },
            new C.ShowValue { Val = chartDataLabel.showValue },
            new C.ShowCategoryName { Val = chartDataLabel.showCategoryName },
            new C.ShowSeriesName { Val = chartDataLabel.showSeriesName },
            new C.ShowPercent { Val = false },
            new C.ShowBubbleSize() { Val = true },
            new C.Separator(chartDataLabel.separator),
            new C.ShowLeaderLines() { Val = false });
        return DataLabels;
    }

    /// <summary>
    /// Create Data Labels Range for the chart.Used in value from Column
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    protected C15.DataLabelsRange CreateDataLabelsRange(string formula, ChartData[] cells)
    {
        return new(new C15.Formula(formula), AddDataLabelCacheValue(cells));
    }

    /// <summary>
    /// Create Data Series for the chart
    /// </summary>
    /// <param name="dataCols">
    /// </param>
    /// <param name="chartDataSetting">
    /// </param>
    /// <returns>
    /// </returns>
    /// <exception cref="ArgumentException">
    /// </exception>
    protected List<ChartDataGrouping> CreateDataSeries(ChartData[][] dataCols, ChartDataSetting chartDataSetting)
    {
        List<uint> SeriesColumns = new();
        for (uint col = chartDataSetting.chartDataColumnStart + 1; col <= (chartDataSetting.chartDataColumnEnd == 0 ? dataCols.Length - 1 : chartDataSetting.chartDataColumnEnd); col++)
        {
            SeriesColumns.Add(col);
        }
        if ((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : chartDataSetting.chartDataRowEnd) - chartDataSetting.chartDataRowStart < 1 || (chartDataSetting.chartDataColumnEnd == 0 ? dataCols.Length : chartDataSetting.chartDataColumnEnd) - chartDataSetting.chartDataColumnStart < 1)
        {
            throw new ArgumentException("Data Series Invalid Range");
        }
        for (int i = 0; i < SeriesColumns.Count; i++)
        {
            uint Column = SeriesColumns[i];
            List<ChartData> XaxisCells = ((ChartData[]?)dataCols[chartDataSetting.chartDataColumnStart].Clone()!).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
            List<ChartData> YaxisCells = ((ChartData[]?)dataCols[Column].Clone()!).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
            ChartDataGrouping ChartDataGrouping = new()
            {
                seriesHeaderFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${chartDataSetting.chartDataRowStart + 1}",
                seriesHeaderCells = ((ChartData[]?)dataCols[Column].Clone()!)[chartDataSetting.chartDataRowStart],
                xAxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)chartDataSetting.chartDataColumnStart + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)chartDataSetting.chartDataColumnStart + 1)}${chartDataSetting.chartDataRowStart + XaxisCells.Count + 1}",
                xAxisCells = XaxisCells.ToArray(),
                yAxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${chartDataSetting.chartDataRowStart + YaxisCells.Count + 1}",
                yAxisCells = YaxisCells.ToArray(),
            };
            if (chartDataSetting.is3Ddata)
            {
                i++;
                Column = SeriesColumns[i];
                List<ChartData> ZaxisCells = ((ChartData[]?)dataCols[Column].Clone()!).Skip((int)chartDataSetting.chartDataRowStart + 1).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
                ChartDataGrouping.zAxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${chartDataSetting.chartDataRowStart + ZaxisCells.Count + 1}";
                ChartDataGrouping.zAxisCells = ZaxisCells.ToArray();
            }
            if (chartDataSetting.valueFromColumn.TryGetValue(Column, out uint DataValueColumn))
            {
                List<ChartData> DataLabelCells = ((ChartData[]?)dataCols[DataValueColumn].Clone()!).Skip((int)chartDataSetting.chartDataRowStart).Take((chartDataSetting.chartDataRowEnd == 0 ? dataCols[0].Length : (int)chartDataSetting.chartDataRowEnd) - (int)chartDataSetting.chartDataRowStart).ToList();
                ChartDataGrouping.dataLabelFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1)}${chartDataSetting.chartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1)}${chartDataSetting.chartDataRowStart + DataLabelCells.Count + 1}";
                ChartDataGrouping.dataLabelCells = DataLabelCells.ToArray();
            }
            chartDataGroupings.Add(ChartDataGrouping);
        }
        return chartDataGroupings;
    }

    /// <summary>
    /// Create Series Text for the chart
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    protected C.SeriesText CreateSeriesText(string formula, ChartData[] cells)
    {
        return new(new C.StringReference(new C.Formula(formula), AddStringCacheValue(cells)));
    }

    /// <summary>
    /// Create Value Axis for the chart
    /// </summary>
    /// <param name="valueAxisSetting">
    /// </param>
    /// <returns>
    /// </returns>
    protected C.ValueAxis CreateValueAxis(ValueAxisSetting valueAxisSetting)
    {
        C.ValueAxis ValueAxis = new(
            new C.AxisId { Val = valueAxisSetting.id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition
            {
                Val = valueAxisSetting.axisPosition switch
                {
                    AxisPosition.LEFT => C.AxisPositionValues.Left,
                    AxisPosition.RIGHT => C.AxisPositionValues.Right,
                    AxisPosition.TOP => C.AxisPositionValues.Top,
                    _ => C.AxisPositionValues.Bottom
                }
            });
        if (chartSetting.chartGridLinesOptions.isMajorValueLinesEnabled)
        {
            ValueAxis.Append(CreateMajorGridLine());
        }
        if (chartSetting.chartGridLinesOptions.isMinorValueLinesEnabled)
        {
            ValueAxis.Append(CreateMinorGridLine());
        }
        C.TextProperties TextProperties = new(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.ParagraphProperties(
                    CreateDefaultRunProperties(new()
                    {
                        fontSize = (int)valueAxisSetting.fontSize * 100,
                        bold = valueAxisSetting.isBold,
                        italic = valueAxisSetting.isItalic,
                        baseline = 0
                    })
                ),
                new A.EndParagraphRunProperties { Language = "en-US" }
            )
        );
        ValueAxis.Append(
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
        ValueAxis.Append(CreateChartShapeProperties());
        ValueAxis.Append(TextProperties);
        ValueAxis.Append(
            new C.CrossingAxis { Val = valueAxisSetting.crossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        return ValueAxis;
    }

    /// <summary>
    /// Create Value Axis Data for the chart
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    /// <exception cref="ArgumentException">
    /// </exception>
    protected C.Values CreateValueAxisData(string formula, ChartData[] cells)
    {
        if (cells.All(v => v.dataType != DataType.NUMBER))
        {
            throw new ArgumentException("Value Axis Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
    }

    /// <summary>
    /// Create X Axis Data for the chart
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    /// <exception cref="ArgumentException">
    /// </exception>
    protected C.XValues CreateXValueAxisData(string formula, ChartData[] cells)
    {
        if (cells.All(v => v.dataType != DataType.NUMBER))
        {
            throw new ArgumentException("X Axis Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
    }

    /// <summary>
    /// Create Y Axis Data for the chart
    /// </summary>
    /// <param name="formula">
    /// </param>
    /// <param name="cells">
    /// </param>
    /// <returns>
    /// </returns>
    /// <exception cref="ArgumentException">
    /// </exception>
    protected C.YValues CreateYValueAxisData(string formula, ChartData[] cells)
    {
        if (cells.All(v => v.dataType != DataType.NUMBER))
        {
            throw new ArgumentException("Y Axis Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(formula), AddNumberCacheValue(cells)));
    }

    /// <summary>
    /// Set chart plot area
    /// </summary>
    /// <param name="plotArea">
    /// </param>
    protected void SetChartPlotArea(C.PlotArea plotArea)
    {
        chart.PlotArea = plotArea;
    }

    #endregion Protected Methods

    #region Private Methods

    private static C15.DataLabelsRangeChache AddDataLabelCacheValue(ChartData[] cells)
    {
        try
        {
            C15.DataLabelsRangeChache DataLabelsRangeChache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in cells)
            {
                C.StringPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
                DataLabelsRangeChache.AppendChild(StringPoint);
                ++count;
            }
            return DataLabelsRangeChache;
        }
        catch
        {
            throw new Exception("Chart. Data Label Ref Error");
        }
    }

    private static C.NumberingCache AddNumberCacheValue(ChartData[] cells)
    {
        try
        {
            C.NumberingCache NumberingCache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in cells)
            {
                C.NumericPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count,
                    FormatCode = Cell.numberFormat
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
                NumberingCache.AppendChild(StringPoint);
                ++count;
            }
            return NumberingCache;
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
            C.StringCache StringCache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in cells)
            {
                C.StringPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.value ?? ""));
                StringCache.AppendChild(StringPoint);
                ++count;
            }
            return StringCache;
        }
        catch
        {
            throw new Exception("Chart. String Ref Error");
        }
    }

    private C.Chart CreateChart()
    {
        C.Chart Chart = new()
        {
            PlotVisibleOnly = new C.PlotVisibleOnly()
            {
                Val = true
            },
            AutoTitleDeleted = new C.AutoTitleDeleted()
            {
                Val = false
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
            Chart.Legend = CreateChartLegend(chartSetting.chartLegendOptions);
        }
        if (chartSetting.title != null)
        {
            Chart.Title = CreateTitle(chartSetting.title);
        }
        return Chart;
    }

    private C.Legend CreateChartLegend(ChartLegendOptions chartLegendOptions)
    {
        C.Legend legend = new();
        legend.Append(new C.LegendPosition()
        {
            Val = chartLegendOptions.legendPosition switch
            {
                ChartLegendOptions.LegendPositionValues.TOP_RIGHT => C.LegendPositionValues.TopRight,
                ChartLegendOptions.LegendPositionValues.TOP => C.LegendPositionValues.Top,
                ChartLegendOptions.LegendPositionValues.BOTTOM => C.LegendPositionValues.Bottom,
                ChartLegendOptions.LegendPositionValues.LEFT => C.LegendPositionValues.Left,
                ChartLegendOptions.LegendPositionValues.RIGHT => C.LegendPositionValues.Right,
                _ => C.LegendPositionValues.Bottom
            }
        });
        legend.Append(new C.Overlay { Val = false });
        legend.Append(CreateChartShapeProperties());
        C.TextProperties TextProperties = new();
        TextProperties.Append(new A.BodyProperties()
        {
            Rotation = 0,
            UseParagraphSpacing = true,
            VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal,
            Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true
        });
        TextProperties.Append(new A.ListStyle());
        A.Paragraph Paragraph = new();
        A.ParagraphProperties ParagraphProperties = new();
        ParagraphProperties.Append(CreateDefaultRunProperties(new()
        {
            solidFill = new()
            {
                schemeColorModel = new()
                {
                    themeColorValues = ThemeColorValues.TEXT_1,
                    luminanceModulation = 65000,
                    luminanceOffset = 35000
                }
            },
            complexScriptFont = "+mn-cs",
            eastAsianFont = "+mn-ea",
            latinFont = "+mn-lt",
            fontSize = (int)chartLegendOptions.fontSize * 100,
            bold = chartLegendOptions.isBold,
            italic = chartLegendOptions.isItalic,
            underline = UnderLineValues.NONE,
            strike = StrikeValues.NO_STRIKE,
            kerning = 1200,
            baseline = 0,
        }));
        Paragraph.Append(ParagraphProperties);
        Paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });
        TextProperties.Append(Paragraph);
        legend.Append(TextProperties);
        return legend;
    }

    private static C.ChartSpace CreateChartSpace()
    {
        C.ChartSpace ChartSpace = new();
        ChartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        ChartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        ChartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        ChartSpace.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
        ChartSpace.RoundedCorners = new C.RoundedCorners()
        {
            Val = false
        };
        ChartSpace.Date1904 = new C.Date1904()
        {
            Val = false
        };
        ChartSpace.EditingLanguage = new C.EditingLanguage()
        {
            Val = "en-US"
        };
        return ChartSpace;
    }

    private static A.Field CreateField(string type, string text)
    {
        return new A.Field(
            new A.RunProperties() { Language = "en-US" },
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
        return new(CreateChartShapeProperties(new()
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
        return new(CreateChartShapeProperties(new()
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

    private C.Title CreateTitle(string strTitle)
    {
        C.RichText RichText = new();
        RichText.Append(new A.BodyProperties()
        {
            Anchor = A.TextAnchoringTypeValues.Center,
            AnchorCenter = true,
            Rotation = 0,
            UseParagraphSpacing = true,
            Vertical = A.TextVerticalValues.Horizontal,
            VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Wrap = A.TextWrappingValues.Square
        });
        RichText.Append(new A.ListStyle());
        RichText.Append(
            new A.Paragraph(new A.ParagraphProperties(CreateDefaultRunProperties()),
            new TextBoxBase(new TextBoxSetting()
            {
                text = strTitle ?? "Chart Title"
            }).GetTextBoxBaseRun()));
        C.Title title = new(new C.ChartText(RichText));
        title.Append(new C.Overlay { Val = false });
        title.Append(CreateChartShapeProperties());
        return title;
    }

    /// <summary>
    /// 
    /// </summary>
    protected C.Marker CreateMarker(MarkerModel marketModel)
    {
        C.Marker marker = new()
        {
            Symbol = new() { Val = MarkerModel.GetMarkerStyleValues(marketModel.markerShapeValues) },
        };
        if (marketModel.markerShapeValues != MarkerModel.MarkerShapeValues.NONE)
        {
            marker.Size = new() { Val = (ByteValue)marketModel.size };
            marker.Append(CreateChartShapeProperties(marketModel.shapeProperties));
        }
        return marker;
    }
    #endregion Private Methods
}