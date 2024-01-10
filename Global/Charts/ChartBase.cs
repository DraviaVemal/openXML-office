/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;

public class ChartBase : CommonProperties
{
    #region Protected Fields

    protected List<ChartDataGrouping> ChartDataGroupings = new();
    protected ChartSetting ChartSetting;

    #endregion Protected Fields

    #region Private Fields

    private readonly C.Chart Chart;

    private readonly C.ChartSpace OpenXMLChartSpace;

    #endregion Private Fields

    #region Protected Constructors

    protected ChartBase(ChartSetting ChartSetting)
    {
        this.ChartSetting = ChartSetting;
        OpenXMLChartSpace = CreateChartSpace();
        Chart = CreateChart();
        GetChartSpace().Append(Chart);
        GetChartSpace().Append(new C.ExternalData(new C.AutoUpdate() { Val = false }) { Id = "rId1" });
    }

    #endregion Protected Constructors

    #region Public Methods

    public C.ChartSpace GetChartSpace()
    {
        return OpenXMLChartSpace;
    }

    public void SetChartPlotArea(C.PlotArea PlotArea)
    {
        Chart.PlotArea = PlotArea;
    }

    #endregion Public Methods

    #region Protected Methods

    protected C.BubbleSize CreateBubbleSizeAxisData(string Formula, ChartData[] Cells)
    {
        if (Cells.All(v => v.DataType != DataType.NUMBER))
        {
            throw new ArgumentException("Bubble Size Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(Formula), AddNumberCacheValue(Cells)));
    }

    protected C.CategoryAxis CreateCategoryAxis(CategoryAxisSetting CategoryAxisSetting)
    {
        C.CategoryAxis CategoryAxis = new(
            new C.AxisId { Val = CategoryAxisSetting.Id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition
            {
                Val = CategoryAxisSetting.AxisPosition switch
                {
                    AxisPosition.LEFT => C.AxisPositionValues.Left,
                    AxisPosition.RIGHT => C.AxisPositionValues.Right,
                    AxisPosition.TOP => C.AxisPositionValues.Top,
                    _ => C.AxisPositionValues.Bottom
                }
            },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = CategoryAxisSetting.CrossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 },
            new C.NoMultiLevelLabels { Val = false });
        C.ShapeProperties ShapeProperties = CreateShapeProperties();
        ShapeProperties.Append(new A.NoFill());
        ShapeProperties.Append(new A.Outline(new A.NoFill()));
        ShapeProperties.Append(new A.EffectList());
        if (ChartSetting.ChartGridLinesOptions.IsMajorCategoryLinesEnabled)
        {
            CategoryAxis.Append(CreateMajorGridLine());
        }
        if (ChartSetting.ChartGridLinesOptions.IsMinorCategoryLinesEnabled)
        {
            CategoryAxis.Append(CreateMinorGridLine());
        }
        CategoryAxis.Append(ShapeProperties);
        return CategoryAxis;
    }

    protected C.CategoryAxisData CreateCategoryAxisData(string Formula, ChartData[] Cells)
    {
        if (Cells.All(v => v.DataType == DataType.NUMBER))
        {
            return new(new C.NumberReference(new C.Formula(Formula), AddNumberCacheValue(Cells)));
        }
        else
        {
            return new(new C.StringReference(new C.Formula(Formula), AddStringCacheValue(Cells)));
        }
    }

    protected CS.ChartStyle CreateChartStyles()
    {
        ChartStyle ChartStyle = new();
        return ChartStyle.CreateChartStyles();
    }

    protected CS.ColorStyle CreateColorStyles()
    {
        ChartColor ChartColor = new();
        return ChartColor.CreateColorStyles();
    }

    protected C.DataLabels CreateDataLabels(ChartDataLabel ChartDataLabel, int? DataLabelCount = 0)
    {
        C.ExtensionList ExtensionList = new(
            new C.Extension(
                new C15.DataLabelFieldTable(),
                new C15.ShowDataLabelsRange() { Val = true }
            )
            {
                Uri = GeneratorUtils.GenerateNewGUID()
            }
        );
        C.DataLabels DataLabels = new(
            new C.ShowLegendKey { Val = ChartDataLabel.ShowLegendKey },
            new C.ShowValue { Val = ChartDataLabel.ShowValue },
            new C.ShowCategoryName { Val = ChartDataLabel.ShowCategoryName },
            new C.ShowSeriesName { Val = ChartDataLabel.ShowSeriesName },
            new C.ShowPercent { Val = false },
            new C.ShowLeaderLines() { Val = false },
            new C.Separator(ChartDataLabel.Separator),
            (OpenXmlElement)ExtensionList.Clone());
        for (int i = 0; i < DataLabelCount; i++)
        {
            A.Paragraph Paragraph = new(CreateField("CELLRANGE", "[CELLRANGE]"));
            if (ChartDataLabel.ShowSeriesName)
            {
                Paragraph.Append(new TextBoxBase(new TextBoxSetting()
                {
                    Text = ChartDataLabel.Separator
                }).GetTextBoxRun());
                Paragraph.Append(CreateField("SERIESNAME", "[SERIES NAME]"));
            }
            if (ChartDataLabel.ShowCategoryName)
            {
                Paragraph.Append(new TextBoxBase(new TextBoxSetting()
                {
                    Text = ChartDataLabel.Separator
                }).GetTextBoxRun());
                Paragraph.Append(CreateField("CATEGORYNAME", "[CATEGORY NAME]"));
            }
            if (ChartDataLabel.ShowValue)
            {
                Paragraph.Append(new TextBoxBase(new TextBoxSetting()
                {
                    Text = ChartDataLabel.Separator
                }).GetTextBoxRun());
                Paragraph.Append(CreateField("VALUE", "[VALUE]"));
            }
            Paragraph.Append(new A.EndParagraphRunProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex() { Val = "000000" }
                    ),
                    new A.Highlight(
                        new A.RgbColorModelHex() { Val = "FFFFFF" }
                    ),
                    new A.LatinFont() { Typeface = "Calibri (Body)" },
                    new A.EastAsianFont() { Typeface = "Calibri (Body)" },
                    new A.ComplexScriptFont() { Typeface = "Calibri (Body)" }
                )
            { Language = "", FontSize = 1800, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Dirty = false });
            DataLabels.Append(new C.DataLabel(
                new C.Index() { Val = (uint)i },
                new C.SeriesText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        Paragraph
                    )
                ),
                new C.ShowLegendKey { Val = ChartDataLabel.ShowLegendKey },
                new C.ShowValue { Val = ChartDataLabel.ShowValue },
                new C.ShowCategoryName { Val = ChartDataLabel.ShowCategoryName },
                new C.ShowSeriesName { Val = ChartDataLabel.ShowSeriesName },
                new C.Separator(ChartDataLabel.Separator),
                (OpenXmlElement)ExtensionList.Clone()
            ));
        }
        return DataLabels;
    }

    protected C15.DataLabelsRange CreateDataLabelsRange(string Formula, ChartData[] Cells)
    {
        return new(new C.Formula(Formula), AddDataLabelCacheValue(Cells));
    }

    protected List<ChartDataGrouping> CreateDataSeries(ChartData[][] DataCols, ChartDataSetting ChartDataSetting)
    {
        List<uint> SeriesColumns = new();
        for (uint col = ChartDataSetting.ChartDataColumnStart + 1; col <= (ChartDataSetting.ChartDataColumnEnd == 0 ? DataCols.Length - 1 : ChartDataSetting.ChartDataColumnEnd); col++)
        {
            SeriesColumns.Add(col);
        }
        if ((ChartDataSetting.ChartDataRowEnd == 0 ? DataCols[0].Length : ChartDataSetting.ChartDataRowEnd) - ChartDataSetting.ChartDataRowStart < 1 || (ChartDataSetting.ChartDataColumnEnd == 0 ? DataCols.Length : ChartDataSetting.ChartDataColumnEnd) - ChartDataSetting.ChartDataColumnStart < 1)
        {
            throw new ArgumentException("Data Series Invalid Range");
        }
        for (int i = 0; i < SeriesColumns.Count; i++)
        {
            uint Column = SeriesColumns[i];
            List<ChartData> XaxisCells = ((ChartData[]?)DataCols[ChartDataSetting.ChartDataColumnStart].Clone()!).Skip((int)ChartDataSetting.ChartDataRowStart + 1).Take((ChartDataSetting.ChartDataRowEnd == 0 ? DataCols[0].Length : (int)ChartDataSetting.ChartDataRowEnd) - (int)ChartDataSetting.ChartDataRowStart).ToList();
            List<ChartData> YaxisCells = ((ChartData[]?)DataCols[Column].Clone()!).Skip((int)ChartDataSetting.ChartDataRowStart + 1).Take((ChartDataSetting.ChartDataRowEnd == 0 ? DataCols[0].Length : (int)ChartDataSetting.ChartDataRowEnd) - (int)ChartDataSetting.ChartDataRowStart).ToList();
            ChartDataGrouping ChartDataGrouping = new()
            {
                SeriesHeaderFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${ChartDataSetting.ChartDataRowStart + 1}",
                SeriesHeaderCells = ((ChartData[]?)DataCols[Column].Clone()!)[ChartDataSetting.ChartDataRowStart],
                XaxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)ChartDataSetting.ChartDataColumnStart + 1)}${ChartDataSetting.ChartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)ChartDataSetting.ChartDataColumnStart + 1)}${ChartDataSetting.ChartDataRowStart + XaxisCells.Count + 1}",
                XaxisCells = XaxisCells.ToArray(),
                YaxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${ChartDataSetting.ChartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${ChartDataSetting.ChartDataRowStart + YaxisCells.Count + 1}",
                YaxisCells = YaxisCells.ToArray(),
            };
            if (ChartDataSetting.Is3Ddata)
            {
                i++;
                Column = SeriesColumns[i];
                List<ChartData> ZaxisCells = ((ChartData[]?)DataCols[Column].Clone()!).Skip((int)ChartDataSetting.ChartDataRowStart + 1).Take((ChartDataSetting.ChartDataRowEnd == 0 ? DataCols[0].Length : (int)ChartDataSetting.ChartDataRowEnd) - (int)ChartDataSetting.ChartDataRowStart).ToList();
                ChartDataGrouping.ZaxisFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${ChartDataSetting.ChartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)Column + 1)}${ChartDataSetting.ChartDataRowStart + ZaxisCells.Count + 1}";
                ChartDataGrouping.ZaxisCells = ZaxisCells.ToArray();
            }
            if (ChartDataSetting.ValueFromColumn.TryGetValue(Column, out uint DataValueColumn))
            {
                List<ChartData> DataLabelCells = ((ChartData[]?)DataCols[DataValueColumn].Clone()!).Skip((int)ChartDataSetting.ChartDataRowStart).Take((ChartDataSetting.ChartDataRowEnd == 0 ? DataCols[0].Length : (int)ChartDataSetting.ChartDataRowEnd) - (int)ChartDataSetting.ChartDataRowStart).ToList();
                ChartDataGrouping.DataLabelFormula = $"Sheet1!${ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1)}${ChartDataSetting.ChartDataRowStart + 2}:${ConverterUtils.ConvertIntToColumnName((int)DataValueColumn + 1)}${ChartDataSetting.ChartDataRowStart + DataLabelCells.Count + 1}";
                ChartDataGrouping.DataLabelCells = DataLabelCells.ToArray();
            }
            ChartDataGroupings.Add(ChartDataGrouping);
        }
        return ChartDataGroupings;
    }

    protected C.SeriesText CreateSeriesText(string Formula, ChartData[] Cells)
    {
        return new(new C.StringReference(new C.Formula(Formula), AddStringCacheValue(Cells)));
    }

    protected C.ShapeProperties CreateShapeProperties()
    {
        return new();
    }

    protected C.ValueAxis CreateValueAxis(ValueAxisSetting ValueAxisSetting)
    {
        C.ValueAxis ValueAxis = new(
            new C.AxisId { Val = ValueAxisSetting.Id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition
            {
                Val = ValueAxisSetting.AxisPosition switch
                {
                    AxisPosition.LEFT => C.AxisPositionValues.Left,
                    AxisPosition.RIGHT => C.AxisPositionValues.Right,
                    AxisPosition.TOP => C.AxisPositionValues.Top,
                    _ => C.AxisPositionValues.Bottom
                }
            },
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = ValueAxisSetting.CrossAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        if (ChartSetting.ChartGridLinesOptions.IsMajorValueLinesEnabled)
        {
            ValueAxis.Append(CreateMajorGridLine());
        }
        if (ChartSetting.ChartGridLinesOptions.IsMinorValueLinesEnabled)
        {
            ValueAxis.Append(CreateMinorGridLine());
        }
        C.ShapeProperties ShapeProperties = CreateShapeProperties();
        ShapeProperties.Append(new A.NoFill());
        ShapeProperties.Append(new A.Outline(new A.NoFill()));
        ShapeProperties.Append(new A.EffectList());
        ValueAxis.Append(ShapeProperties);
        return ValueAxis;
    }

    protected C.Values CreateValueAxisData(string Formula, ChartData[] Cells)
    {
        if (Cells.All(v => v.DataType != DataType.NUMBER))
        {
            throw new ArgumentException("Value Axis Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(Formula), AddNumberCacheValue(Cells)));
    }

    protected C.XValues CreateXValueAxisData(string Formula, ChartData[] Cells)
    {
        if (Cells.All(v => v.DataType != DataType.NUMBER))
        {
            throw new ArgumentException("X Axis Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(Formula), AddNumberCacheValue(Cells)));
    }

    protected C.YValues CreateYValueAxisData(string Formula, ChartData[] Cells)
    {
        if (Cells.All(v => v.DataType != DataType.NUMBER))
        {
            throw new ArgumentException("Y Axis Data Should Be numaric");
        }
        return new(new C.NumberReference(new C.Formula(Formula), AddNumberCacheValue(Cells)));
    }

    #endregion Protected Methods

    #region Private Methods

    private C15.DataLabelsRangeChache AddDataLabelCacheValue(ChartData[] Cells)
    {
        try
        {
            C15.DataLabelsRangeChache DataLabelsRangeChache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)Cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in Cells)
            {
                C.StringPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.Value ?? ""));
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

    private C.NumberingCache AddNumberCacheValue(ChartData[] Cells)
    {
        try
        {
            C.NumberingCache NumberingCache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)Cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in Cells)
            {
                C.NumericPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count,
                    FormatCode = Cell.NumberFormat
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.Value ?? ""));
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

    private C.StringCache AddStringCacheValue(ChartData[] Cells)
    {
        try
        {
            C.StringCache StringCache = new()
            {
                PointCount = new C.PointCount()
                {
                    Val = (UInt32Value)(uint)Cells.Length
                },
            };
            int count = 0;
            foreach (ChartData Cell in Cells)
            {
                C.StringPoint StringPoint = new()
                {
                    Index = (UInt32Value)(uint)count
                };
                StringPoint.AppendChild(new C.NumericValue(Cell.Value ?? ""));
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
        if (ChartSetting.ChartLegendOptions.IsEnableLegend)
        {
            Chart.Legend = CreateChartLegend(ChartSetting.ChartLegendOptions);
        }
        if (ChartSetting.Title != null)
        {
            Chart.Title = CreateTitle(ChartSetting.Title);
        }
        return Chart;
    }

    private C.Legend CreateChartLegend(ChartLegendOptions objChartLegendOptions)
    {
        C.Legend legend = new();
        legend.Append(new C.LegendPosition()
        {
            Val = objChartLegendOptions.LegendPosition switch
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
        C.ShapeProperties ShapeProperties = CreateShapeProperties();
        ShapeProperties.Append(new A.NoFill());
        A.Outline Outline = new();
        Outline.Append(new A.NoFill());
        ShapeProperties.Append(Outline);
        ShapeProperties.Append(new A.EffectList());
        legend.Append(ShapeProperties);
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
        A.DefaultRunProperties DefaultRunProperties = new()
        {
            FontSize = 1197,
            Bold = false,
            Italic = false,
            Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike,
            Kerning = 1200,
            Baseline = 0
        };
        DefaultRunProperties.Append(new A.SolidFill(new A.SchemeColor(
        new A.LuminanceModulation { Val = 65000 },
        new A.LuminanceOffset { Val = 35000 })
        { Val = A.SchemeColorValues.Text1 }));
        DefaultRunProperties.Append(new A.LatinFont { Typeface = "+mn-lt" });
        DefaultRunProperties.Append(new A.EastAsianFont { Typeface = "+mn-ea" });
        DefaultRunProperties.Append(new A.ComplexScriptFont { Typeface = "+mn-cs" });
        ParagraphProperties.Append(DefaultRunProperties);
        Paragraph.Append(ParagraphProperties);
        Paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });
        TextProperties.Append(Paragraph);
        legend.Append(TextProperties);
        return legend;
    }

    private C.ChartSpace CreateChartSpace()
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

    private A.Field CreateField(string type, string text)
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
        return new(new C.ShapeProperties(
                        new A.Outline(
                            new A.SolidFill(
                                new A.SchemeColor(
                                    new A.LuminanceModulation { Val = 15000 },
                                    new A.LuminanceOffset { Val = 85000 })
                                { Val = A.SchemeColorValues.Text1 }
                            )
                        )
                        {
                            Width = 9525,
                            CapType = A.LineCapValues.Flat,
                            CompoundLineType = A.CompoundLineValues.Single,
                            Alignment = A.PenAlignmentValues.Center
                        }
                    )
                );
    }

    private C.MinorGridlines CreateMinorGridLine()
    {
        return new(new C.ShapeProperties(
                        new A.Outline(
                            new A.SolidFill(
                                new A.SchemeColor(
                                    new A.LuminanceModulation { Val = 5000 },
                                    new A.LuminanceOffset { Val = 95000 })
                                { Val = A.SchemeColorValues.Text1 }
                            )
                        )
                        {
                            Width = 9525,
                            CapType = A.LineCapValues.Flat,
                            CompoundLineType = A.CompoundLineValues.Single,
                            Alignment = A.PenAlignmentValues.Center
                        }
                    )
                );
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
        A.DefaultRunProperties DefaultRunProperties = new();
        DefaultRunProperties.Append(new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 65000 }, new A.LuminanceOffset { Val = 35000 }) { Val = A.SchemeColorValues.Text1 }));
        DefaultRunProperties.Append(new A.LatinFont { Typeface = "+mn-lt" });
        DefaultRunProperties.Append(new A.EastAsianFont { Typeface = "+mn-ea" });
        DefaultRunProperties.Append(new A.ComplexScriptFont { Typeface = "+mn-cs" });
        RichText.Append(
            new A.Paragraph(new A.ParagraphProperties(DefaultRunProperties),
            new TextBoxBase(new TextBoxSetting()
            {
                Text = strTitle ?? "Chart Title"
            }).GetTextBoxRun()));
        C.Title title = new(new C.ChartText(RichText));
        title.Append(new C.Overlay { Val = false });
        C.ShapeProperties ShapeProperties = CreateShapeProperties();
        ShapeProperties.Append(new A.NoFill());
        A.Outline Outline = new();
        Outline.Append(new A.NoFill());
        ShapeProperties.Append(Outline);
        ShapeProperties.Append(new A.EffectList());
        title.Append(ShapeProperties);
        return title;
    }

    #endregion Private Methods
}