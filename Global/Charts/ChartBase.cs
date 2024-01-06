using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;

public class ChartBase
{
    #region Protected Fields

    protected ChartSetting ChartSetting;
    protected List<ChartDataGrouping> ChartDataGroupings = new();
    #endregion Protected Fields

    #region Private Fields

    private C.Chart Chart;

    private C.ChartSpace OpenXMLChartSpace;

    #endregion Private Fields

    #region Protected Constructors

    protected ChartBase(ChartSetting ChartSetting)
    {
        this.ChartSetting = ChartSetting;
        OpenXMLChartSpace = CreateChartSpace();
        Chart = CreateChart();
        GetChartSpace().Append(Chart);
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

    protected C15.DataLabelsRangeChache AddDataLabelCacheValue(ChartData[] Cells)
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

    protected C.NumberingCache AddNumberCacheValue(ChartData[] Cells, ChartSeriesSetting? ChartSeriesSetting)
    {
        try
        {
            C.NumberingCache NumberingCache = new()
            {
                FormatCode = new C.FormatCode(ChartSeriesSetting?.NumberFormat ?? "General"),
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

    protected C.StringCache AddStringCacheValue(ChartData[] Cells)
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
    protected List<ChartDataGrouping> CreateDataSeries(C.BarChart ColumnChart, ChartData[][] DataCols, ChartDataSetting ChartDataSetting)
    {
        if (ChartDataSetting.ChartDataRowStart < 1)
        {
            throw new ArgumentOutOfRangeException("Data Range Cannot Be Less Than 0");
        }
        List<uint> SeriesColumns = new();
        long TotalSeriesCount = (ChartDataSetting.ChartDataColumnEnd == 0 ? DataCols.Length : ChartDataSetting.ChartDataColumnEnd) - ChartDataSetting.ChartDataColumnStart - ChartDataSetting.ValueFromColumn.Count;
        for (uint col = ChartDataSetting.ChartDataColumnStart; col < ChartDataSetting.ChartDataColumnEnd; col++)
        {
            if (!(ChartDataSetting.ValueFromColumn.TryGetValue(col, out _) || col == ChartDataSetting.ChartRowHeader))
            {
                SeriesColumns.Add(col);
            }
        }
        if (TotalSeriesCount != SeriesColumns.Count)
        {
            throw new ArgumentOutOfRangeException("Data Series Column Miss Match");
        }
        foreach (uint Column in SeriesColumns)
        {
            ChartDataGrouping ChartDataGrouping = new()
            {
                XaxisCells = (ChartData[]?)DataCols[ChartDataSetting.ChartRowHeader].Clone(),
                YaxisCells = (ChartData[]?)DataCols[Column].Clone(),
            };
            if (ChartDataSetting.ValueFromColumn.TryGetValue(Column, out uint DataValueColumn))
            {
                ChartDataGrouping.DataLabelCells = (ChartData[]?)DataCols[DataValueColumn].Clone();
            }
            ChartDataGroupings.Add(ChartDataGrouping);
        }
        return ChartDataGroupings;
    }
    protected C.CategoryAxis CreateCategoryAxis(UInt32Value axisId, C.AxisPositionValues? AxisPositionValues = null)
    {
        C.CategoryAxis CategoryAxis = new(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = AxisPositionValues ?? C.AxisPositionValues.Bottom },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = axisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 },
            new C.NoMultiLevelLabels { Val = false });
        C.ShapeProperties ShapeProperties = new();
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

    protected C.ValueAxis CreateValueAxis(UInt32Value axisId, C.AxisPositionValues? AxisPositionValues = null)
    {
        C.ValueAxis ValueAxis = new(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = AxisPositionValues ?? C.AxisPositionValues.Left },
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = axisId },
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
        C.ShapeProperties ShapeProperties = new();
        ShapeProperties.Append(new A.NoFill());
        ShapeProperties.Append(new A.Outline(new A.NoFill()));
        ShapeProperties.Append(new A.EffectList());
        ValueAxis.Append(ShapeProperties);
        return ValueAxis;
    }

    protected A.SolidFill GetSolidFill(List<string> FillColors, int index)
    {
        if (FillColors.Count > 0)
        {
            return new A.SolidFill(new A.RgbColorModelHex() { Val = FillColors[index % FillColors.Count] });
        }
        return new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(index % 6) + 1}") });
    }

    #endregion Protected Methods

    #region Private Methods

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
            Val = objChartLegendOptions.legendPosition switch
            {
                ChartLegendOptions.eLegendPosition.TOP_RIGHT => C.LegendPositionValues.TopRight,
                ChartLegendOptions.eLegendPosition.TOP => C.LegendPositionValues.Top,
                ChartLegendOptions.eLegendPosition.BOTTOM => C.LegendPositionValues.Bottom,
                ChartLegendOptions.eLegendPosition.LEFT => C.LegendPositionValues.Left,
                ChartLegendOptions.eLegendPosition.RIGHT => C.LegendPositionValues.Right,
                _ => C.LegendPositionValues.Bottom
            }
        });
        legend.Append(new C.Overlay { Val = false });
        C.ShapeProperties ShapeProperties = new();
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
            new TextBox(new TextBoxSetting()
            {
                Text = strTitle ?? "Chart Title"
            }).GetTextBoxRun()));
        C.Title title = new(new C.ChartText(RichText));
        title.Append(new C.Overlay { Val = false });
        C.ShapeProperties ShapeProperties = new();
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