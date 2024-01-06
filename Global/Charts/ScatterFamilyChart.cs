using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class ScatterFamilyChart : ChartBase
    {
        #region Protected Fields

        protected ScatterChartSetting ScatterChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        protected ScatterFamilyChart(ScatterChartSetting ScatterChartSetting, ChartData[][] DataCols) : base(ScatterChartSetting)
        {
            this.ScatterChartSetting = ScatterChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.ScatterChart ScatterChart = new(
                new C.ScatterStyle
                {
                    Val = ScatterChartSetting.ScatterChartTypes switch
                    {
                        ScatterChartTypes.SCATTER_SMOOTH => C.ScatterStyleValues.Smooth,
                        ScatterChartTypes.SCATTER_SMOOTH_MARKER => C.ScatterStyleValues.SmoothMarker,
                        ScatterChartTypes.SCATTER_STRIGHT => C.ScatterStyleValues.Line,
                        ScatterChartTypes.SCATTER_STRIGHT_MARKER => C.ScatterStyleValues.LineMarker,
                        // Clusted
                        _ => C.ScatterStyleValues.LineMarker,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                C.Marker Marker = new[] { ScatterChartTypes.SCATTER, ScatterChartTypes.SCATTER_SMOOTH_MARKER, ScatterChartTypes.SCATTER_STRIGHT_MARKER }.Contains(ScatterChartSetting.ScatterChartTypes) ? new(
                    new C.Symbol { Val = ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.SCATTER ? C.MarkerStyleValues.Auto : C.MarkerStyleValues.Circle },
                    new C.Size { Val = 5 },
                    new C.ShapeProperties(
                        new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(seriesIndex % 6) + 1}") }),
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(seriesIndex % 6) + 1}") })),
                        new A.EffectList()
                    )) :
                    new(new C.Symbol()
                    {
                        Val = C.MarkerStyleValues.None
                    });
                ScatterChart.Append(CreateScatterChartSeries(seriesIndex,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}${DataCols[0].Length}",
                    col.Skip(1).ToArray(),
                    Marker,
                     ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.SCATTER ? new A.Outline(new A.NoFill()) : new A.Outline(GetSolidFill(ScatterChartSetting.ScatterChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex)),
                    GetDataLabels(ScatterChartSetting, seriesIndex)
                ));
                seriesIndex++;
            }
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false });
            ScatterChart.Append(DataLabels);
            ScatterChart.Append(new C.Smooth { Val = false });
            ScatterChart.Append(new C.AxisId { Val = 1362418656 });
            ScatterChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(ScatterChart);
            plotArea.Append(CreateValueAxis(1362418656, C.AxisPositionValues.Bottom));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels CreateDataLabel(ScatterChartDataLabel ScatterChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = ScatterChartDataLabel.DataLabelPosition != ScatterChartDataLabel.eDataLabelPosition.NONE },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            if (ScatterChartDataLabel.DataLabelPosition != ScatterChartDataLabel.eDataLabelPosition.NONE)
            {
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = ScatterChartDataLabel.DataLabelPosition switch
                    {
                        ScatterChartDataLabel.eDataLabelPosition.LEFT => C.DataLabelPositionValues.Left,
                        ScatterChartDataLabel.eDataLabelPosition.RIGHT => C.DataLabelPositionValues.Right,
                        ScatterChartDataLabel.eDataLabelPosition.ABOVE => C.DataLabelPositionValues.Top,
                        ScatterChartDataLabel.eDataLabelPosition.BELOW => C.DataLabelPositionValues.Bottom,
                        //Center
                        _ => C.DataLabelPositionValues.Center,
                    }
                }, 0);
                DataLabels.InsertAt(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()), 0);
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = 1197,
                    Bold = false,
                    Italic = false,
                    Underline = A.TextUnderlineValues.None,
                    Strike = A.TextStrikeValues.NoStrike,
                    Kerning = 1200,
                    Baseline = 0
                }), new A.EndParagraphRunProperties() { Language = "en-US" });
                DataLabels.InsertAt(new C.TextProperties(new A.BodyProperties(new A.ShapeAutoFit())
                {
                    Rotation = 0,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    LeftInset = 38100,
                    TopInset = 19050,
                    RightInset = 38100,
                    BottomInset = 19050,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                }, new A.ListStyle(),
               Paragraph), 0);
            }
            return DataLabels;
        }

        private C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, string seriesTextFormula, ChartData[] seriesTextCells,
                                                        string xFormula, ChartData[] xCells, string yFormula,
                                                        ChartData[] yCells, C.Marker Marker, A.Outline Outline,
                                                        C.DataLabels DataLabels)
        {
            C.ScatterChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))),
                Marker);
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(Outline);
            ShapeProperties.Append(new A.EffectList());
            series.Append(DataLabels);
            series.Append(ShapeProperties);
            series.Append(new C.XValues(new C.NumberReference(new C.Formula(xFormula), AddNumberCacheValue(xCells, null))));
            series.Append(new C.YValues(new C.NumberReference(new C.Formula(yFormula), AddNumberCacheValue(yCells, null))));
            series.Append(new C.Smooth() { Val = new[] { ScatterChartTypes.SCATTER_SMOOTH, ScatterChartTypes.SCATTER_SMOOTH_MARKER }.Contains(ScatterChartSetting.ScatterChartTypes) });
            return series;
        }

        private C.DataLabels GetDataLabels(ScatterChartSetting ScatterChartSetting, int index)
        {
            if (index < ScatterChartSetting.ScatterChartSeriesSettings.Count)
            {
                return CreateDataLabel(ScatterChartSetting.ScatterChartSeriesSettings?[index]?.ScatterChartDataLabel ?? new ScatterChartDataLabel());
            }
            return CreateDataLabel(new ScatterChartDataLabel());
        }

        #endregion Private Methods
    }
}