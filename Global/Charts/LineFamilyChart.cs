using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class LineFamilyChart : ChartBase
    {
        #region Protected Methods
        protected LineChartDataLabel LineChartDataLabel = new();
        protected C.PlotArea CreateChartPlotArea(ChartData[][] DataCols, C.GroupingValues groupingValue, bool isMarkerEnabled = false)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.LineChart LineChart = new(
                new C.Grouping { Val = groupingValue },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                C.Marker Marker = isMarkerEnabled ? new(
                    new C.Symbol { Val = C.MarkerStyleValues.Circle },
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
                LineChart.Append(CreateLineChartSeries(seriesIndex,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}${DataCols[0].Length}",
                    col.Skip(1).ToArray(),
                    $"accent{(seriesIndex % 6) + 1}",
                    Marker
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
            LineChart.Append(DataLabels);
            LineChart.Append(new C.Smooth { Val = false });
            LineChart.Append(new C.AxisId { Val = 1362418656 });
            LineChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(LineChart);
            plotArea.Append(CreateCategoryAxis(1362418656));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        #endregion Protected Methods

        #region Private Methods
        private C.DataLabels CreateDataLabel()
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = LineChartDataLabel.DataLabelPosition != LineChartDataLabel.eDataLabelPosition.NONE },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            if (LineChartDataLabel.DataLabelPosition != LineChartDataLabel.eDataLabelPosition.NONE)
            {
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = LineChartDataLabel.DataLabelPosition switch
                    {
                        LineChartDataLabel.eDataLabelPosition.LEFT => C.DataLabelPositionValues.Left,
                        LineChartDataLabel.eDataLabelPosition.RIGHT => C.DataLabelPositionValues.Right,
                        LineChartDataLabel.eDataLabelPosition.ABOVE => C.DataLabelPositionValues.Top,
                        LineChartDataLabel.eDataLabelPosition.BELOW => C.DataLabelPositionValues.Bottom,
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
        private C.LineChartSeries CreateLineChartSeries(int seriesIndex, string seriesTextFormula, ChartData[] seriesTextCells, string categoryFormula, ChartData[] categoryCells, string valueFormula, ChartData[] valueCells, string accent, C.Marker Marker)
        {
            C.LineChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))),
                Marker);
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues(accent) }), new A.Round()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(CreateDataLabel());
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            return series;
        }

        #endregion Private Methods
    }
}