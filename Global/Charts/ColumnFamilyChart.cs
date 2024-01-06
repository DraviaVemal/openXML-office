using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class ColumnFamilyChart : ChartBase
    {
        #region Protected Fields

        protected ColumnChartSetting ColumnChartSetting;

        #endregion Protected Fields

        #region Public Constructors

        public ColumnFamilyChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting)
        {
            this.ColumnChartSetting = ColumnChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Public Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.BarChart ColumnChart = new(
                new C.BarDirection { Val = C.BarDirectionValues.Column },
                new C.BarGrouping
                {
                    Val = ColumnChartSetting.ColumnChartTypes switch
                    {
                        ColumnChartTypes.STACKED => C.BarGroupingValues.Stacked,
                        ColumnChartTypes.PERCENT_STACKED => C.BarGroupingValues.PercentStacked,
                        // Clusted
                        _ => C.BarGroupingValues.Clustered,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                ColumnChart.Append(CreateColumnChartSeries(seriesIndex,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}${DataCols[0].Length}",
                    col.Skip(1).ToArray(),
                    GetSolidFill(ColumnChartSetting.ColumnChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels(ColumnChartSetting, seriesIndex)
                ));
                seriesIndex++;
            }
            if (ColumnChartSetting.ColumnChartTypes == ColumnChartTypes.CLUSTERED)
            {
                ColumnChart.Append(new C.GapWidth { Val = 219 });
                ColumnChart.Append(new C.Overlap { Val = -27 });
            }
            else
            {
                ColumnChart.Append(new C.GapWidth { Val = 150 });
                ColumnChart.Append(new C.Overlap { Val = 100 });
            }
            ColumnChart.Append(new C.AxisId { Val = 1362418656 });
            ColumnChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(ColumnChart);
            plotArea.Append(CreateCategoryAxis(1362418656));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.BarChartSeries CreateColumnChartSeries(int seriesIndex, string seriesTextFormula, ChartData[] seriesTextCells,
                                                        string categoryFormula, ChartData[] categoryCells, string valueFormula,
                                                        ChartData[] valueCells, A.SolidFill SolidFill, C.DataLabels DataLabels)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(SolidFill);
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(DataLabels);
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            series.Append(new C.Smooth()
            {
                Val = false
            });
            return series;
        }

        private C.DataLabels CreateDataLabel(ColumnChartDataLabel ColumnChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = ColumnChartDataLabel.DataLabelPosition != ColumnChartDataLabel.eDataLabelPosition.NONE },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            if (ColumnChartDataLabel.DataLabelPosition != ColumnChartDataLabel.eDataLabelPosition.NONE)
            {
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = ColumnChartDataLabel.DataLabelPosition switch
                    {
                        ColumnChartDataLabel.eDataLabelPosition.CENTER => C.DataLabelPositionValues.Center,
                        ColumnChartDataLabel.eDataLabelPosition.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                        ColumnChartDataLabel.eDataLabelPosition.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
                        _ => C.DataLabelPositionValues.OutsideEnd
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

        private C.DataLabels GetDataLabels(ColumnChartSetting ColumnChartSetting, int index)
        {
            if (index < ColumnChartSetting.ColumnChartSeriesSettings.Count)
            {
                return CreateDataLabel(ColumnChartSetting.ColumnChartSeriesSettings[index]?.ColumnChartDataLabel ?? new ColumnChartDataLabel());
            }
            return CreateDataLabel(new ColumnChartDataLabel());
        }

        #endregion Private Methods
    }
}