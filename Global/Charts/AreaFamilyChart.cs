using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class AreaFamilyChart : ChartBase
    {
        #region Protected Fields

        protected readonly AreaChartSetting AreaChartSetting;

        #endregion Protected Fields

        #region Public Constructors

        public AreaFamilyChart(AreaChartSetting AreaChartSetting, ChartData[][] DataCols) : base(AreaChartSetting)
        {
            this.AreaChartSetting = AreaChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Public Constructors

        #region Private Methods

        private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, string seriesTextFormula,
                                                        ChartData[] seriesTextCells, string categoryFormula, ChartData[] categoryCells,
                                                        string valueFormula, ChartData[] valueCells, A.SolidFill SolidFill,
                                                        C.DataLabels DataLabels)
        {
            C.AreaChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.Outline(SolidFill, new A.Outline(new A.NoFill())));
            ShapeProperties.Append(new A.EffectList());
            series.Append(DataLabels);
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            return series;
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.AreaChart AreaChart = new(
                new C.Grouping
                {
                    Val = AreaChartSetting.AreaChartTypes switch
                    {
                        AreaChartTypes.STACKED => C.GroupingValues.Stacked,
                        AreaChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
                        // Clusted
                        _ => C.GroupingValues.Standard,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                AreaChart.Append(CreateAreaChartSeries(seriesIndex,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 2)}${DataCols[0].Length}",
                    col.Skip(1).ToArray(),
                    GetSolidFill(AreaChartSetting.AreaChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels(seriesIndex)
                ));
                seriesIndex++;
            }
            AreaChart.Append(new C.AxisId { Val = 1362418656 });
            AreaChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(AreaChart);
            plotArea.Append(CreateCategoryAxis(1362418656));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels CreateDataLabel(AreaChartDataLabel AreaChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = AreaChartDataLabel.DataLabelPosition != AreaChartDataLabel.eDataLabelPosition.NONE },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            if (AreaChartDataLabel.DataLabelPosition != AreaChartDataLabel.eDataLabelPosition.NONE)
            {
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = AreaChartDataLabel.DataLabelPosition switch
                    {
                        //Show
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

        private C.DataLabels GetDataLabels(int index)
        {
            if (index < AreaChartSetting.AreaChartSeriesSettings.Count)
            {
                return CreateDataLabel(AreaChartSetting.AreaChartSeriesSettings?[index]?.AreaChartDataLabel ?? new AreaChartDataLabel());
            }
            return CreateDataLabel(new AreaChartDataLabel());
        }

        #endregion Private Methods
    }
}