using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class BarFamilyChart : ChartBase
    {
        #region Protected Fields

        protected readonly BarChartSetting BarChartSetting;

        #endregion Protected Fields

        #region Public Constructors

        public BarFamilyChart(BarChartSetting BarChartSetting, ChartData[][] DataCols) : base(BarChartSetting)
        {
            this.BarChartSetting = BarChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Public Constructors

        #region Private Methods

        private C.BarChartSeries CreateBarChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, A.SolidFill SolidFill, C.DataLabels DataLabels)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(ChartDataGrouping.SeriesHeaderFormula!), AddStringCacheValue(new[] { ChartDataGrouping.SeriesHeaderCells! }))),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(SolidFill);
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(DataLabels);
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(ChartDataGrouping.XaxisFormula!), AddStringCacheValue(ChartDataGrouping.XaxisCells!))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(ChartDataGrouping.YaxisFormula!), AddNumberCacheValue(ChartDataGrouping.YaxisCells!, null))));
            series.Append(new C.Smooth()
            {
                Val = false
            });
            return series;
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.BarChart BarChart = new(
                new C.BarDirection { Val = C.BarDirectionValues.Bar },
                new C.BarGrouping
                {
                    Val = BarChartSetting.BarChartTypes switch
                    {
                        BarChartTypes.STACKED => C.BarGroupingValues.Stacked,
                        BarChartTypes.PERCENT_STACKED => C.BarGroupingValues.PercentStacked,
                        // Clusted
                        _ => C.BarGroupingValues.Clustered
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, BarChartSetting.ChartDataSetting)
            .ForEach(Series =>
            {
                BarChart.Append(CreateBarChartSeries(seriesIndex, Series,
                    CreateSolidFill(BarChartSetting.BarChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels(BarChartSetting, seriesIndex)));
                seriesIndex++;
            });
            if (BarChartSetting.BarChartTypes == BarChartTypes.CLUSTERED)
            {
                BarChart.Append(new C.GapWidth { Val = 219 });
                BarChart.Append(new C.Overlap { Val = -27 });
            }
            else
            {
                BarChart.Append(new C.GapWidth { Val = 150 });
                BarChart.Append(new C.Overlap { Val = 100 });
            }
            BarChart.Append(new C.AxisId { Val = 1362418656 });
            BarChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(BarChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                Id = 1362418656
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels CreateDataLabel(BarChartDataLabel BarChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = BarChartDataLabel.ShowLegendKey },
                new C.ShowValue { Val = BarChartDataLabel.ShowValue },
                new C.ShowCategoryName { Val = BarChartDataLabel.ShowCategoryName },
                new C.ShowSeriesName { Val = BarChartDataLabel.ShowSeriesName },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            DataLabels.InsertAt(new C.DataLabelPosition()
            {
                Val = BarChartDataLabel.DataLabelPosition switch
                {
                    BarChartDataLabel.eDataLabelPosition.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                    BarChartDataLabel.eDataLabelPosition.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                    BarChartDataLabel.eDataLabelPosition.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
                    _ => C.DataLabelPositionValues.Center
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
            return DataLabels;
        }

        private C.DataLabels GetDataLabels(BarChartSetting BarChartSetting, int index)
        {
            if (index < BarChartSetting.BarChartSeriesSettings.Count)
            {
                return CreateDataLabel(BarChartSetting.BarChartSeriesSettings?[index]?.BarChartDataLabel ?? new BarChartDataLabel());
            }
            return CreateDataLabel(new BarChartDataLabel());
        }

        #endregion Private Methods
    }
}