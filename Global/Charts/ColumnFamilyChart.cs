using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;

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
            int SeriesIndex = 0;
            CreateDataSeries(DataCols, ColumnChartSetting.ChartDataSetting)
            .ForEach(Series =>
            {
                ColumnChart.Append(CreateColumnChartSeries(SeriesIndex, Series,
                                    CreateSolidFill(ColumnChartSetting.ColumnChartSeriesSettings
                                            .Where(item => item.FillColor != null)
                                            .Select(item => item.FillColor!)
                                            .ToList(), SeriesIndex),
                                    GetDataLabels(ColumnChartSetting, SeriesIndex)));
                SeriesIndex++;
            });
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

        private C.BarChartSeries CreateColumnChartSeries(int SeriesIndex, ChartDataGrouping ChartDataGrouping, A.SolidFill SolidFill, C.DataLabels DataLabels)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)SeriesIndex) },
                new C.Order { Val = new UInt32Value((uint)SeriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(SolidFill);
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(DataLabels);
            series.Append(ShapeProperties);
            series.Append(CreateCategoryAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!));
            series.Append(CreateValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!));
            if (ChartDataGrouping.DataLabelFormula != null && ChartDataGrouping.DataLabelCells != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(new C15.DataLabelsRange(new C15.Formula(ChartDataGrouping.DataLabelFormula), AddDataLabelCacheValue(ChartDataGrouping.DataLabelCells)))));
            }
            return series;
        }

        private C.DataLabels CreateDataLabel(ColumnChartDataLabel ColumnChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = ColumnChartDataLabel.ShowLegendKey },
                new C.ShowValue { Val = ColumnChartDataLabel.ShowValue },
                new C.ShowCategoryName { Val = ColumnChartDataLabel.ShowCategoryName },
                new C.ShowSeriesName { Val = ColumnChartDataLabel.ShowSeriesName },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            DataLabels.InsertAt(new C.DataLabelPosition()
            {
                Val = ColumnChartDataLabel.DataLabelPosition switch
                {
                    ColumnChartDataLabel.eDataLabelPosition.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                    ColumnChartDataLabel.eDataLabelPosition.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                    ColumnChartDataLabel.eDataLabelPosition.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
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