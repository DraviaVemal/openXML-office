/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

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

        private C.BarChartSeries CreateBarChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, BarChartSeriesSetting BarChartSeriesSetting, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(SolidFill);
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateCategoryAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!, BarChartSeriesSetting));
            series.Append(CreateValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!, BarChartSeriesSetting));
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
                BarChart.Append(CreateBarChartSeries(seriesIndex, Series, BarChartSetting.BarChartSeriesSettings.Count > seriesIndex ? BarChartSetting.BarChartSeriesSettings[seriesIndex] : new BarChartSeriesSetting(),
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
            C.DataLabels? DataLabels = CreateBarDataLabel(BarChartSetting.BarChartDataLabel);
            if (DataLabels != null)
            {
                BarChart.Append(DataLabels);
            }
            BarChart.Append(new C.AxisId { Val = 1362418656 });
            BarChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(BarChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                Id = 1362418656,
                AxisPosition = AxisPosition.LEFT,
                CrossAxisId = 1358349936
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936,
                AxisPosition = AxisPosition.BOTTOM,
                CrossAxisId = 1362418656
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels? CreateBarDataLabel(BarChartDataLabel BarChartDataLabel)
        {
            if (BarChartDataLabel.GetType().GetProperties()
                .Where(Prop => Prop.PropertyType == typeof(bool))
                .Any(Prop => (bool)Prop.GetValue(BarChartDataLabel)!))
            {
                C.DataLabels DataLabels = CreateDataLabel(BarChartDataLabel);
                if (BarChartSetting.BarChartTypes != BarChartTypes.CLUSTERED && BarChartDataLabel.DataLabelPosition == BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END)
                {
                    throw new ArgumentException("'Outside End' Data Label Is only Available with Cluster chart type");
                }
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = BarChartDataLabel.DataLabelPosition switch
                    {
                        BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                        BarChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                        BarChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
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
            return null;
        }

        private C.DataLabels? GetDataLabels(BarChartSetting BarChartSetting, int index)
        {
            if (index < BarChartSetting.BarChartSeriesSettings.Count)
            {
                return CreateBarDataLabel(BarChartSetting.BarChartSeriesSettings?[index]?.BarChartDataLabel ?? new BarChartDataLabel());
            }
            return null;
        }

        #endregion Private Methods
    }
}