/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class LineFamilyChart : ChartBase
    {
        #region Protected Fields

        protected LineChartSetting LineChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        protected LineFamilyChart(LineChartSetting LineChartSetting, ChartData[][] DataCols) : base(LineChartSetting)
        {
            this.LineChartSetting = LineChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.LineChart LineChart = new(
                new C.Grouping
                {
                    Val = LineChartSetting.LineChartTypes switch
                    {
                        LineChartTypes.STACKED => C.GroupingValues.Stacked,
                        LineChartTypes.STACKED_MARKER => C.GroupingValues.Stacked,
                        LineChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
                        LineChartTypes.PERCENT_STACKED_MARKER => C.GroupingValues.PercentStacked,
                        // Clusted
                        _ => C.GroupingValues.Standard,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, LineChartSetting.ChartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < LineChartSetting.LineChartSeriesSettings.Count)
                    {
                        return CreateLineDataLabels(LineChartSetting.LineChartSeriesSettings?[seriesIndex]?.LineChartDataLabel ?? new LineChartDataLabel(), Series.DataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                C.Marker Marker = new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(LineChartSetting.LineChartTypes) ? new(
                    new C.Symbol { Val = C.MarkerStyleValues.Circle },
                    new C.Size { Val = 5 },
                    new C.ShapeProperties(
                        CreateSolidFill(new List<string>(), seriesIndex),
                        new A.Outline(CreateSolidFill(new List<string>(), seriesIndex)),
                        new A.EffectList()
                    )) :
                    new(new C.Symbol()
                    {
                        Val = C.MarkerStyleValues.None
                    });
                LineChart.Append(CreateLineChartSeries(seriesIndex, Series, LineChartSetting.LineChartSeriesSettings.Count > seriesIndex ? LineChartSetting.LineChartSeriesSettings[seriesIndex] : new LineChartSeriesSetting(), Marker,
                     CreateSolidFill(LineChartSetting.LineChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels()));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateLineDataLabels(LineChartSetting.LineChartDataLabel);
            if (DataLabels != null)
            {
                LineChart.Append(DataLabels);
            }
            LineChart.Append(new C.AxisId { Val = 1362418656 });
            LineChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(LineChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                Id = 1362418656,
                CrossAxisId = 1358349936,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936,
                CrossAxisId = 1362418656
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, LineChartSeriesSetting LineChartSeriesSetting, C.Marker Marker, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.LineChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }),
                Marker);
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.Outline(SolidFill, new A.Round()));
            ShapeProperties.Append(new A.EffectList());
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateCategoryAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!, LineChartSeriesSetting));
            series.Append(CreateValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!, LineChartSeriesSetting));
            if (ChartDataGrouping.DataLabelCells != null && ChartDataGrouping.DataLabelFormula != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(ChartDataGrouping.DataLabelFormula, ChartDataGrouping.DataLabelCells.Skip(1).ToArray(), LineChartSeriesSetting)
                )
                { Uri = GeneratorUtils.GenerateNewGUID() }));
            }
            return series;
        }

        private C.DataLabels? CreateLineDataLabels(LineChartDataLabel LineChartDataLabel, int? DataLabelCounter = 0)
        {
            if (LineChartDataLabel.ShowValue || LineChartDataLabel.ShowCategoryName || LineChartDataLabel.ShowLegendKey || LineChartDataLabel.ShowSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(LineChartDataLabel, DataLabelCounter);
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = LineChartDataLabel.DataLabelPosition switch
                    {
                        LineChartDataLabel.DataLabelPositionValues.LEFT => C.DataLabelPositionValues.Left,
                        LineChartDataLabel.DataLabelPositionValues.RIGHT => C.DataLabelPositionValues.Right,
                        LineChartDataLabel.DataLabelPositionValues.ABOVE => C.DataLabelPositionValues.Top,
                        LineChartDataLabel.DataLabelPositionValues.BELOW => C.DataLabelPositionValues.Bottom,
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
                return DataLabels;
            }
            return null;
        }

        #endregion Private Methods
    }
}