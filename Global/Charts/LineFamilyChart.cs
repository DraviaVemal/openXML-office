// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C16 = DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a line chart.
    /// </summary>
    public class LineFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// The settings for the line chart.
        /// </summary>
        protected LineChartSetting lineChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Line Chart with provided settings
        /// </summary>
        /// <param name="LineChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        protected LineFamilyChart(LineChartSetting LineChartSetting, ChartData[][] DataCols) : base(LineChartSetting)
        {
            lineChartSetting = LineChartSetting;
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
                    Val = lineChartSetting.lineChartTypes switch
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
            CreateDataSeries(DataCols, lineChartSetting.chartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < lineChartSetting.lineChartSeriesSettings.Count)
                    {
                        return CreateLineDataLabels(lineChartSetting.lineChartSeriesSettings?[seriesIndex]?.lineChartDataLabel ?? new LineChartDataLabel(), Series.dataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                C.Marker Marker = new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(lineChartSetting.lineChartTypes) ? new(
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
                LineChart.Append(CreateLineChartSeries(seriesIndex, Series, Marker,
                     CreateSolidFill(lineChartSetting.lineChartSeriesSettings
                            .Where(item => item?.fillColor != null)
                            .Select(item => item?.fillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels()));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateLineDataLabels(lineChartSetting.lineChartDataLabel);
            if (DataLabels != null)
            {
                LineChart.Append(DataLabels);
            }
            LineChart.Append(new C.AxisId { Val = 1362418656 });
            LineChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(LineChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                id = 1362418656,
                crossAxisId = 1358349936,
                fontSize = lineChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = lineChartSetting.chartAxesOptions.isHorizontalBold,
                isItalic = lineChartSetting.chartAxesOptions.isHorizontalItalic,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                crossAxisId = 1362418656,
                fontSize = lineChartSetting.chartAxesOptions.verticalFontSize,
                isBold = lineChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = lineChartSetting.chartAxesOptions.isVerticalItalic,
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, C.Marker Marker, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.Extension extension = new(
                new C16.UniqueID() { Val = GeneratorUtils.GenerateNewGUID() }
            )
            { Uri = GeneratorUtils.GenerateNewGUID() };
            C.LineChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.seriesHeaderFormula!, new[] { ChartDataGrouping.seriesHeaderCells! }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.Outline(SolidFill, new A.Round()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(ShapeProperties);
            series.Append(Marker);
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(CreateCategoryAxisData(ChartDataGrouping.xAxisFormula!, ChartDataGrouping.xAxisCells!));
            series.Append(CreateValueAxisData(ChartDataGrouping.yAxisFormula!, ChartDataGrouping.yAxisCells!));
            if (ChartDataGrouping.dataLabelCells != null && ChartDataGrouping.dataLabelFormula != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(ChartDataGrouping.dataLabelFormula, ChartDataGrouping.dataLabelCells.Skip(1).ToArray())
                )
                { Uri = GeneratorUtils.GenerateNewGUID() }));
            }
            return series;
        }

        private C.DataLabels? CreateLineDataLabels(LineChartDataLabel LineChartDataLabel, int? DataLabelCounter = 0)
        {
            if (LineChartDataLabel.showValue || LineChartDataLabel.showValueFromColumn || LineChartDataLabel.showCategoryName || LineChartDataLabel.showLegendKey || LineChartDataLabel.showSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(LineChartDataLabel, DataLabelCounter);
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = LineChartDataLabel.dataLabelPosition switch
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
                    FontSize = (int)LineChartDataLabel.fontSize * 100,
                    Bold = LineChartDataLabel.isBold,
                    Italic = LineChartDataLabel.isItalic,
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