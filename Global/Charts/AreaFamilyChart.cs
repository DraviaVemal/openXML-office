// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Aread Chart Core data
    /// </summary>
    public class AreaFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// Area Chart Setting
        /// </summary>
        protected readonly AreaChartSetting areaChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Area Chart with provided settings
        /// </summary>
        /// <param name="AreaChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        protected AreaFamilyChart(AreaChartSetting AreaChartSetting, ChartData[][] DataCols) : base(AreaChartSetting)
        {
            areaChartSetting = AreaChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.AreaChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.seriesHeaderFormula!, new[] { ChartDataGrouping.seriesHeaderCells! }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.Outline(SolidFill, new A.Outline(new A.NoFill())));
            ShapeProperties.Append(new A.EffectList());
            series.Append(ShapeProperties);
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
                { Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
            }
            return series;
        }

        private C.DataLabels? CreateAreaDataLabels(AreaChartDataLabel AreaChartDataLabel, int? DataLabelCounter = 0)
        {
            if (AreaChartDataLabel.showValue || AreaChartDataLabel.showValueFromColumn || AreaChartDataLabel.showCategoryName || AreaChartDataLabel.showLegendKey || AreaChartDataLabel.showSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(AreaChartDataLabel, DataLabelCounter);
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = AreaChartDataLabel.dataLabelPosition switch
                    {
                        //Show
                        _ => C.DataLabelPositionValues.Center,
                    }
                }, 0);
                DataLabels.Append(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()));
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = (int)AreaChartDataLabel.fontSize * 100,
                    Bold = AreaChartDataLabel.isBold,
                    Italic = AreaChartDataLabel.isItalic,
                    Underline = A.TextUnderlineValues.None,
                    Strike = A.TextStrikeValues.NoStrike,
                    Kerning = 1200,
                    Baseline = 0
                }), new A.EndParagraphRunProperties() { Language = "en-US" });
                DataLabels.Append(new C.TextProperties(new A.BodyProperties(new A.ShapeAutoFit())
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
               Paragraph));
                return DataLabels;
            }
            return null;
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.AreaChart AreaChart = new(
                new C.Grouping
                {
                    Val = areaChartSetting.areaChartTypes switch
                    {
                        AreaChartTypes.STACKED => C.GroupingValues.Stacked,
                        AreaChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
                        // Clusted
                        _ => C.GroupingValues.Standard,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, areaChartSetting.chartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < areaChartSetting.areaChartSeriesSettings.Count)
                    {
                        return CreateAreaDataLabels(areaChartSetting.areaChartSeriesSettings?[seriesIndex]?.areaChartDataLabel ?? new AreaChartDataLabel(), Series.dataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                AreaChart.Append(CreateAreaChartSeries(seriesIndex, Series,
                                CreateSolidFill(areaChartSetting.areaChartSeriesSettings
                                        .Where(item => item?.fillColor != null)
                                        .Select(item => item?.fillColor!)
                                        .ToList(), seriesIndex),
                                GetDataLabels()));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateAreaDataLabels(areaChartSetting.areaChartDataLabel);
            if (DataLabels != null)
            {
                AreaChart.Append(DataLabels);
            }
            AreaChart.Append(new C.AxisId { Val = 1362418656 });
            AreaChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(AreaChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                id = 1362418656,
                crossAxisId = 1358349936,
                fontSize = areaChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = areaChartSetting.chartAxesOptions.isHorizontalBold,
                isItalic = areaChartSetting.chartAxesOptions.isHorizontalItalic,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                crossAxisId = 1362418656,
                fontSize = areaChartSetting.chartAxesOptions.verticalFontSize,
                isBold = areaChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = areaChartSetting.chartAxesOptions.isVerticalItalic,
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        #endregion Private Methods
    }
}