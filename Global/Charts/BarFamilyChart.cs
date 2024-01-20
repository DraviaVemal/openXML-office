// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a bar chart.
    /// </summary>
    public class BarFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// Bar Chart Setting
        /// </summary>
        protected readonly BarChartSetting barChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Bar Chart with provided settings
        /// </summary>
        /// <param name="BarChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        protected BarFamilyChart(BarChartSetting BarChartSetting, ChartData[][] DataCols) : base(BarChartSetting)
        {
            barChartSetting = BarChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.BarChartSeries CreateBarChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.seriesHeaderFormula!, new[] { ChartDataGrouping.seriesHeaderCells! }),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(SolidFill);
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
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
                { Uri = GeneratorUtils.GenerateNewGUID() }));
            }
            return series;
        }

        private C.DataLabels? CreateBarDataLabels(BarChartDataLabel BarChartDataLabel, int? DataLabelCounter = 0)
        {
            if (BarChartDataLabel.showValue || BarChartDataLabel.showValueFromColumn || BarChartDataLabel.showCategoryName || BarChartDataLabel.showLegendKey || BarChartDataLabel.showSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(BarChartDataLabel, DataLabelCounter);
                if (barChartSetting.barChartTypes != BarChartTypes.CLUSTERED && BarChartDataLabel.dataLabelPosition == BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END)
                {
                    throw new ArgumentException("'Outside End' Data Label Is only Available with Cluster chart type");
                }
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = BarChartDataLabel.dataLabelPosition switch
                    {
                        BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                        BarChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                        BarChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
                        _ => C.DataLabelPositionValues.Center
                    }
                }, 0);
                DataLabels.Append(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()));
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = (int)BarChartDataLabel.fontSize * 100,
                    Bold = BarChartDataLabel.isBold,
                    Italic = BarChartDataLabel.isItalic,
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
            C.BarChart BarChart = new(
                new C.BarDirection { Val = C.BarDirectionValues.Bar },
                new C.BarGrouping
                {
                    Val = barChartSetting.barChartTypes switch
                    {
                        BarChartTypes.STACKED => C.BarGroupingValues.Stacked,
                        BarChartTypes.PERCENT_STACKED => C.BarGroupingValues.PercentStacked,
                        // Clusted
                        _ => C.BarGroupingValues.Clustered
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, barChartSetting.chartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < barChartSetting.barChartSeriesSettings.Count)
                    {
                        return CreateBarDataLabels(barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataLabel ?? new BarChartDataLabel(), Series.dataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                BarChart.Append(CreateBarChartSeries(seriesIndex, Series,
                    CreateSolidFill(barChartSetting.barChartSeriesSettings
                            .Where(item => item?.fillColor != null)
                            .Select(item => item?.fillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels()));
                seriesIndex++;
            });
            if (barChartSetting.barChartTypes == BarChartTypes.CLUSTERED)
            {
                BarChart.Append(new C.GapWidth { Val = (UInt16Value)barChartSetting.barGraphicsSetting.categoryGap });
                BarChart.Append(new C.Overlap { Val = (SByteValue)barChartSetting.barGraphicsSetting.seriesGap });
            }
            else
            {
                BarChart.Append(new C.GapWidth { Val = 150 });
                BarChart.Append(new C.Overlap { Val = 100 });
            }
            C.DataLabels? DataLabels = CreateBarDataLabels(barChartSetting.barChartDataLabel);
            if (DataLabels != null)
            {
                BarChart.Append(DataLabels);
            }
            BarChart.Append(new C.AxisId { Val = 1362418656 });
            BarChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(BarChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                id = 1362418656,
                axisPosition = AxisPosition.LEFT,
                crossAxisId = 1358349936,
                fontSize = barChartSetting.chartAxesOptions.verticalFontSize,
                isBold = barChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = barChartSetting.chartAxesOptions.isVerticalItalic,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                axisPosition = AxisPosition.BOTTOM,
                crossAxisId = 1362418656,
                fontSize = barChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = barChartSetting.chartAxesOptions.isHorizontalBold,
                isItalic = barChartSetting.chartAxesOptions.isHorizontalItalic,
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