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
        /// <param name="barChartSetting">
        /// </param>
        /// <param name="dataCols">
        /// </param>
        protected BarFamilyChart(BarChartSetting barChartSetting, ChartData[][] dataCols) : base(barChartSetting)
        {
            this.barChartSetting = barChartSetting;
            SetChartPlotArea(CreateChartPlotArea(dataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.BarChartSeries CreateBarChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
        {
            SolidFillModel GetSolidFill()
            {
                SolidFillModel solidFillModel = new();
                string? hexColor = barChartSetting.barChartSeriesSettings?
                            .Where(item => item?.fillColor != null)
                            .Select(item => item?.fillColor!)
                            .ToList().ElementAtOrDefault(seriesIndex);
                if (hexColor != null)
                {
                    solidFillModel.hexColor = hexColor;
                    return solidFillModel;
                }
                else
                {
                    solidFillModel.schemeColorModel = new()
                    {
                        themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % 6),
                    };
                }
                return solidFillModel;
            }
            C.DataLabels? dataLabels = seriesIndex < barChartSetting.barChartSeriesSettings.Count ? CreateBarDataLabels(barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataLabel ?? new BarChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            ShapePropertiesModel shapePropertiesModel = new()
            {
                solidFill = GetSolidFill(),
                outline = new()
                {
                    solidFill = GetSolidFill()
                }
            };
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }),
                new C.InvertIfNegative { Val = true });
            series.Append(CreateChartShapeProperties(shapePropertiesModel));
            if (dataLabels != null)
            {
                series.Append(dataLabels);
            }
            series.Append(CreateCategoryAxisData(chartDataGrouping.xAxisFormula!, chartDataGrouping.xAxisCells!));
            series.Append(CreateValueAxisData(chartDataGrouping.yAxisFormula!, chartDataGrouping.yAxisCells!));
            if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
                )
                { Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
            }
            return series;
        }

        private C.DataLabels? CreateBarDataLabels(BarChartDataLabel barChartDataLabel, int? dataLabelCounter = 0)
        {
            if (barChartDataLabel.showValue || barChartDataLabel.showValueFromColumn || barChartDataLabel.showCategoryName || barChartDataLabel.showLegendKey || barChartDataLabel.showSeriesName || dataLabelCounter > 0)
            {
                C.DataLabels dataLabels = CreateDataLabels(barChartDataLabel, dataLabelCounter);
                if (barChartSetting.barChartTypes != BarChartTypes.CLUSTERED && barChartDataLabel.dataLabelPosition == BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END)
                {
                    throw new ArgumentException("'Outside End' Data Label Is only Available with Cluster chart type");
                }
                dataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = barChartDataLabel.dataLabelPosition switch
                    {
                        BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                        BarChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                        BarChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
                        _ => C.DataLabelPositionValues.Center
                    }
                }, 0);
                dataLabels.Append(CreateChartShapeProperties());
                dataLabels.Append(CreateChartTextProperties(new()
                {
                    bodyProperties = new()
                    {
                        rotation = 0,
                        anchorCenter = true,
                        anchor = TextAnchoringValues.CENTER,
                        bottomInset = 19050,
                        leftInset = 38100,
                        rightInset = 38100,
                        topInset = 19050,
                        useParagraphSpacing = true,
                        vertical = TextVerticalAlignmentValues.HORIZONTAL,
                        verticalOverflow = TextVerticalOverflowValues.ELLIPSIS,
                        wrap = TextWrappingValues.SQUARE,
                    },
                    drawingParagraph = new()
                    {
                        paragraphPropertiesModel = new()
                        {
                            defaultRunProperties = new()
                            {
                                solidFill = new()
                                {
                                    schemeColorModel = new()
                                    {
                                        themeColorValues = ThemeColorValues.TEXT_1,
                                        luminanceModulation = 7500,
                                        luminanceOffset = 2500
                                    }
                                },
                                complexScriptFont = "+mn-cs",
                                eastAsianFont = "+mn-ea",
                                latinFont = "+mn-lt",
                                fontSize = (int)barChartDataLabel.fontSize * 100,
                                bold = barChartDataLabel.isBold,
                                italic = barChartDataLabel.isItalic,
                                underline = UnderLineValues.NONE,
                                strike = StrikeValues.NO_STRIKE,
                                kerning = 1200,
                                baseline = 0,
                            }
                        }
                    }
                }));
                return dataLabels;
            }
            return null;
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
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
            CreateDataSeries(dataCols, barChartSetting.chartDataSetting).ForEach(Series =>
            {
                BarChart.Append(CreateBarChartSeries(seriesIndex, Series));
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
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        #endregion Private Methods
    }
}