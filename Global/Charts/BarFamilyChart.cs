// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a bar chart.
    /// </summary>
    public class BarFamilyChart : ChartBase
    {
        private const int DefaultGapWidth = 150;
        private const int DefaultOverlap = 100;
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
            SolidFillModel GetFillSolidFill()
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
                        themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
                    };
                }
                return solidFillModel;
            }
            SolidFillModel GetOutlineSolidFill()
            {
                SolidFillModel solidFillModel = new();
                string? hexColor = barChartSetting.barChartSeriesSettings?
                            .Where(item => item?.borderColor != null)
                            .Select(item => item?.borderColor!)
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
                        themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
                    };
                }
                return solidFillModel;
            }
            C.DataLabels? dataLabels = seriesIndex < barChartSetting.barChartSeriesSettings.Count ?
                CreateBarDataLabels(barChartSetting.barChartSeriesSettings?[seriesIndex]?.barChartDataLabel ?? new BarChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            ShapePropertiesModel shapePropertiesModel = new()
            {
                solidFill = GetFillSolidFill(),
                outline = new()
                {
                    solidFill = GetOutlineSolidFill()
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
                BarChart.Append(new C.GapWidth { Val = DefaultGapWidth });
                BarChart.Append(new C.Overlap { Val = DefaultOverlap });
            }
            C.DataLabels? DataLabels = CreateBarDataLabels(barChartSetting.barChartDataLabel);
            if (DataLabels != null)
            {
                BarChart.Append(DataLabels);
            }
            BarChart.Append(new C.AxisId { Val = CategoryAxisId });
            BarChart.Append(new C.AxisId { Val = ValueAxisId });
            plotArea.Append(BarChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                id = CategoryAxisId,
                crossAxisId = ValueAxisId,
                axisPosition = AxisPosition.LEFT,
                fontSize = barChartSetting.chartAxesOptions.verticalFontSize,
                isBold = barChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = barChartSetting.chartAxesOptions.isVerticalItalic,
                isVisible = barChartSetting.chartAxesOptions.isVerticalAxesEnabled,
                invertOrder = barChartSetting.chartAxesOptions.invertVerticalAxesOrder,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = ValueAxisId,
                crossAxisId = CategoryAxisId,
                axisPosition = AxisPosition.BOTTOM,
                fontSize = barChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = barChartSetting.chartAxesOptions.isHorizontalBold,
                isItalic = barChartSetting.chartAxesOptions.isHorizontalItalic,
                isVisible = barChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
                invertOrder = barChartSetting.chartAxesOptions.invertHorizontalAxesOrder,
            }));
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        #endregion Private Methods
    }
}