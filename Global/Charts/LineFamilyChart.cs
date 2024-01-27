// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
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
        /// <param name="lineChartSetting">
        /// </param>
        /// <param name="dataCols">
        /// </param>
        protected LineFamilyChart(LineChartSetting lineChartSetting, ChartData[][] dataCols) : base(lineChartSetting)
        {
            this.lineChartSetting = lineChartSetting;
            SetChartPlotArea(CreateChartPlotArea(dataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
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
            CreateDataSeries(dataCols, lineChartSetting.chartDataSetting).ForEach(Series =>
            {
                LineChart.Append(CreateLineChartSeries(seriesIndex, Series));
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
                isVisible = lineChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                crossAxisId = 1362418656,
                fontSize = lineChartSetting.chartAxesOptions.verticalFontSize,
                isBold = lineChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = lineChartSetting.chartAxesOptions.isVerticalItalic,
                isVisible = lineChartSetting.chartAxesOptions.isVerticalAxesEnabled,
            }));
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        private C.LineChartSeries CreateLineChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
        {
            MarkerModel marketModel = new()
            {
                markerShapeValues = MarkerModel.MarkerShapeValues.NONE,
            };
            if (new[] { LineChartTypes.CLUSTERED_MARKER, LineChartTypes.STACKED_MARKER, LineChartTypes.PERCENT_STACKED_MARKER }.Contains(lineChartSetting.lineChartTypes))
            {
                marketModel.markerShapeValues = MarkerModel.MarkerShapeValues.CIRCLE;
                marketModel.shapeProperties = new()
                {
                    solidFill = new()
                    {
                        schemeColorModel = new()
                        {
                            themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
                        }
                    },
                    outline = new()
                    {
                        solidFill = new()
                        {
                            schemeColorModel = new()
                            {
                                themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % AccentColurCount),
                            }
                        }
                    }
                };
            }
            C.DataLabels? dataLabels = seriesIndex < lineChartSetting.lineChartSeriesSettings.Count ?
                CreateLineDataLabels(lineChartSetting.lineChartSeriesSettings?[seriesIndex]?.lineChartDataLabel ?? new LineChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            SolidFillModel GetSolidFill()
            {
                SolidFillModel solidFillModel = new();
                string? hexColor = lineChartSetting.lineChartSeriesSettings?
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
            C.LineChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
            ShapePropertiesModel shapePropertiesModel = new()
            {
                outline = new()
                {
                    solidFill = GetSolidFill()
                }
            };
            series.Append(CreateChartShapeProperties(shapePropertiesModel));
            series.Append(CreateMarker(marketModel));
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

        private C.DataLabels? CreateLineDataLabels(LineChartDataLabel lineChartDataLabel, int? dataLabelCounter = 0)
        {
            if (lineChartDataLabel.showValue || lineChartDataLabel.showValueFromColumn || lineChartDataLabel.showCategoryName || lineChartDataLabel.showLegendKey || lineChartDataLabel.showSeriesName || dataLabelCounter > 0)
            {
                C.DataLabels dataLabels = CreateDataLabels(lineChartDataLabel, dataLabelCounter);
                dataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = lineChartDataLabel.dataLabelPosition switch
                    {
                        LineChartDataLabel.DataLabelPositionValues.LEFT => C.DataLabelPositionValues.Left,
                        LineChartDataLabel.DataLabelPositionValues.RIGHT => C.DataLabelPositionValues.Right,
                        LineChartDataLabel.DataLabelPositionValues.ABOVE => C.DataLabelPositionValues.Top,
                        LineChartDataLabel.DataLabelPositionValues.BELOW => C.DataLabelPositionValues.Bottom,
                        //Center
                        _ => C.DataLabelPositionValues.Center,
                    }
                }, 0);
                return dataLabels;
            }
            return null;
        }

        #endregion Private Methods
    }
}