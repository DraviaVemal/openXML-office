// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
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
        /// <param name="areaChartSetting">
        /// </param>
        /// <param name="dataCols">
        /// </param>
        protected AreaFamilyChart(AreaChartSetting areaChartSetting, ChartData[][] dataCols) : base(areaChartSetting)
        {
            this.areaChartSetting = areaChartSetting;
            SetChartPlotArea(CreateChartPlotArea(dataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping)
        {
            SolidFillModel GetFillSolidFill()
            {
                SolidFillModel solidFillModel = new();
                string? hexColor = areaChartSetting.areaChartSeriesSettings?
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
                string? hexColor = areaChartSetting.areaChartSeriesSettings?
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
            ShapePropertiesModel shapePropertiesModel = new()
            {
                solidFill = GetFillSolidFill(),
                outline = new()
                {
                    solidFill = GetOutlineSolidFill()
                }
            };
            C.DataLabels? dataLabels = seriesIndex < areaChartSetting.areaChartSeriesSettings.Count ?
                CreateAreaDataLabels(areaChartSetting.areaChartSeriesSettings?[seriesIndex]?.areaChartDataLabel ?? new AreaChartDataLabel(), ChartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            C.AreaChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.seriesHeaderFormula!, new[] { ChartDataGrouping.seriesHeaderCells! }));
            series.Append(CreateChartShapeProperties(shapePropertiesModel));
            if (dataLabels != null)
            {
                series.Append(dataLabels);
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

        private C.DataLabels? CreateAreaDataLabels(AreaChartDataLabel areaChartDataLabel, int? dataLabelCounter = 0)
        {
            if (areaChartDataLabel.showValue || areaChartDataLabel.showValueFromColumn || areaChartDataLabel.showCategoryName || areaChartDataLabel.showLegendKey || areaChartDataLabel.showSeriesName || dataLabelCounter > 0)
            {
                C.DataLabels dataLabels = CreateDataLabels(areaChartDataLabel, dataLabelCounter);
                dataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = areaChartDataLabel.dataLabelPosition switch
                    {
                        //Show
                        _ => C.DataLabelPositionValues.Center,
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
            plotArea.Append(CreateAreaChart(dataCols));
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                id = 1362418656,
                crossAxisId = 1358349936,
                fontSize = areaChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = areaChartSetting.chartAxesOptions.isHorizontalBold,
                isItalic = areaChartSetting.chartAxesOptions.isHorizontalItalic,
                isVisible = areaChartSetting.chartAxesOptions.isHorizontalAxesEnabled,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                crossAxisId = 1362418656,
                fontSize = areaChartSetting.chartAxesOptions.verticalFontSize,
                isBold = areaChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = areaChartSetting.chartAxesOptions.isVerticalItalic,
                isVisible = areaChartSetting.chartAxesOptions.isVerticalAxesEnabled,
            }));
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        private C.AreaChart CreateAreaChart(ChartData[][] dataCols)
        {
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
            CreateDataSeries(dataCols, areaChartSetting.chartDataSetting).ForEach(Series =>
            {
                AreaChart.Append(CreateAreaChartSeries(seriesIndex, Series));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateAreaDataLabels(areaChartSetting.areaChartDataLabel);
            if (DataLabels != null)
            {
                AreaChart.Append(DataLabels);
            }
            AreaChart.Append(new C.AxisId { Val = 1362418656 });
            AreaChart.Append(new C.AxisId { Val = 1358349936 });
            return AreaChart;
        }

        #endregion Private Methods
    }
}