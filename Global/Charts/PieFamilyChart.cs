// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of pie charts.
    /// </summary>
    public class PieFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// The settings for the pie chart.
        /// </summary>
        protected PieChartSetting pieChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Pie Chart with provided settings
        /// </summary>
        /// <param name="pieChartSetting">
        /// </param>
        /// <param name="dataCols">
        /// </param>
        protected PieFamilyChart(PieChartSetting pieChartSetting, ChartData[][] dataCols) : base(pieChartSetting)
        {
            this.pieChartSetting = pieChartSetting;
            switch (pieChartSetting.pieChartTypes)
            {
                case PieChartTypes.DOUGHNUT:
                    SetChartPlotArea(CreateChartPlotArea(dataCols));
                    break;

                default:
                    SetChartPlotArea(CreateChartPlotArea(dataCols));
                    break;
            };
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            OpenXmlCompositeElement Chart = pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT ? new C.DoughnutChart(
                new C.VaryColors { Val = true }) : new C.PieChart(
                new C.VaryColors { Val = true });
            int seriesIndex = 0;
            CreateDataSeries(dataCols, pieChartSetting.chartDataSetting).ForEach(Series =>
            {
                Chart.Append(CreateChartSeries(seriesIndex, Series));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreatePieDataLabels(pieChartSetting.pieChartDataLabel);
            if (DataLabels != null)
            {
                Chart.Append(DataLabels);
            }
            Chart.Append(new C.FirstSliceAngle { Val = 0 });
            Chart.Append(new C.HoleSize { Val = 50 });
            plotArea.Append(Chart);
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        private C.PieChartSeries CreateChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
        {
            C.DataLabels? dataLabels = seriesIndex < pieChartSetting.pieChartSeriesSettings.Count ?
                CreatePieDataLabels(pieChartSetting.pieChartSeriesSettings?[seriesIndex]?.pieChartDataLabel ?? new PieChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            C.PieChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
            for (uint index = 0; index < chartDataGrouping.xAxisCells!.Length; index++)
            {
                C.DataPoint dataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
                ShapePropertiesModel shapePropertiesModel = new()
                {
                    solidFill = new()
                    {
                        schemeColorModel = new()
                        {
                            themeColorValues = ThemeColorValues.ACCENT_1 + ((int)index % AccentColurCount),
                        }
                    }
                };
                if (pieChartSetting.pieChartTypes != PieChartTypes.DOUGHNUT)
                {
                    shapePropertiesModel.outline = new()
                    {
                        solidFill = new()
                        {
                            schemeColorModel = new()
                            {
                                themeColorValues = ThemeColorValues.ACCENT_1 + ((int)index % AccentColurCount),
                            }

                        }
                    };
                }
                dataPoint.Append(CreateChartShapeProperties(shapePropertiesModel));
                if (dataLabels != null)
                {
                    series.Append(dataLabels);
                }
                series.Append(dataPoint);
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

        private C.DataLabels? CreatePieDataLabels(PieChartDataLabel pieChartDataLabel, int? dataLabelCounter = 0)
        {
            if (pieChartDataLabel.showValue || pieChartDataLabel.showCategoryName || pieChartDataLabel.showLegendKey || pieChartDataLabel.showSeriesName || dataLabelCounter > 0)
            {
                C.DataLabels dataLabels = CreateDataLabels(pieChartDataLabel, dataLabelCounter);
                if (pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT &&
                    new[] { PieChartDataLabel.DataLabelPositionValues.CENTER, PieChartDataLabel.DataLabelPositionValues.INSIDE_END, PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END, PieChartDataLabel.DataLabelPositionValues.BEST_FIT }.Contains(pieChartDataLabel.dataLabelPosition))
                {
                    throw new ArgumentException("DataLabelPosition is not supported for Doughnut Chart.");
                }
                if (pieChartSetting.pieChartTypes != PieChartTypes.DOUGHNUT)
                {
                    dataLabels.InsertAt(new C.DataLabelPosition()
                    {
                        Val = pieChartDataLabel.dataLabelPosition switch
                        {
                            PieChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                            PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                            PieChartDataLabel.DataLabelPositionValues.BEST_FIT => C.DataLabelPositionValues.BestFit,
                            //Center
                            _ => C.DataLabelPositionValues.Center,
                        }
                    }, 0);
                }
                return dataLabels;
            }
            return null;
        }

        #endregion Private Methods
    }
}