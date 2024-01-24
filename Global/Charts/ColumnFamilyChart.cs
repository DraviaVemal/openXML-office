// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a column chart.
    /// </summary>
    public class ColumnFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// Column Chart Setting
        /// </summary>
        protected ColumnChartSetting columnChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Column Chart with provided settings
        /// </summary>
        /// <param name="columnChartSetting">
        /// </param>
        /// <param name="dataCols">
        /// </param>
        protected ColumnFamilyChart(ColumnChartSetting columnChartSetting, ChartData[][] dataCols) : base(columnChartSetting)
        {
            this.columnChartSetting = columnChartSetting;
            SetChartPlotArea(CreateChartPlotArea(dataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.BarChart ColumnChart = new(
                new C.BarDirection { Val = C.BarDirectionValues.Column },
                new C.BarGrouping
                {
                    Val = columnChartSetting.columnChartTypes switch
                    {
                        ColumnChartTypes.STACKED => C.BarGroupingValues.Stacked,
                        ColumnChartTypes.PERCENT_STACKED => C.BarGroupingValues.PercentStacked,
                        // Clusted
                        _ => C.BarGroupingValues.Clustered,
                    }
                },
                new C.VaryColors { Val = false });
            int SeriesIndex = 0;
            CreateDataSeries(dataCols, columnChartSetting.chartDataSetting).ForEach(Series =>
            {
                ColumnChart.Append(CreateColumnChartSeries(SeriesIndex, Series));
                SeriesIndex++;
            });
            if (columnChartSetting.columnChartTypes == ColumnChartTypes.CLUSTERED)
            {
                ColumnChart.Append(new C.GapWidth { Val = (UInt16Value)columnChartSetting.columnGraphicsSetting.categoryGap });
                ColumnChart.Append(new C.Overlap { Val = (SByteValue)columnChartSetting.columnGraphicsSetting.seriesGap });
            }
            else
            {
                ColumnChart.Append(new C.GapWidth { Val = 150 });
                ColumnChart.Append(new C.Overlap { Val = 100 });
            }
            C.DataLabels? DataLabels = CreateColumnDataLabels(columnChartSetting.columnChartDataLabel);
            if (DataLabels != null)
            {
                ColumnChart.Append(DataLabels);
            }
            ColumnChart.Append(new C.AxisId { Val = 1362418656 });
            ColumnChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(ColumnChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                id = 1362418656,
                crossAxisId = 1358349936,
                fontSize = columnChartSetting.chartAxesOptions.horizontalFontSize,
                isBold = columnChartSetting.chartAxesOptions.isVerticalItalic,
                isItalic = columnChartSetting.chartAxesOptions.isVerticalItalic,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                crossAxisId = 1362418656,
                fontSize = columnChartSetting.chartAxesOptions.verticalFontSize,
                isBold = columnChartSetting.chartAxesOptions.isVerticalBold,
                isItalic = columnChartSetting.chartAxesOptions.isVerticalItalic,
            }));
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        private C.BarChartSeries CreateColumnChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
        {
            C.DataLabels? dataLabels = seriesIndex < columnChartSetting.columnChartSeriesSettings.Count ? CreateColumnDataLabels(columnChartSetting.columnChartSeriesSettings[seriesIndex]?.columnChartDataLabel ?? new ColumnChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            SolidFillModel GetSolidFill()
            {
                SolidFillModel solidFillModel = new();
                string? hexColor = columnChartSetting.columnChartSeriesSettings?
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

        private C.DataLabels? CreateColumnDataLabels(ColumnChartDataLabel columnChartDataLabel, int? dataLabelCounter = 0)
        {
            if (columnChartDataLabel.showValue || columnChartDataLabel.showValueFromColumn || columnChartDataLabel.showCategoryName || columnChartDataLabel.showLegendKey || columnChartDataLabel.showSeriesName || dataLabelCounter > 0)
            {
                C.DataLabels dataLabels = CreateDataLabels(columnChartDataLabel, dataLabelCounter);
                dataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = columnChartDataLabel.dataLabelPosition switch
                    {
                        ColumnChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                        ColumnChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                        ColumnChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
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
                                fontSize = (int)columnChartDataLabel.fontSize * 100,
                                bold = columnChartDataLabel.isBold,
                                italic = columnChartDataLabel.isItalic,
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

        #endregion Private Methods
    }
}