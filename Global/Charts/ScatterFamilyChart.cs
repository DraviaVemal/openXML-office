// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the types of scatter charts.
    /// </summary>
    public class ScatterFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// Scatter Chart Setting
        /// </summary>
        protected ScatterChartSetting scatterChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Scatter Chart with provided settings
        /// </summary>
        /// <param name="scatterChartSetting">
        /// </param>
        /// <param name="dataCols">
        /// </param>
        protected ScatterFamilyChart(ScatterChartSetting scatterChartSetting, ChartData[][] dataCols) : base(scatterChartSetting)
        {
            this.scatterChartSetting = scatterChartSetting;
            SetChartPlotArea(CreateChartPlotArea(dataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] dataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            OpenXmlCompositeElement Chart = scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE ? new C.BubbleChart() : new C.ScatterChart(
                new C.ScatterStyle
                {
                    Val = scatterChartSetting.scatterChartTypes switch
                    {
                        ScatterChartTypes.SCATTER_SMOOTH => C.ScatterStyleValues.Smooth,
                        ScatterChartTypes.SCATTER_SMOOTH_MARKER => C.ScatterStyleValues.SmoothMarker,
                        ScatterChartTypes.SCATTER_STRIGHT => C.ScatterStyleValues.Line,
                        ScatterChartTypes.SCATTER_STRIGHT_MARKER => C.ScatterStyleValues.LineMarker,
                        // Clusted
                        _ => C.ScatterStyleValues.LineMarker,
                    }
                });
            Chart.Append(new C.VaryColors() { Val = false });
            if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                scatterChartSetting.chartDataSetting.is3Ddata = true;
                if ((dataCols.Length - 1) % 2 != 0)
                {
                    throw new ArgumentOutOfRangeException("Required 3D Data Size is not met.");
                }
            }
            int seriesIndex = 0;
            CreateDataSeries(dataCols, scatterChartSetting.chartDataSetting).ForEach(Series =>
            {
                Chart.Append(CreateScatterChartSeries(seriesIndex, Series));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateScatterDataLabels(scatterChartSetting.scatterChartDataLabel);
            if (DataLabels != null)
            {
                Chart.Append(DataLabels);
            }
            if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                Chart.Append(new C.BubbleScale() { Val = 100 });
                Chart.Append(new C.ShowNegativeBubbles() { Val = false });
            }
            Chart.Append(new C.AxisId { Val = 1362418656 });
            Chart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(Chart);
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1362418656,
                axisPosition = AxisPosition.BOTTOM,
                crossAxisId = 1358349936,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                id = 1358349936,
                crossAxisId = 1362418656
            }));
            plotArea.Append(CreateChartShapeProperties());
            return plotArea;
        }

        private C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ChartDataGrouping chartDataGrouping)
        {
            C.DataLabels? dataLabels = seriesIndex < scatterChartSetting.scatterChartSeriesSettings.Count ? CreateScatterDataLabels(scatterChartSetting.scatterChartSeriesSettings?[seriesIndex]?.scatterChartDataLabel ?? new ScatterChartDataLabel(), chartDataGrouping.dataLabelCells?.Length ?? 0) : null;
            SolidFillModel GetSolidFill()
            {
                SolidFillModel solidFillModel = new();
                string? hexColor = scatterChartSetting.scatterChartSeriesSettings?
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
            MarkerModel markerModel = new();
            if (new[] { ScatterChartTypes.SCATTER, ScatterChartTypes.SCATTER_SMOOTH_MARKER, ScatterChartTypes.SCATTER_STRIGHT_MARKER }.Contains(scatterChartSetting.scatterChartTypes))
            {
                markerModel.markerShapeValues = scatterChartSetting.scatterChartTypes == ScatterChartTypes.SCATTER ? MarkerModel.MarkerShapeValues.AUTO : MarkerModel.MarkerShapeValues.CIRCLE;
                markerModel.shapeProperties = new()
                {
                    solidFill = new()
                    {
                        schemeColorModel = new()
                        {
                            themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % 6),
                        }
                    },
                    outline = new()
                    {
                        solidFill = new()
                        {
                            schemeColorModel = new()
                            {
                                themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % 6),
                            }
                        }
                    }
                };
            }
            C.ScatterChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(chartDataGrouping.seriesHeaderFormula!, new[] { chartDataGrouping.seriesHeaderCells! }));
            ShapePropertiesModel shapePropertiesModel = new()
            {
                outline = new()
                {
                    solidFill = scatterChartSetting.scatterChartTypes == ScatterChartTypes.SCATTER ? null : GetSolidFill(),
                }
            };
            if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                shapePropertiesModel.solidFill = new()
                {
                    schemeColorModel = new()
                    {
                        themeColorValues = ThemeColorValues.ACCENT_1 + (seriesIndex % 6),
                        tint = 75000,

                    }
                };
                series.Append(new C.InvertIfNegative() { Val = false });
            }
            series.Append(CreateChartShapeProperties(shapePropertiesModel));
            if (scatterChartSetting.scatterChartTypes != ScatterChartTypes.BUBBLE)
            {
                series.Append(CreateMarker(markerModel));
            }
            if (dataLabels != null)
            {
                series.Append(dataLabels);
            }
            series.Append(CreateXValueAxisData(chartDataGrouping.xAxisFormula!, chartDataGrouping.xAxisCells!));
            series.Append(CreateYValueAxisData(chartDataGrouping.yAxisFormula!, chartDataGrouping.yAxisCells!));
            if (scatterChartSetting.scatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                series.Append(CreateBubbleSizeAxisData(chartDataGrouping.zAxisFormula!, chartDataGrouping.zAxisCells!));
                series.Append(new C.Bubble3D() { Val = false });
            }
            else
            {
                series.Append(new C.Smooth() { Val = new[] { ScatterChartTypes.SCATTER_SMOOTH, ScatterChartTypes.SCATTER_SMOOTH_MARKER }.Contains(scatterChartSetting.scatterChartTypes) });
            }
            if (chartDataGrouping.dataLabelCells != null && chartDataGrouping.dataLabelFormula != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(chartDataGrouping.dataLabelFormula, chartDataGrouping.dataLabelCells.Skip(1).ToArray())
                )
                { Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" }));
            }
            return series;
        }

        private C.DataLabels? CreateScatterDataLabels(ScatterChartDataLabel scatterChartDataLabel, int? dataLabelCounter = 0)
        {
            if (scatterChartDataLabel.showValue || scatterChartDataLabel.showValueFromColumn || scatterChartDataLabel.showCategoryName || scatterChartDataLabel.showLegendKey || scatterChartDataLabel.showSeriesName || scatterChartDataLabel.showBubbleSize || dataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(scatterChartDataLabel, dataLabelCounter);
                DataLabels.Append(new C.ShowBubbleSize { Val = scatterChartDataLabel.showBubbleSize });
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = scatterChartDataLabel.dataLabelPosition switch
                    {
                        ScatterChartDataLabel.DataLabelPositionValues.LEFT => C.DataLabelPositionValues.Left,
                        ScatterChartDataLabel.DataLabelPositionValues.RIGHT => C.DataLabelPositionValues.Right,
                        ScatterChartDataLabel.DataLabelPositionValues.ABOVE => C.DataLabelPositionValues.Top,
                        ScatterChartDataLabel.DataLabelPositionValues.BELOW => C.DataLabelPositionValues.Bottom,
                        //Center
                        _ => C.DataLabelPositionValues.Center,
                    }
                }, 0);
                DataLabels.Append(CreateChartShapeProperties());
                A.Paragraph Paragraph = new(new A.ParagraphProperties(CreateDefaultRunProperties(new()
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
                    fontSize = (int)scatterChartDataLabel.fontSize * 100,
                    bold = scatterChartDataLabel.isBold,
                    italic = scatterChartDataLabel.isItalic,
                    underline = UnderLineValues.NONE,
                    strike = StrikeValues.NO_STRIKE,
                    kerning = 1200,
                    baseline = 0,
                })), new A.EndParagraphRunProperties() { Language = "en-US" });
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

        #endregion Private Methods
    }
}