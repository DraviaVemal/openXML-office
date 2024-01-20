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
        /// <param name="ColumnChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        protected ColumnFamilyChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting)
        {
            columnChartSetting = ColumnChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
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
            CreateDataSeries(DataCols, columnChartSetting.chartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (SeriesIndex < columnChartSetting.columnChartSeriesSettings.Count)
                    {
                        return CreateColumnDataLabels(columnChartSetting.columnChartSeriesSettings[SeriesIndex]?.columnChartDataLabel ?? new ColumnChartDataLabel(), Series.dataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                ColumnChart.Append(CreateColumnChartSeries(SeriesIndex, Series,
                                    CreateSolidFill(columnChartSetting.columnChartSeriesSettings
                                            .Where(item => item?.fillColor != null)
                                            .Select(item => item?.fillColor!)
                                            .ToList(), SeriesIndex),
                                    GetDataLabels()));
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
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.BarChartSeries CreateColumnChartSeries(int SeriesIndex, ChartDataGrouping ChartDataGrouping, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)SeriesIndex) },
                new C.Order { Val = new UInt32Value((uint)SeriesIndex) },
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

        private C.DataLabels? CreateColumnDataLabels(ColumnChartDataLabel ColumnChartDataLabel, int? DataLabelCounter = 0)
        {
            if (ColumnChartDataLabel.showValue || ColumnChartDataLabel.showValueFromColumn || ColumnChartDataLabel.showCategoryName || ColumnChartDataLabel.showLegendKey || ColumnChartDataLabel.showSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(ColumnChartDataLabel, DataLabelCounter);
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = ColumnChartDataLabel.dataLabelPosition switch
                    {
                        ColumnChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                        ColumnChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                        ColumnChartDataLabel.DataLabelPositionValues.INSIDE_BASE => C.DataLabelPositionValues.InsideBase,
                        _ => C.DataLabelPositionValues.Center
                    }
                }, 0);
                DataLabels.InsertAt(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()), 0);
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = (int)ColumnChartDataLabel.fontSize * 100,
                    Bold = ColumnChartDataLabel.isBold,
                    Italic = ColumnChartDataLabel.isItalic,
                    Underline = A.TextUnderlineValues.None,
                    Strike = A.TextStrikeValues.NoStrike,
                    Kerning = 1200,
                    Baseline = 0
                }), new A.EndParagraphRunProperties() { Language = "en-US" });
                DataLabels.InsertAt(
                    new C.TextProperties(
                        new A.BodyProperties(
                            new A.ShapeAutoFit()
                            )
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
                        },
                        new A.ListStyle(),
                        Paragraph), 0);
                return DataLabels;
            }
            return null;
        }

        #endregion Private Methods
    }
}