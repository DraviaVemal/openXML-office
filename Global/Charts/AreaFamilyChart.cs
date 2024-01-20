// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global {
    /// <summary>
    /// Aread Chart Core data
    /// </summary>
    public class AreaFamilyChart : ChartBase {
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
        protected AreaFamilyChart(AreaChartSetting AreaChartSetting,ChartData[][] DataCols) : base(AreaChartSetting) {
            this.areaChartSetting = AreaChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex,ChartDataGrouping ChartDataGrouping,A.SolidFill SolidFill,C.DataLabels? DataLabels) {
            C.AreaChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.seriesHeaderFormula!,new[] { ChartDataGrouping.seriesHeaderCells! }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.Outline(SolidFill,new A.Outline(new A.NoFill())));
            ShapeProperties.Append(new A.EffectList());
            if(DataLabels != null) {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateCategoryAxisData(ChartDataGrouping.xAxisFormula!,ChartDataGrouping.xAxisCells!));
            series.Append(CreateValueAxisData(ChartDataGrouping.yAxisFormula!,ChartDataGrouping.yAxisCells!));
            if(ChartDataGrouping.dataLabelCells != null && ChartDataGrouping.dataLabelFormula != null) {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(ChartDataGrouping.dataLabelFormula,ChartDataGrouping.dataLabelCells.Skip(1).ToArray())
                ) { Uri = GeneratorUtils.GenerateNewGUID() }));
            }
            return series;
        }

        private C.DataLabels? CreateAreaDataLabels(AreaChartDataLabel AreaChartDataLabel,int? DataLabelCounter = 0) {
            if(AreaChartDataLabel.showValue || AreaChartDataLabel.showValueFromColumn || AreaChartDataLabel.showCategoryName || AreaChartDataLabel.showLegendKey || AreaChartDataLabel.showSeriesName || DataLabelCounter > 0) {
                C.DataLabels DataLabels = CreateDataLabels(AreaChartDataLabel,DataLabelCounter);
                DataLabels.InsertAt(new C.DataLabelPosition() {
                    Val = AreaChartDataLabel.dataLabelPosition switch {
                        //Show
                        _ => C.DataLabelPositionValues.Center,
                    }
                },0);
                DataLabels.InsertAt(new C.ShapeProperties(new A.NoFill(),new A.Outline(new A.NoFill()),new A.EffectList()),0);
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = (int)AreaChartDataLabel.FontSize * 100,
                    Bold = AreaChartDataLabel.IsBold,
                    Italic = AreaChartDataLabel.IsItalic,
                    Underline = A.TextUnderlineValues.None,
                    Strike = A.TextStrikeValues.NoStrike,
                    Kerning = 1200,
                    Baseline = 0
                }),new A.EndParagraphRunProperties() { Language = "en-US" });
                DataLabels.InsertAt(new C.TextProperties(new A.BodyProperties(new A.ShapeAutoFit()) {
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
                },new A.ListStyle(),
               Paragraph),0);
                return DataLabels;
            }
            return null;
        }

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols) {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.AreaChart AreaChart = new(
                new C.Grouping {
                    Val = areaChartSetting.areaChartTypes switch {
                        AreaChartTypes.STACKED => C.GroupingValues.Stacked,
                        AreaChartTypes.PERCENT_STACKED => C.GroupingValues.PercentStacked,
                        // Clusted
                        _ => C.GroupingValues.Standard,
                    }
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols,areaChartSetting.chartDataSetting).ForEach(Series => {
                C.DataLabels? GetDataLabels() {
                    if(seriesIndex < areaChartSetting.areaChartSeriesSettings.Count) {
                        return CreateAreaDataLabels(areaChartSetting.areaChartSeriesSettings?[seriesIndex]?.areaChartDataLabel ?? new AreaChartDataLabel(),Series.dataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                AreaChart.Append(CreateAreaChartSeries(seriesIndex,Series,
                                CreateSolidFill(areaChartSetting.areaChartSeriesSettings
                                        .Where(item => item?.fillColor != null)
                                        .Select(item => item?.fillColor!)
                                        .ToList(),seriesIndex),
                                GetDataLabels()));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateAreaDataLabels(areaChartSetting.areaChartDataLabel);
            if(DataLabels != null) {
                AreaChart.Append(DataLabels);
            }
            AreaChart.Append(new C.AxisId { Val = 1362418656 });
            AreaChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(AreaChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                Id = 1362418656,
                CrossAxisId = 1358349936,
                FontSize = AreaChartSetting.ChartAxesOptions.HorizontalFontSize,
                IsBold = AreaChartSetting.ChartAxesOptions.IsHorizontalBold,
                IsItalic = AreaChartSetting.ChartAxesOptions.IsHorizontalItalic,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936,
                CrossAxisId = 1362418656,
                FontSize = AreaChartSetting.ChartAxesOptions.VerticalFontSize,
                IsBold = AreaChartSetting.ChartAxesOptions.IsVerticalBold,
                IsItalic = AreaChartSetting.ChartAxesOptions.IsVerticalItalic,
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