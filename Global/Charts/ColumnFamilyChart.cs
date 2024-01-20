/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

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
        protected ColumnChartSetting ColumnChartSetting;

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
            this.ColumnChartSetting = ColumnChartSetting;
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
                    Val = ColumnChartSetting.ColumnChartTypes switch
                    {
                        ColumnChartTypes.STACKED => C.BarGroupingValues.Stacked,
                        ColumnChartTypes.PERCENT_STACKED => C.BarGroupingValues.PercentStacked,
                        // Clusted
                        _ => C.BarGroupingValues.Clustered,
                    }
                },
                new C.VaryColors { Val = false });
            int SeriesIndex = 0;
            CreateDataSeries(DataCols, ColumnChartSetting.ChartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (SeriesIndex < ColumnChartSetting.ColumnChartSeriesSettings.Count)
                    {
                        return CreateColumnDataLabels(ColumnChartSetting.ColumnChartSeriesSettings[SeriesIndex]?.ColumnChartDataLabel ?? new ColumnChartDataLabel(), Series.DataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                ColumnChart.Append(CreateColumnChartSeries(SeriesIndex, Series,
                                    CreateSolidFill(ColumnChartSetting.ColumnChartSeriesSettings
                                            .Where(item => item?.FillColor != null)
                                            .Select(item => item?.FillColor!)
                                            .ToList(), SeriesIndex),
                                    GetDataLabels()));
                SeriesIndex++;
            });
            if (ColumnChartSetting.ColumnChartTypes == ColumnChartTypes.CLUSTERED)
            {
                ColumnChart.Append(new C.GapWidth { Val = 219 });
                ColumnChart.Append(new C.Overlap { Val = -27 });
            }
            else
            {
                ColumnChart.Append(new C.GapWidth { Val = 150 });
                ColumnChart.Append(new C.Overlap { Val = 100 });
            }
            C.DataLabels? DataLabels = CreateColumnDataLabels(ColumnChartSetting.ColumnChartDataLabel);
            if (DataLabels != null)
            {
                ColumnChart.Append(DataLabels);
            }
            ColumnChart.Append(new C.AxisId { Val = 1362418656 });
            ColumnChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(ColumnChart);
            plotArea.Append(CreateCategoryAxis(new CategoryAxisSetting()
            {
                Id = 1362418656,
                CrossAxisId = 1358349936,
                FontSize = ColumnChartSetting.ChartAxesOptions.HorizontalFontSize,
                IsBold = ColumnChartSetting.ChartAxesOptions.IsHorizontalBold,
                IsItalic = ColumnChartSetting.ChartAxesOptions.IsHorizontalItalic,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936,
                CrossAxisId = 1362418656,
                FontSize = ColumnChartSetting.ChartAxesOptions.VerticalFontSize,
                IsBold = ColumnChartSetting.ChartAxesOptions.IsVerticalBold,
                IsItalic = ColumnChartSetting.ChartAxesOptions.IsVerticalItalic,
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
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(SolidFill);
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateCategoryAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!));
            series.Append(CreateValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!));
            if (ChartDataGrouping.DataLabelCells != null && ChartDataGrouping.DataLabelFormula != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(ChartDataGrouping.DataLabelFormula, ChartDataGrouping.DataLabelCells.Skip(1).ToArray())
                )
                { Uri = GeneratorUtils.GenerateNewGUID() }));
            }
            return series;
        }

        private C.DataLabels? CreateColumnDataLabels(ColumnChartDataLabel ColumnChartDataLabel, int? DataLabelCounter = 0)
        {
            if (ColumnChartDataLabel.ShowValue || ColumnChartDataLabel.ShowValueFromColumn || ColumnChartDataLabel.ShowCategoryName || ColumnChartDataLabel.ShowLegendKey || ColumnChartDataLabel.ShowSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(ColumnChartDataLabel, DataLabelCounter);
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = ColumnChartDataLabel.DataLabelPosition switch
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
                    FontSize = (int)ColumnChartDataLabel.FontSize * 100,
                    Bold = ColumnChartDataLabel.IsBold,
                    Italic = ColumnChartDataLabel.IsItalic,
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