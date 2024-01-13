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
    /// Represents the types of scatter charts.
    /// </summary>
    public class ScatterFamilyChart : ChartBase
    {
        #region Protected Fields

        /// <summary>
        /// Scatter Chart Setting
        /// </summary>
        protected ScatterChartSetting ScatterChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        /// <summary>
        /// Create Scatter Chart with provided settings
        /// </summary>
        /// <param name="ScatterChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        protected ScatterFamilyChart(ScatterChartSetting ScatterChartSetting, ChartData[][] DataCols) : base(ScatterChartSetting)
        {
            this.ScatterChartSetting = ScatterChartSetting;
            SetChartPlotArea(CreateChartPlotArea(DataCols));
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            OpenXmlCompositeElement Chart = ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.BUBBLE ? new C.BubbleChart() : new C.ScatterChart(
                new C.ScatterStyle
                {
                    Val = ScatterChartSetting.ScatterChartTypes switch
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
            if (ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                ScatterChartSetting.ChartDataSetting.Is3Ddata = true;
                if ((DataCols.Length - 1) % 2 != 0)
                {
                    throw new ArgumentOutOfRangeException("Required 3D Data Size is not met.");
                }
            }
            int seriesIndex = 0;
            CreateDataSeries(DataCols, ScatterChartSetting.ChartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < ScatterChartSetting.ScatterChartSeriesSettings.Count)
                    {
                        return CreateScatterDataLabels(ScatterChartSetting.ScatterChartSeriesSettings?[seriesIndex]?.ScatterChartDataLabel ?? new ScatterChartDataLabel(), Series.DataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                C.Marker Marker = new[] { ScatterChartTypes.SCATTER, ScatterChartTypes.SCATTER_SMOOTH_MARKER, ScatterChartTypes.SCATTER_STRIGHT_MARKER }.Contains(ScatterChartSetting.ScatterChartTypes) ? new(
                    new C.Symbol { Val = ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.SCATTER ? C.MarkerStyleValues.Auto : C.MarkerStyleValues.Circle },
                    new C.Size { Val = 5 },
                    new C.ShapeProperties(
                        new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(seriesIndex % 6) + 1}") }),
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(seriesIndex % 6) + 1}") })),
                        new A.EffectList()
                    )) :
                    new(new C.Symbol()
                    {
                        Val = C.MarkerStyleValues.None
                    });
                Chart.Append(CreateScatterChartSeries(seriesIndex, Series,
                    ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.BUBBLE ? null : Marker,
                    ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.SCATTER ? new A.Outline(new A.NoFill()) : new A.Outline(CreateSolidFill(ScatterChartSetting.ScatterChartSeriesSettings
                            .Where(item => item?.FillColor != null)
                            .Select(item => item?.FillColor!)
                            .ToList(), seriesIndex)),
                    GetDataLabels()));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreateScatterDataLabels(ScatterChartSetting.ScatterChartDataLabel);
            if (DataLabels != null)
            {
                Chart.Append(DataLabels);
            }
            if (ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                Chart.Append(new C.BubbleScale() { Val = 100 });
                Chart.Append(new C.ShowNegativeBubbles() { Val = false });
            }
            Chart.Append(new C.AxisId { Val = 1362418656 });
            Chart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(Chart);
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1362418656,
                AxisPosition = AxisPosition.BOTTOM,
                CrossAxisId = 1358349936,
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936,
                CrossAxisId = 1362418656
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, C.Marker? Marker, A.Outline Outline, C.DataLabels? DataLabels)
        {
            C.ScatterChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }));
            if (Marker != null)
            {
                series.Append(Marker);
            }
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            if (ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                ShapeProperties.Append(new A.SolidFill(new A.SchemeColor(new A.Alpha() { Val = 75000 }) { Val = A.SchemeColorValues.Accent1 }));
                ShapeProperties.Append(new A.Outline(new A.NoFill()));
                series.Append(new C.InvertIfNegative() { Val = false });
            }
            else
            {
                ShapeProperties.Append(Outline);
                ShapeProperties.Append(new A.EffectList());
            }
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateXValueAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!));
            series.Append(CreateYValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!));
            if (ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.BUBBLE)
            {
                series.Append(CreateBubbleSizeAxisData(ChartDataGrouping.ZaxisFormula!, ChartDataGrouping.ZaxisCells!));
                series.Append(new C.Bubble3D() { Val = false });
            }
            else
            {
                series.Append(new C.Smooth() { Val = new[] { ScatterChartTypes.SCATTER_SMOOTH, ScatterChartTypes.SCATTER_SMOOTH_MARKER }.Contains(ScatterChartSetting.ScatterChartTypes) });
            }
            if (ChartDataGrouping.DataLabelCells != null && ChartDataGrouping.DataLabelFormula != null)
            {
                series.Append(new C.ExtensionList(new C.Extension(
                    CreateDataLabelsRange(ChartDataGrouping.DataLabelFormula, ChartDataGrouping.DataLabelCells.Skip(1).ToArray())
                )
                { Uri = GeneratorUtils.GenerateNewGUID() }));
            }
            return series;
        }

        private C.DataLabels? CreateScatterDataLabels(ScatterChartDataLabel ScatterChartDataLabel, int? DataLabelCounter = 0)
        {
            if (ScatterChartDataLabel.ShowValue || ScatterChartDataLabel.ShowValueFromColumn || ScatterChartDataLabel.ShowCategoryName || ScatterChartDataLabel.ShowLegendKey || ScatterChartDataLabel.ShowSeriesName || ScatterChartDataLabel.ShowBubbleSize || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(ScatterChartDataLabel, DataLabelCounter);
                DataLabels.Append(new C.ShowBubbleSize { Val = ScatterChartDataLabel.ShowBubbleSize });
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = ScatterChartDataLabel.DataLabelPosition switch
                    {
                        ScatterChartDataLabel.DataLabelPositionValues.LEFT => C.DataLabelPositionValues.Left,
                        ScatterChartDataLabel.DataLabelPositionValues.RIGHT => C.DataLabelPositionValues.Right,
                        ScatterChartDataLabel.DataLabelPositionValues.ABOVE => C.DataLabelPositionValues.Top,
                        ScatterChartDataLabel.DataLabelPositionValues.BELOW => C.DataLabelPositionValues.Bottom,
                        //Center
                        _ => C.DataLabelPositionValues.Center,
                    }
                }, 0);
                DataLabels.InsertAt(new C.ShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()), new A.EffectList()), 0);
                A.Paragraph Paragraph = new(new A.ParagraphProperties(new A.DefaultRunProperties(
                    new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation() { Val = 75000 }, new A.LuminanceOffset() { Val = 25000 }) { Val = A.SchemeColorValues.Text1 }),
                    new A.LatinFont() { Typeface = "+mn-lt" }, new A.EastAsianFont() { Typeface = "+mn-ea" }, new A.ComplexScriptFont() { Typeface = "+mn-cs" })
                {
                    FontSize = 1197,
                    Bold = false,
                    Italic = false,
                    Underline = A.TextUnderlineValues.None,
                    Strike = A.TextStrikeValues.NoStrike,
                    Kerning = 1200,
                    Baseline = 0
                }), new A.EndParagraphRunProperties() { Language = "en-US" });
                DataLabels.InsertAt(new C.TextProperties(new A.BodyProperties(new A.ShapeAutoFit())
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
               Paragraph), 0);
                return DataLabels;
            }
            return null;
        }

        #endregion Private Methods
    }
}