using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class ScatterFamilyChart : ChartBase
    {
        #region Protected Fields

        protected ScatterChartSetting ScatterChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

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
            C.ScatterChart ScatterChart = new(
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
                },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, ScatterChartSetting.ChartDataSetting)
            .ForEach(Series =>
            {
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
                ScatterChart.Append(CreateScatterChartSeries(seriesIndex, Series, ScatterChartSetting.ScatterChartSeriesSettings.Count > seriesIndex ? ScatterChartSetting.ScatterChartSeriesSettings[seriesIndex] : new ScatterChartSeriesSetting(), Marker,
                     ScatterChartSetting.ScatterChartTypes == ScatterChartTypes.SCATTER ? new A.Outline(new A.NoFill()) : new A.Outline(CreateSolidFill(ScatterChartSetting.ScatterChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex)),
                    GetDataLabels(ScatterChartSetting, seriesIndex)));
                seriesIndex++;
            });

            ScatterChart.Append(new C.AxisId { Val = 1362418656 });
            ScatterChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(ScatterChart);
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1362418656,
                AxisPosition = AxisPosition.BOTTOM
            }));
            plotArea.Append(CreateValueAxis(new ValueAxisSetting()
            {
                Id = 1358349936
            }));
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.DataLabels? CreateDataLabel(ScatterChartDataLabel ScatterChartDataLabel)
        {
            if (ScatterChartDataLabel.GetType().GetProperties()
                .Where(Prop => Prop.PropertyType == typeof(bool))
                .Any(Prop => (bool)Prop.GetValue(ScatterChartDataLabel)!))
            {
                C.DataLabels DataLabels = new(
                    new C.ShowLegendKey { Val = ScatterChartDataLabel.ShowLegendKey },
                    new C.ShowValue { Val = ScatterChartDataLabel.ShowValue },
                    new C.ShowCategoryName { Val = ScatterChartDataLabel.ShowCategoryName },
                    new C.ShowSeriesName { Val = ScatterChartDataLabel.ShowSeriesName },
                    new C.ShowPercent { Val = false },
                    new C.ShowBubbleSize { Val = false },
                    new C.ShowLeaderLines() { Val = false });
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = ScatterChartDataLabel.DataLabelPosition switch
                    {
                        ScatterChartDataLabel.eDataLabelPosition.LEFT => C.DataLabelPositionValues.Left,
                        ScatterChartDataLabel.eDataLabelPosition.RIGHT => C.DataLabelPositionValues.Right,
                        ScatterChartDataLabel.eDataLabelPosition.ABOVE => C.DataLabelPositionValues.Top,
                        ScatterChartDataLabel.eDataLabelPosition.BELOW => C.DataLabelPositionValues.Bottom,
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

        private C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, ScatterChartSeriesSetting ScatterChartSeriesSetting, C.Marker Marker, A.Outline Outline, C.DataLabels? DataLabels)
        {
            C.ScatterChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }),
                Marker);
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(Outline);
            ShapeProperties.Append(new A.EffectList());
            if (DataLabels != null)
            {
                series.Append(DataLabels);
            }
            series.Append(ShapeProperties);
            series.Append(CreateXValueAxisData(ChartDataGrouping.XaxisFormula!, ChartDataGrouping.XaxisCells!, ScatterChartSeriesSetting));
            series.Append(CreateYValueAxisData(ChartDataGrouping.YaxisFormula!, ChartDataGrouping.YaxisCells!, ScatterChartSeriesSetting));
            series.Append(new C.Smooth() { Val = new[] { ScatterChartTypes.SCATTER_SMOOTH, ScatterChartTypes.SCATTER_SMOOTH_MARKER }.Contains(ScatterChartSetting.ScatterChartTypes) });
            return series;
        }

        private C.DataLabels? GetDataLabels(ScatterChartSetting ScatterChartSetting, int index)
        {
            if (index < ScatterChartSetting.ScatterChartSeriesSettings.Count)
            {
                return CreateDataLabel(ScatterChartSetting.ScatterChartSeriesSettings?[index]?.ScatterChartDataLabel ?? new ScatterChartDataLabel());
            }
            return null;
        }

        #endregion Private Methods
    }
}