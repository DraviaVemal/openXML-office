/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class PieFamilyChart : ChartBase
    {
        #region Protected Fields

        protected PieChartSetting PieChartSetting;

        #endregion Protected Fields

        #region Protected Constructors

        protected PieFamilyChart(PieChartSetting PieChartSetting, ChartData[][] DataCols) : base(PieChartSetting)
        {
            this.PieChartSetting = PieChartSetting;
            switch (PieChartSetting.PieChartTypes)
            {
                case PieChartTypes.DOUGHNUT:
                    SetChartPlotArea(CreateChartPlotArea(DataCols));
                    break;

                default:
                    SetChartPlotArea(CreateChartPlotArea(DataCols));
                    break;
            };
        }

        #endregion Protected Constructors

        #region Private Methods

        private C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            OpenXmlCompositeElement Chart = PieChartSetting.PieChartTypes == PieChartTypes.DOUGHNUT ? new C.DoughnutChart(
                new C.VaryColors { Val = true }) : new C.PieChart(
                new C.VaryColors { Val = true });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, PieChartSetting.ChartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < PieChartSetting.PieChartSeriesSettings.Count)
                    {
                        return CreatePieDataLabels(PieChartSetting.PieChartSeriesSettings?[seriesIndex]?.PieChartDataLabel ?? new PieChartDataLabel(), Series.DataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                Chart.Append(CreateChartSeries(seriesIndex, Series, PieChartSetting.PieChartSeriesSettings.Count > seriesIndex ? PieChartSetting.PieChartSeriesSettings[seriesIndex] : new PieChartSeriesSetting(),
                    CreateSolidFill(PieChartSetting.PieChartSeriesSettings
                            .Where(item => item.FillColor != null)
                            .Select(item => item.FillColor!)
                            .ToList(), seriesIndex),
                    GetDataLabels()));
                seriesIndex++;
            });
            C.DataLabels? DataLabels = CreatePieDataLabels(PieChartSetting.PieChartDataLabel);
            if (DataLabels != null)
            {
                Chart.Append(DataLabels);
            }
            Chart.Append(new C.FirstSliceAngle { Val = 0 });
            Chart.Append(new C.HoleSize { Val = 50 });
            plotArea.Append(Chart);
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.PieChartSeries CreateChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, PieChartSeriesSetting PieChartSeriesSetting, A.SolidFill SolidFill, C.DataLabels? DataLabels)
        {
            C.PieChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.SeriesHeaderFormula!, new[] { ChartDataGrouping.SeriesHeaderCells! }));
            for (uint index = 0; index < ChartDataGrouping.XaxisCells!.Length; index++)
            {
                C.DataPoint DataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
                C.ShapeProperties ShapeProperties = CreateShapeProperties();
                ShapeProperties.Append(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(index % 6) + 1}") }));
                if (PieChartSetting.PieChartTypes == PieChartTypes.DOUGHNUT)
                {
                    ShapeProperties.Append(new A.Outline(new A.NoFill()));
                }
                else
                {
                    ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 })) { Width = 19050 });
                }
                ShapeProperties.Append(new A.EffectList());
                // series.Append(DataLabels);
                DataPoint.Append(ShapeProperties);
                series.Append(DataPoint);
            }
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

        private C.DataLabels? CreatePieDataLabels(PieChartDataLabel PieChartDataLabel, int? DataLabelCounter = 0)
        {
            if (PieChartDataLabel.ShowValue || PieChartDataLabel.ShowCategoryName || PieChartDataLabel.ShowLegendKey || PieChartDataLabel.ShowSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(PieChartDataLabel, DataLabelCounter);
                if (PieChartSetting.PieChartTypes == PieChartTypes.DOUGHNUT &&
                    new[] { PieChartDataLabel.DataLabelPositionValues.CENTER, PieChartDataLabel.DataLabelPositionValues.INSIDE_END, PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END, PieChartDataLabel.DataLabelPositionValues.BEST_FIT }.Contains(PieChartDataLabel.DataLabelPosition))
                    DataLabels.InsertAt(new C.DataLabelPosition()
                    {
                        Val = PieChartDataLabel.DataLabelPosition switch
                        {
                            PieChartDataLabel.DataLabelPositionValues.INSIDE_END => C.DataLabelPositionValues.InsideEnd,
                            PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END => C.DataLabelPositionValues.OutsideEnd,
                            PieChartDataLabel.DataLabelPositionValues.BEST_FIT => C.DataLabelPositionValues.BestFit,
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