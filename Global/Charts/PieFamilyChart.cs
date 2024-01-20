// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
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
        /// <param name="PieChartSetting">
        /// </param>
        /// <param name="DataCols">
        /// </param>
        protected PieFamilyChart(PieChartSetting PieChartSetting, ChartData[][] DataCols) : base(PieChartSetting)
        {
            pieChartSetting = PieChartSetting;
            switch (PieChartSetting.pieChartTypes)
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
            OpenXmlCompositeElement Chart = pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT ? new C.DoughnutChart(
                new C.VaryColors { Val = true }) : new C.PieChart(
                new C.VaryColors { Val = true });
            int seriesIndex = 0;
            CreateDataSeries(DataCols, pieChartSetting.chartDataSetting).ForEach(Series =>
            {
                C.DataLabels? GetDataLabels()
                {
                    if (seriesIndex < pieChartSetting.pieChartSeriesSettings.Count)
                    {
                        return CreatePieDataLabels(pieChartSetting.pieChartSeriesSettings?[seriesIndex]?.pieChartDataLabel ?? new PieChartDataLabel(), Series.dataLabelCells?.Length ?? 0);
                    }
                    return null;
                }
                Chart.Append(CreateChartSeries(seriesIndex, Series, GetDataLabels()));
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
            C.ShapeProperties ShapeProperties = CreateShapeProperties();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        private C.PieChartSeries CreateChartSeries(int seriesIndex, ChartDataGrouping ChartDataGrouping, C.DataLabels? DataLabels)
        {
            C.PieChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                CreateSeriesText(ChartDataGrouping.seriesHeaderFormula!, new[] { ChartDataGrouping.seriesHeaderCells! }));
            for (uint index = 0; index < ChartDataGrouping.xAxisCells!.Length; index++)
            {
                C.DataPoint DataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
                C.ShapeProperties ShapeProperties = CreateShapeProperties();
                ShapeProperties.Append(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(index % 6) + 1}") }));
                if (pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT)
                {
                    ShapeProperties.Append(new A.Outline(new A.NoFill()));
                }
                else
                {
                    ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 })) { Width = 19050 });
                }
                ShapeProperties.Append(new A.EffectList());
                DataPoint.Append(ShapeProperties);
                if (DataLabels != null)
                {
                    series.Append(DataLabels);
                }
                series.Append(DataPoint);
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

        private C.DataLabels? CreatePieDataLabels(PieChartDataLabel PieChartDataLabel, int? DataLabelCounter = 0)
        {
            if (PieChartDataLabel.showValue || PieChartDataLabel.showCategoryName || PieChartDataLabel.showLegendKey || PieChartDataLabel.showSeriesName || DataLabelCounter > 0)
            {
                C.DataLabels DataLabels = CreateDataLabels(PieChartDataLabel, DataLabelCounter);
                if (pieChartSetting.pieChartTypes == PieChartTypes.DOUGHNUT &&
                    new[] { PieChartDataLabel.DataLabelPositionValues.CENTER, PieChartDataLabel.DataLabelPositionValues.INSIDE_END, PieChartDataLabel.DataLabelPositionValues.OUTSIDE_END, PieChartDataLabel.DataLabelPositionValues.BEST_FIT }.Contains(PieChartDataLabel.dataLabelPosition))
                {
                    throw new ArgumentException("DataLabelPosition is not supported for Doughnut Chart.");
                }
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = PieChartDataLabel.dataLabelPosition switch
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
                    FontSize = (int)PieChartDataLabel.fontSize * 100,
                    Bold = PieChartDataLabel.isBold,
                    Italic = PieChartDataLabel.isItalic,
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