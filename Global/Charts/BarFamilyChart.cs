using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class BarFamilyChart : ChartBase
    {
        #region Protected Methods

        private C.BarChartSeries CreateBarChartSeries(int seriesIndex, string seriesTextFormula, List<ChartData> seriesTextCells, string categoryFormula, List<ChartData> categoryCells, string valueFormula, List<ChartData> valueCells, string accent)
        {
            C.BarChartSeries series = new(
                new C.Index() { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order() { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))),
                new C.InvertIfNegative() { Val = false });
            C.ShapeProperties spPr = new();
            spPr.Append(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }));
            spPr.Append(new A.Outline(new A.NoFill()));
            spPr.Append(new A.EffectList());
            series.Append(spPr);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            return series;
        }

        private C.ValueAxis CreateValueAxis(UInt32Value axisId)
        {
            C.ValueAxis valAx = new(
                new C.AxisId() { Val = axisId },
                new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
                new C.Delete() { Val = false },
                new C.AxisPosition() { Val = C.AxisPositionValues.Left },
                new C.MajorGridlines(
                    new C.ShapeProperties(
                        new A.Outline(
                            new A.SolidFill(
                                new A.SchemeColor(
                                    new A.LuminanceModulation() { Val = 15000 },
                                    new A.LuminanceOffset() { Val = 85000 })
                                { Val = A.SchemeColorValues.Text1 }
                            )
                        )
                        {
                            Width = 9525,
                            CapType = A.LineCapValues.Flat,
                            CompoundLineType = A.CompoundLineValues.Single,
                            Alignment = A.PenAlignmentValues.Center
                        },
                        new A.Round()
                    )
                ),
                new C.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new C.MajorTickMark() { Val = C.TickMarkValues.None },
                new C.MinorTickMark() { Val = C.TickMarkValues.None },
                new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis() { Val = axisId },
                new C.Crosses() { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween() { Val = C.CrossBetweenValues.Between });
            C.ShapeProperties spPr = new();
            spPr.Append(new A.NoFill());
            spPr.Append(new A.Outline(new A.NoFill()));
            spPr.Append(new A.EffectList());
            valAx.Append(spPr);
            return valAx;
        }


        private C.CategoryAxis CreateCategoryAxis(UInt32Value axisId, string formula)
        {
            C.CategoryAxis catAx = new(
                new C.AxisId() { Val = axisId },
                new C.Scaling(new C.Orientation() { Val = C.OrientationValues.MinMax }),
                new C.Delete() { Val = false },
                new C.AxisPosition() { Val = C.AxisPositionValues.Bottom },
                new C.MajorTickMark() { Val = C.TickMarkValues.None },
                new C.MinorTickMark() { Val = C.TickMarkValues.None },
                new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis() { Val = axisId },
                new C.Crosses() { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled() { Val = true },
                new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset() { Val = 100 },
                new C.NoMultiLevelLabels() { Val = false });
            C.ShapeProperties spPr = new();
            spPr.Append(new A.NoFill());
            spPr.Append(new A.Outline(new A.NoFill()));
            spPr.Append(new A.EffectList());
            catAx.Append(spPr);
            return catAx;
        }

        protected C.PlotArea CreateChartPlotArea(ChartData[][] DataCols)
        {
            List<ChartData> chartDataList = new()
        {
            new ChartData { Value = "Series 1" },
        };
            List<ChartData> chartDataList1 = new()
        {
            new ChartData { Value = "Series 2" },
        };
            List<ChartData> chartDataList2 = new()
        {
            new ChartData { Value = "Category 1" },
            new ChartData { Value = "Category 2" },
            new ChartData { Value = "Category 3" },
            new ChartData { Value = "Category 4" },
        };
            List<ChartData> chartDataList3 = new()
        {
            new ChartData { Value = "4.3" },
            new ChartData { Value = "2.5" },
            new ChartData { Value = "3.5" },
            new ChartData { Value = "4.5" },
        };
            List<ChartData> chartDataList4 = new()
        {
            new ChartData { Value = "2.4" },
            new ChartData { Value = "4.4000000000000004" },
            new ChartData { Value = "1.8" },
            new ChartData { Value = "2.8" },
        };
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.BarChart BarChart = new(
                new C.BarDirection() { Val = C.BarDirectionValues.Bar },
                new C.BarGrouping() { Val = C.BarGroupingValues.Clustered },
                new C.VaryColors() { Val = false });
            BarChart.Append(CreateBarChartSeries(0, "Sheet1!$B$1", chartDataList, "Sheet1!$A$2:$A$5", chartDataList2, "Sheet1!$B$2:$B$5", chartDataList3, "accent1"));
            BarChart.Append(CreateBarChartSeries(1, "Sheet1!$C$1", chartDataList1, "Sheet1!$A$2:$A$5", chartDataList2, "Sheet1!$C$2:$C$5", chartDataList4, "accent2"));
            C.DataLabels dLbls = new(
                new C.ShowLegendKey() { Val = false },
                new C.ShowValue() { Val = false },
                new C.ShowCategoryName() { Val = false },
                new C.ShowSeriesName() { Val = false },
                new C.ShowPercent() { Val = false },
                new C.ShowBubbleSize() { Val = false });
            BarChart.Append(dLbls);
            BarChart.Append(new C.GapWidth() { Val = 219 });
            BarChart.Append(new C.Overlap() { Val = -27 });
            BarChart.Append(new C.AxisId() { Val = 1362418656 });
            BarChart.Append(new C.AxisId() { Val = 1358349936 });
            plotArea.Append(BarChart);
            plotArea.Append(CreateCategoryAxis(1362418656, "Sheet1!$A$2:$A$5"));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties spPr = new();
            spPr.Append(new A.NoFill());
            spPr.Append(new A.Outline(new A.NoFill()));
            spPr.Append(new A.EffectList());
            plotArea.Append(spPr);
            return plotArea;
        }

        #endregion Protected Methods
    }
}