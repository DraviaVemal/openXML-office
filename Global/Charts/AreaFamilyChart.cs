using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class AreaFamilyChart : ChartBase
    {
        #region Protected Methods

        protected C.PlotArea CreateChartPlotArea(ChartData[][] DataCols, C.GroupingValues groupingValue)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.AreaChart AreaChart = new(
                new C.Grouping() { Val = groupingValue },
                new C.VaryColors() { Val = false });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                AreaChart.Append(CreateAreaChartSeries(seriesIndex++,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}${DataCols[0].Length}",
                    col.Skip(1).ToArray(),
                    $"accent{(seriesIndex % 6) + 1}"
                ));
            }
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey() { Val = false },
                new C.ShowValue() { Val = false },
                new C.ShowCategoryName() { Val = false },
                new C.ShowSeriesName() { Val = false },
                new C.ShowPercent() { Val = false },
                new C.ShowBubbleSize() { Val = false });
            AreaChart.Append(DataLabels);
            AreaChart.Append(new C.AxisId() { Val = 1362418656 });
            AreaChart.Append(new C.AxisId() { Val = 1358349936 });
            plotArea.Append(AreaChart);
            plotArea.Append(CreateCategoryAxis(1362418656, $"Sheet1!$A$2:$A${DataCols[0].Length}"));
            plotArea.Append(CreateValueAxis(1358349936));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        #endregion Protected Methods

        #region Private Methods

        private C.AreaChartSeries CreateAreaChartSeries(int seriesIndex, string seriesTextFormula, ChartData[] seriesTextCells, string categoryFormula, ChartData[] categoryCells, string valueFormula, ChartData[] valueCells, string accent)
        {
            C.AreaChartSeries series = new(
                new C.Index() { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order() { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))));
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.SchemeColor() { Val = new A.SchemeColorValues(accent) }), new A.Outline(new A.NoFill())));
            ShapeProperties.Append(new A.EffectList());
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            return series;
        }

        private C.CategoryAxis CreateCategoryAxis(UInt32Value axisId, string formula)
        {
            C.CategoryAxis CategoryAxis = new(
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
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            CategoryAxis.Append(ShapeProperties);
            return CategoryAxis;
        }

        private C.ValueAxis CreateValueAxis(UInt32Value axisId)
        {
            C.ValueAxis ValueAxis = new(
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
                new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory });
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            ValueAxis.Append(ShapeProperties);
            return ValueAxis;
        }

        #endregion Private Methods
    }
}