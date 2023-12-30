using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class BarFamilyChart : ChartBase
    {
        #region Protected Methods

        protected C.PlotArea CreateChartPlotArea(ChartData[][] DataCols, C.BarDirectionValues barDirectionValue, C.BarGroupingValues barGroupingValue)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.BarChart BarChart = new(
                new C.BarDirection { Val = barDirectionValue },
                new C.BarGrouping { Val = barGroupingValue },
                new C.VaryColors { Val = false });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                BarChart.Append(CreateBarChartSeries(seriesIndex++,
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
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false });
            BarChart.Append(DataLabels);
            if (barGroupingValue == C.BarGroupingValues.Clustered)
            {
                BarChart.Append(new C.GapWidth { Val = 219 });
                BarChart.Append(new C.Overlap { Val = -27 });
            }
            else
            {
                BarChart.Append(new C.GapWidth { Val = 150 });
                BarChart.Append(new C.Overlap { Val = 100 });
            }
            BarChart.Append(new C.AxisId { Val = 1362418656 });
            BarChart.Append(new C.AxisId { Val = 1358349936 });
            plotArea.Append(BarChart);
            plotArea.Append(CreateCategoryAxis(1362418656));
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

        private C.BarChartSeries CreateBarChartSeries(int seriesIndex, string seriesTextFormula, ChartData[] seriesTextCells, string categoryFormula, ChartData[] categoryCells, string valueFormula, ChartData[] valueCells, string accent)
        {
            C.BarChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))),
                new C.InvertIfNegative { Val = true });
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues(accent) }));
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            series.Append(ShapeProperties);
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            series.Append(new C.Smooth()
            {
                Val = false
            });
            return series;
        }

        #endregion Private Methods
    }
}