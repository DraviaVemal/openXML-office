using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class PieFamilyChart : ChartBase
    {
        #region Protected Methods

        protected C.PlotArea CreateDoughnutChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.DoughnutChart DoughnutChart = new(
                new C.VaryColors() { Val = true });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                DoughnutChart.Append(CreateChartSeries(seriesIndex++,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}${DataCols[0].Length}",
                    col.Skip(1).ToArray()
                ));
            }
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey() { Val = false },
                new C.ShowValue() { Val = false },
                new C.ShowCategoryName() { Val = false },
                new C.ShowSeriesName() { Val = false },
                new C.ShowPercent() { Val = false },
                new C.ShowBubbleSize() { Val = false },
                new C.ShowLeaderLines() { Val = true });
            DoughnutChart.Append(DataLabels);
            DoughnutChart.Append(new C.FirstSliceAngle() { Val = 0 });
            DoughnutChart.Append(new C.HoleSize() { Val = 50 });
            plotArea.Append(DoughnutChart);
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        protected C.PlotArea CreatePieChartPlotArea(ChartData[][] DataCols)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.PieChart PieChart = new(
                new C.VaryColors() { Val = true });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).Take(1).ToArray())
            {
                PieChart.Append(CreateChartSeries(seriesIndex++,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}${DataCols[0].Length}",
                    col.Skip(1).ToArray()
                ));
            }
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey() { Val = false },
                new C.ShowValue() { Val = false },
                new C.ShowCategoryName() { Val = false },
                new C.ShowSeriesName() { Val = false },
                new C.ShowPercent() { Val = false },
                new C.ShowBubbleSize() { Val = false },
                new C.ShowLeaderLines() { Val = true });
            PieChart.Append(DataLabels);
            PieChart.Append(new C.FirstSliceAngle() { Val = 0 });
            plotArea.Append(PieChart);
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        #endregion Protected Methods

        #region Private Methods

        private C.PieChartSeries CreateChartSeries(int seriesIndex, string seriesTextFormula, ChartData[] seriesTextCells, string categoryFormula, ChartData[] categoryCells, string valueFormula, ChartData[] valueCells, bool IsDoughnut = false)
        {
            C.PieChartSeries series = new(
                new C.Index() { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order() { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))));
            for (uint index = 0; index < categoryCells.Length; index++)
            {
                C.DataPoint DataPoint = new(new C.Index() { Val = index }, new C.Bubble3D() { Val = false });
                C.ShapeProperties ShapeProperties = new();
                ShapeProperties.Append(new A.SolidFill(new A.SchemeColor() { Val = new A.SchemeColorValues($"accent{(index % 6) + 1}") }));
                if (IsDoughnut)
                {
                    ShapeProperties.Append(new A.Outline(new A.NoFill()));
                }
                else
                {
                    ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.Light1 })) { Width = 19050 });
                }
                ShapeProperties.Append(new A.EffectList());
                DataPoint.Append(ShapeProperties);
                series.Append(DataPoint);
            }
            series.Append(new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), AddStringCacheValue(categoryCells))));
            series.Append(new C.Values(new C.NumberReference(new C.Formula(valueFormula), AddNumberCacheValue(valueCells, null))));
            return series;
        }

        #endregion Private Methods
    }
}