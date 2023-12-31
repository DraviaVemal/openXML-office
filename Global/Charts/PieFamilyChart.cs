using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class PieFamilyChart : ChartBase
    {
        #region Protected Methods
        protected C.PlotArea CreateDoughnutChartPlotArea(ChartData[][] DataCols, PieChartSetting chartSetting)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.DoughnutChart DoughnutChart = new(
                new C.VaryColors { Val = true });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).ToArray())
            {
                DoughnutChart.Append(CreateChartSeries(seriesIndex,
                    chartSetting,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}${DataCols[0].Length}",
                    col.Skip(1).ToArray()
                ));
                seriesIndex++;
            }
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines { Val = true });
            DoughnutChart.Append(DataLabels);
            DoughnutChart.Append(new C.FirstSliceAngle { Val = 0 });
            DoughnutChart.Append(new C.HoleSize { Val = 50 });
            plotArea.Append(DoughnutChart);
            C.ShapeProperties ShapeProperties = new();
            ShapeProperties.Append(new A.NoFill());
            ShapeProperties.Append(new A.Outline(new A.NoFill()));
            ShapeProperties.Append(new A.EffectList());
            plotArea.Append(ShapeProperties);
            return plotArea;
        }

        protected C.PlotArea CreatePieChartPlotArea(ChartData[][] DataCols, PieChartSetting chartSetting)
        {
            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());
            C.PieChart PieChart = new(
                new C.VaryColors { Val = true });
            int seriesIndex = 0;
            foreach (ChartData[] col in DataCols.Skip(1).Take(1).ToArray())
            {
                PieChart.Append(CreateChartSeries(seriesIndex,
                    chartSetting,
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$1",
                    col.Take(1).ToArray(),
                    $"Sheet1!$A$2:$A${DataCols[0].Length}",
                    DataCols[0].Skip(1).ToArray(),
                    $"Sheet1!${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}$2:${ConverterUtils.ConvertIntToColumnName(seriesIndex + 1)}${DataCols[0].Length}",
                    col.Skip(1).ToArray()));
                seriesIndex++;
            }
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines { Val = true });
            PieChart.Append(DataLabels);
            PieChart.Append(new C.FirstSliceAngle { Val = 0 });
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
        private C.DataLabels CreateDataLabel(PieChartDataLabel PieChartDataLabel)
        {
            C.DataLabels DataLabels = new(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = PieChartDataLabel.DataLabelPosition != PieChartDataLabel.eDataLabelPosition.NONE },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false },
                new C.ShowLeaderLines() { Val = false });
            if (PieChartDataLabel.DataLabelPosition != PieChartDataLabel.eDataLabelPosition.NONE)
            {
                DataLabels.InsertAt(new C.DataLabelPosition()
                {
                    Val = PieChartDataLabel.DataLabelPosition switch
                    {
                        //Show
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
            }
            return DataLabels;
        }
        private C.PieChartSeries CreateChartSeries(int seriesIndex, PieChartSetting PieChartSetting, string seriesTextFormula, ChartData[] seriesTextCells, string categoryFormula, ChartData[] categoryCells, string valueFormula, ChartData[] valueCells, bool IsDoughnut = false)
        {
            C.PieChartSeries series = new(
                new C.Index { Val = new UInt32Value((uint)seriesIndex) },
                new C.Order { Val = new UInt32Value((uint)seriesIndex) },
                new C.SeriesText(new C.StringReference(new C.Formula(seriesTextFormula), AddStringCacheValue(seriesTextCells))));
            for (uint index = 0; index < categoryCells.Length; index++)
            {
                C.DataPoint DataPoint = new(new C.Index { Val = index }, new C.Bubble3D { Val = false });
                C.ShapeProperties ShapeProperties = new();
                ShapeProperties.Append(new A.SolidFill(new A.SchemeColor { Val = new A.SchemeColorValues($"accent{(index % 6) + 1}") }));
                if (IsDoughnut)
                {
                    ShapeProperties.Append(new A.Outline(new A.NoFill()));
                }
                else
                {
                    ShapeProperties.Append(new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 })) { Width = 19050 });
                }
                ShapeProperties.Append(new A.EffectList());
                series.Append(CreateDataLabel(PieChartSetting.SeriesSettings?[seriesIndex]?.PieChartDataLabel ?? new PieChartDataLabel()));
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