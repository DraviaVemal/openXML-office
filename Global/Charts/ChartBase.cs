using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global;
public class ChartBase
{
    public ChartSpace CreateChartSpace()
    {
        ChartSpace ChartSpace = new();
        Chart Chart = CreateChart();
        Title Title = CreateTitle();
        PlotArea PlotArea = CreatePlotArea();
        Layout Layout = CreateLayout();
        Legend Legend = CreateLegend();
        Chart.Append(Title);
        PlotArea.Append(Layout);
        Chart.Append(PlotArea);
        Chart.Append(Legend);
        ChartSpace.Append(Chart);
        return ChartSpace;
    }

    protected Chart CreateChart()
    {
        return new();
    }

    protected Title CreateTitle()
    {
        return new();
    }

    protected PlotArea CreatePlotArea()
    {
        return new();
    }

    protected Legend CreateLegend()
    {
        return new();
    }

    protected Layout CreateLayout()
    {
        return new();
    }
}
