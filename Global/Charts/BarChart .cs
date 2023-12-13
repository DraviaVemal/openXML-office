using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global;
public class BarChart : BarFamilyChart
{
    public ChartSpace GetChartSpace()
    {
        return CreateChartSpace();
    }

    public ChartStyle GetChartStyle()
    {
        return CreateChartStyles();
    }
    public ColorStyle GetColorStyle()
    {
        return CreateColorStyles();
    }
}
