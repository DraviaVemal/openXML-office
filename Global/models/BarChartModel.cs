namespace OpenXMLOffice.Global
{

    public class BarChartDataLabel
    {
        public enum eDataLabelPosition
        {
            NONE,
            CENTER,
            INSIDE_END,
            INSIDE_BASE,
            OUTSIDE_END,
            // CALLOUT
        }

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;
    }

    public class BarChartSetting : ChartSetting
    {
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public BarChartDataLabel BarChartDataLabel = new();
    }
}