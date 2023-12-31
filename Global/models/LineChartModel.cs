namespace OpenXMLOffice.Global
{
    public class LineChartDataLabel
    {
        public enum eDataLabelPosition
        {
            NONE,
            LEFT,
            RIGHT,
            CENTER,
            ABOVE,
            BELOW,
            // CALLOUT
        }

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;
    }
    public class LineChartSetting : ChartSetting
    {
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public LineChartDataLabel LineChartDataLabel = new();
    }
}