namespace OpenXMLOffice.Global
{
    public class ColumnChartDataLabel
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
    public class ColumnChartSetting : ChartSetting
    {
        public ChartAxisOptions ChartAxisOptions = new();
        public ChartAxesOptions ChartAxesOptions = new();
        public ChartGridLinesOptions ChartGridLinesOptions = new();
        public ColumnChartDataLabel ColumnChartDataLabel = new();
    }
}