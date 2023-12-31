namespace OpenXMLOffice.Global
{
    public class PieChartDataLabel
    {
        public enum eDataLabelPosition
        {
            NONE,
            SHOW,
            // CALLOUT
        }

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;
    }
    public class PieChartSetting : ChartSetting
    {

        public PieChartDataLabel PieChartDataLabel = new();
    }
}