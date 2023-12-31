namespace OpenXMLOffice.Global
{
    public class PieChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields
        public string? FillColor;
        public string? BorderColor;
        public PieChartDataLabel PieChartDataLabel = new();
        #endregion Public Fields
    }

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
        public List<PieChartSeriesSetting>? SeriesSettings;
    }
}