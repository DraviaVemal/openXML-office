namespace OpenXMLOffice.Global
{
    public enum PieChartTypes
    {
        PIE,

        // PIE_3D, PIE_PIE, PIE_BAR,
        DOUGHNUT
    }

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
        public PieChartTypes PieChartTypes = PieChartTypes.PIE;
        public PieChartDataLabel PieChartDataLabel = new();
        public List<PieChartSeriesSetting> PieChartSeriesSettings = new();
    }
}