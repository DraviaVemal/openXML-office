namespace OpenXMLOffice.Global
{
    public enum PieChartTypes
    {
        PIE,

        // PIE_3D, PIE_PIE, PIE_BAR,
        DOUGHNUT
    }

    public class PieChartDataLabel
    {
        #region Public Fields

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.SHOW;
        public bool ShowCategoryName = false;
        public bool ShowLegendKey = false;
        public bool ShowSeriesName = false;
        public bool ShowValue = false;

        #endregion Public Fields

        #region Public Enums

        public enum eDataLabelPosition
        {
            SHOW,
            // CALLOUT
        }

        #endregion Public Enums
    }

    public class PieChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public string? BorderColor;
        public string? FillColor;
        public PieChartDataLabel PieChartDataLabel = new();

        #endregion Public Fields
    }

    public class PieChartSetting : ChartSetting
    {
        #region Public Fields

        public PieChartDataLabel PieChartDataLabel = new();
        public List<PieChartSeriesSetting> PieChartSeriesSettings = new();
        public PieChartTypes PieChartTypes = PieChartTypes.PIE;

        #endregion Public Fields
    }
}