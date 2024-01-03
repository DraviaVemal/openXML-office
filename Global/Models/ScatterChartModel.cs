namespace OpenXMLOffice.Global
{
    public enum ScatterChartTypes
    {
        SCATTER,
        SCATTER_SMOOTH,
        SCATTER_SMOOTH_MARKER,
        SCATTER_STRIGHT,
        SCATTER_STRIGHT_MARKER,
        BUBBLE,
        // BUBBLE_3D
    }

    public class ScatterChartDataLabel
    {
        #region Public Fields

        public eDataLabelPosition DataLabelPosition = eDataLabelPosition.NONE;

        #endregion Public Fields

        #region Public Enums

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

        #endregion Public Enums
    }

    public class ScatterChartSeriesSetting : ChartSeriesSetting
    {
        #region Public Fields

        public string? BorderColor;
        public string? FillColor;
        public ScatterChartDataLabel ScatterChartDataLabel = new();

        #endregion Public Fields
    }

    public class ScatterChartSetting : ChartSetting
    {
        #region Public Fields

        public ChartAxesOptions ChartAxesOptions = new();
        public ChartAxisOptions ChartAxisOptions = new();
        public List<ScatterChartSeriesSetting> ScatterChartSeriesSettings = new();
        public ScatterChartTypes ScatterChartTypes = ScatterChartTypes.SCATTER;

        #endregion Public Fields
    }
}