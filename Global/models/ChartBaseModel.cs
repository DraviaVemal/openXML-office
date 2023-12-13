namespace OpenXMLOffice.Global
{
    public class ChartData
    {
        #region Public Fields

        public string? Value;

        #endregion Public Fields
    }

    public class ChartSeriesSetting
    {
        #region Public Fields

        public string? NumberFormat;

        #endregion Public Fields
    }

    public class ChartSetting
    {
        #region Public Fields

        public List<ChartSeriesSetting>? SeriesSettings;

        #endregion Public Fields
    }
}