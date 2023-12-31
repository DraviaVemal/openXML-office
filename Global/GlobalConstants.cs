namespace OpenXMLOffice.Global
{
    public static class GlobalConstants
    {
        #region Public Enums
        public const int DefaultDPI = 96;
        public const int DefaultSlideWidthEmu = 9144000;
        public const int DefaultSlideHeightEmu = 6858000;
        public enum ColumnChartTypes
        {
            CLUSTERED,
            STACKED,
            PERCENT_STACKED,
            // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D, COLUMN_3D
        }

        public enum BarChartTypes
        {
            CLUSTERED,
            STACKED,
            PERCENT_STACKED,
            // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D,
        }

        public enum LineChartTypes
        {
            CLUSTERED,
            STACKED,
            PERCENT_STACKED,
            CLUSTERED_MARKER,
            STACKED_MARKER,
            PERCENT_STACKED_MARKER,
            // CLUSTERED_3D
        }

        public enum AreaChartTypes
        {
            CLUSTERED,
            STACKED,
            PERCENT_STACKED,
            // CLUSTERED_3D, STACKED_3D, PERCENT_STACKED_3D
        }

        public enum PieChartTypes
        {
            PIE,

            // PIE_3D, PIE_PIE, PIE_BAR,
            DOUGHNUT
        }

        #endregion Public Enums
    }
}