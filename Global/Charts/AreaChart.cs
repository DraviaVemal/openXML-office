// Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License. See License in
// the project root for license information.
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class AreaChart : AreaFamilyChart
    {
        #region Public Constructors

        public AreaChart(AreaChartSetting AreaChartSetting, ChartData[][] DataCols) : base(AreaChartSetting, DataCols)
        {
        }

        #endregion Public Constructors

        #region Public Methods

        public CS.ChartStyle GetChartStyle()
        {
            return CreateChartStyles();
        }

        public CS.ColorStyle GetColorStyle()
        {
            return CreateColorStyles();
        }

        #endregion Public Methods
    }
}