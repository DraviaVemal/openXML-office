// Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License. See License in
// the project root for license information.
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global
{
    public class ColumnChart : ColumnFamilyChart
    {
        #region Public Constructors

        public ColumnChart(ColumnChartSetting ColumnChartSetting, ChartData[][] DataCols) : base(ColumnChartSetting, DataCols)
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