/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the Data type of the chart data.
    /// </summary>
    public enum DataType
    {
        /// <summary>
        /// Date Data Type
        /// </summary>
        DATE,
        /// <summary>
        /// Number Data Type
        /// </summary>
        NUMBER,
        /// <summary>
        /// String Data Type
        /// </summary>
        STRING
    }
    /// <summary>
    /// Represents the settings for a chart data.
    /// </summary>
    public class ChartData
    {
        #region Public Fields
        /// <summary>
        /// The data type of the chart data.
        /// </summary>
        public DataType DataType = DataType.STRING;
        /// <summary>
        /// The value of the chart data.
        /// </summary>
        public string? Value;
        /// <summary>
        /// Number Format for Chart Data (Default: General)
        /// </summary>
        public string NumberFormat = "General";

        #endregion Public Fields
    }
}