// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// Represents the properties of a column in a worksheet.
    /// </summary>
    public class SpreadsheetProperties
    {
        #region Public Fields

        /// <summary>
        /// Spreadsheet settings
        /// </summary>
        public SpreadsheetSettings Settings = new();

        /// <summary>
        /// Spreadsheet theme settings
        /// </summary>
        public ThemePallet Theme = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings of a spreadsheet.
    /// </summary>
    public class SpreadsheetSettings
    {
    }
}