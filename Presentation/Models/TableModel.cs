/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Presentation Table Cell Class for setting the cell properties.
    /// </summary>
    public class TableCell
    {
        #region Public Fields

        /// <summary>
        /// Enable Bottom Border
        /// </summary>
        public bool BottomBorder = false;
        /// <summary>
        /// Enable Top Left to Bottom Right Border
        /// </summary>
        public bool TopLeftToBottomRightBorder = false;
        /// <summary>
        /// Enable Bottom Left to Top Right Border
        /// </summary>
        public bool BottomLeftToTopRightBorder = false;

        /// <summary>
        /// Cell Background Color
        /// </summary>
        public string? CellBackground;

        /// <summary>
        /// Cell Font Family
        /// Default: Calibri (Body)
        /// </summary>
        public string FontFamily = "Calibri (Body)";

        /// <summary>
        /// Cell Font Size
        /// </summary>
        public int FontSize = 16;

        /// <summary>
        /// Is Bold text
        /// </summary>
        public bool IsBold = false;

        /// <summary>
        /// Is Italic text
        /// </summary>
        public bool IsItalic = false;

        /// <summary>
        /// Is Underline text
        /// </summary>
        public bool IsUnderline = false;

        /// <summary>
        /// Enable Left Border
        /// </summary>
        public bool LeftBorder = false;

        /// <summary>
        /// Enable Right Border
        /// </summary>
        public bool RightBorder = false;

        /// <summary>
        /// Text Background Color aka Highlight Color
        /// </summary>
        public string? TextBackground;

        /// <summary>
        /// Text Color
        /// </summary>
        public string TextColor = "000000";

        /// <summary>
        /// Enable Top Border
        /// </summary>
        public bool TopBorder = false;

        /// <summary>
        /// Cell Value
        /// </summary>
        public string? Value;

        /// <summary>
        /// Cell Alignment Option
        /// </summary>
        public AlignmentValues? Alignment;

        /// <summary>
        /// Cell Vertical Alignment Option
        /// </summary>
        public enum AlignmentValues
        {
            /// <summary>
            /// Align Left
            /// </summary>
            LEFT,

            /// <summary>
            /// Align Center
            /// </summary>
            CENTER,

            /// <summary>
            /// Align Right
            /// </summary>
            RIGHT,
            /// <summary>
            /// Align Justify
            /// </summary>
            JUSTIFY
        }

        #endregion Public Fields
    }

    /// <summary>
    /// Table Row Customisation Properties
    /// </summary>
    public class TableRow
    {
        #region Public Fields

        /// <summary>
        /// Row Height
        /// </summary>
        public int Height = 370840;

        /// <summary>
        /// Row Background Color.Will get overriden by TableCell.CellBackground
        /// </summary>
        public string? RowBackground;

        /// <summary>
        /// Table Cell List
        /// </summary>
        public List<TableCell> TableCells = new();

        /// <summary>
        /// Default Text Color for the row. Will get overriden by TableCell.TextColor
        /// </summary>
        public string TextColor = "000000";

        #endregion Public Fields
    }

    /// <summary>
    /// Table Customisation Properties
    /// </summary>
    public class TableSetting
    {
        #region Public Fields

        /// <summary>
        /// Overall Table Height
        /// </summary>
        public uint Height = 741680;

        /// <summary>
        /// Table Name. Default: Table 1
        /// </summary>
        public string Name = "Table 1";

        /// <summary>
        /// Table Column Width List.Works based on WidthType Setting
        /// </summary>
        public List<float> TableColumnWidth = new();

        /// <summary>
        /// Overall Table Width
        /// </summary>
        public uint Width = 8128000;

        /// <summary>
        /// AUTO - Ignore User Width value and space the colum equally EMU - (English Metric Units)
        /// Direct PPT standard Sizing 1 Inch * 914400 EMU's PIXEL - Based on Target DPI the pixel
        /// is converted to EMU and used when running PERCENTAGE - 0-100 Width percentage split for
        /// each column RATIO - 0-10 Width ratio of each column
        /// </summary>
        public WidthOptionValues WidthType = WidthOptionValues.AUTO;

        /// <summary>
        /// Table X Position in the slide
        /// </summary>
        public uint X = 0;

        /// <summary>
        /// Table Y Position in the slide
        /// </summary>
        public uint Y = 0;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// Width Option Values
        /// </summary>
        public enum WidthOptionValues
        {
            /// <summary>
            /// AUTO - Ignore User Width value and space the colum equally
            /// </summary>
            AUTO,

            /// <summary>
            /// EMU - (English Metric Units) Direct PPT standard Sizing 1 Inch * 914400 EMU's
            /// </summary>
            EMU,

            /// <summary>
            /// PIXEL - Based on Target DPI the pixel is converted to EMU and used when running
            /// </summary>
            PIXEL,

            /// <summary>
            /// PERCENTAGE - 0-100 Width percentage split for each column
            /// </summary>
            PERCENTAGE,

            /// <summary>
            /// RATIO - 0-10 Width ratio of each column
            /// </summary>
            RATIO
        }

        #endregion Public Enums
    }
}