// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Presentation {
    /// <summary>
    /// Presentation Table Cell Class for setting the cell properties.
    /// </summary>
    public class TableCell {
        #region Public Fields

        /// <summary>
        /// Cell Alignment Option
        /// </summary>
        public AlignmentValues? alignment;

        /// <summary>
        /// Enable Bottom Border
        /// </summary>
        public bool bottomBorder = false;

        /// <summary>
        /// Enable Bottom Left to Top Right Border
        /// </summary>
        public bool bottomLeftToTopRightBorder = false;

        /// <summary>
        /// Cell Background Color
        /// </summary>
        public string? cellBackground;

        /// <summary>
        /// Cell Font Family
        /// Default: Calibri (Body)
        /// </summary>
        public string fontFamily = "Calibri (Body)";

        /// <summary>
        /// Cell Font Size
        /// </summary>
        public int fontSize = 16;

        /// <summary>
        /// Is Bold text
        /// </summary>
        public bool isBold = false;

        /// <summary>
        /// Is Italic text
        /// </summary>
        public bool isItalic = false;

        /// <summary>
        /// Is Underline text
        /// </summary>
        public bool isUnderline = false;

        /// <summary>
        /// Enable Left Border
        /// </summary>
        public bool leftBorder = false;

        /// <summary>
        /// Enable Right Border
        /// </summary>
        public bool rightBorder = false;

        /// <summary>
        /// Text Background Color aka Highlight Color
        /// </summary>
        public string? textBackground;

        /// <summary>
        /// Text Color
        /// </summary>
        public string textColor = "000000";

        /// <summary>
        /// Enable Top Border
        /// </summary>
        public bool topBorder = false;

        /// <summary>
        /// Enable Top Left to Bottom Right Border
        /// </summary>
        public bool topLeftToBottomRightBorder = false;

        /// <summary>
        /// Cell Value
        /// </summary>
        public string? value;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// Cell Vertical Alignment Option
        /// </summary>
        public enum AlignmentValues {
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

        #endregion Public Enums
    }

    /// <summary>
    /// Table Row Customisation Properties
    /// </summary>
    public class TableRow {
        #region Public Fields

        /// <summary>
        /// Row Height
        /// </summary>
        public int height = 370840;

        /// <summary>
        /// Row Background Color.Will get overriden by TableCell.CellBackground
        /// </summary>
        public string? rowBackground;

        /// <summary>
        /// Table Cell List
        /// </summary>
        public List<TableCell> tableCells = new();

        /// <summary>
        /// Default Text Color for the row. Will get overriden by TableCell.TextColor
        /// </summary>
        public string textColor = "000000";

        #endregion Public Fields
    }

    /// <summary>
    /// Table Customisation Properties
    /// </summary>
    public class TableSetting {
        #region Public Fields

        /// <summary>
        /// Overall Table Height
        /// </summary>
        public uint height = 741680;

        /// <summary>
        /// Table Name. Default: Table 1
        /// </summary>
        public string name = "Table 1";

        /// <summary>
        /// Table Column Width List.Works based on WidthType Setting
        /// </summary>
        public List<float> tableColumnWidth = new();

        /// <summary>
        /// Overall Table Width
        /// </summary>
        public uint width = 8128000;

        /// <summary>
        /// AUTO - Ignore User Width value and space the colum equally EMU - (English Metric Units)
        /// Direct PPT standard Sizing 1 Inch * 914400 EMU's PIXEL - Based on Target DPI the pixel
        /// is converted to EMU and used when running PERCENTAGE - 0-100 Width percentage split for
        /// each column RATIO - 0-10 Width ratio of each column
        /// </summary>
        public WidthOptionValues widthType = WidthOptionValues.AUTO;

        /// <summary>
        /// Table X Position in the slide
        /// </summary>
        public uint x = 0;

        /// <summary>
        /// Table Y Position in the slide
        /// </summary>
        public uint y = 0;

        #endregion Public Fields

        #region Public Enums

        /// <summary>
        /// Width Option Values
        /// </summary>
        public enum WidthOptionValues {
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