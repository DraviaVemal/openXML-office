/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// Represents the base class for a border in a style.
    /// </summary>
    public class BorderBase
    {
        #region Public Fields
        /// <summary>
        /// Gets or sets the color of the border.
        /// </summary>
        public string Color = "64";
        /// <summary>
        /// Gets or sets the style of the border.
        /// </summary>
        public StyleValues Style = StyleValues.THIN;

        #endregion Public Fields

        #region Public Enums
        /// <summary>
        /// Border style values
        /// </summary>
        public enum StyleValues
        {
            /// <summary>
            /// Thin Border option
            /// </summary>
            THIN
        }

        #endregion Public Enums
    }
    /// <summary>
    /// Represents the border style of a cell in a worksheet.
    /// </summary>
    public class BorderStyle
    {
        #region Public Fields
        /// <summary>
        /// Bottom border style
        /// </summary>
        public BottomBorder Bottom = new();
        /// <summary>
        /// Left border style
        /// </summary>
        public LeftBorder Left = new();
        /// <summary>
        /// Right border style
        /// </summary>
        public RightBorder Right = new();
        /// <summary>
        /// Top border style
        /// </summary>
        public TopBorder Top = new();

        #endregion Public Fields

        #region Public Properties
        /// <summary>
        /// Gets or sets the ID of the border style.
        /// </summary>
        public int Id { get; set; }

        #endregion Public Properties

        #region Public Classes
        /// <summary>
        /// Represents the bottom border style of a cell in a worksheet.
        /// </summary>
        public class BottomBorder : BorderBase
        {
        }
        /// <summary>
        /// Represents the left border style of a cell in a worksheet.
        /// </summary>
        public class LeftBorder : BorderBase
        {
        }
        /// <summary>
        /// Represents the right border style of a cell in a worksheet.
        /// </summary>
        public class RightBorder : BorderBase
        {
        }
        /// <summary>
        /// Represents the top border style of a cell in a worksheet.
        /// </summary>
        public class TopBorder : BorderBase
        {
        }

        #endregion Public Classes
    }
    /// <summary>
    /// Represents the fill style of a cell in a worksheet.
    /// </summary>
    public class CellStyleSetting
    {
        #region Public Fields
        /// <summary>
        /// Gets or sets the background color of the cell.
        /// </summary>
        public string? BackgroundColor;
        /// <summary>
        /// Gets or sets the font family of the cell. default is Calibri
        /// </summary>
        public string FontFamily = "Calibri";
        /// <summary>
        /// Gets or sets the font size of the cell. default is 11
        /// </summary>
        public int FontSize = 11;
        /// <summary>
        /// Horizontal alignment of the cell. default is left
        /// </summary>
        public HorizontalAlignmentValues HorizontalAlignment = HorizontalAlignmentValues.LEFT;
        /// <summary>
        /// Is Cell Bold. default is false
        /// </summary>
        public bool IsBold = false;
        /// <summary>
        /// Is Cell Double Strick. default is false
        /// </summary>
        public bool IsDoubleStrick = false;
        /// <summary>
        /// Is Cell Italic. default is false
        /// </summary>
        public bool IsItalic = false;
        /// <summary>
        /// Is Cell Strick. default is false
        /// </summary>
        public bool IsStrick = false;
        /// <summary>
        /// Is Wrap Text. default is false
        /// </summary>
        public bool IsWrapText = false;
        /// <summary>
        /// Gets or sets the number format of the cell. default is General
        /// </summary>
        public string NumberFormat = "General";
        /// <summary>
        /// Gets or sets the text color of the cell. default is 000000
        /// </summary>
        public string TextColor = "000000";
        /// <summary>
        /// Vertical alignment of the cell. default is bottom
        /// </summary>
        public VerticalAlignmentValues VerticalAlignment = VerticalAlignmentValues.BOTTOM;

        #endregion Public Fields

        #region Public Enums
        /// <summary>
        /// Horizontal alignment values
        /// </summary>
        public enum HorizontalAlignmentValues
        {
            /// <summary>
            /// Left alignment
            /// </summary>
            LEFT,
            /// <summary>
            /// Center alignment
            /// </summary>
            CENTER,
            /// <summary>
            /// Right alignment
            /// </summary>
            RIGHT
        }
        /// <summary>
        /// Vertical alignment values
        /// </summary>
        public enum VerticalAlignmentValues
        {
            /// <summary>
            /// Top alignment
            /// </summary>
            TOP,
            /// <summary>
            /// Middle alignment
            /// </summary>
            MIDDLE,
            /// <summary>
            /// Bottom alignment
            /// </summary>
            BOTTOM
        }

        #endregion Public Enums
    }
    /// <summary>
    /// Represents the fill style of a cell in a worksheet.
    /// </summary>
    public class FillStyle
    {
        #region Public Properties
        /// <summary>
        /// Fill style ID
        /// </summary>
        public int Id { get; set; }

        #endregion Public Properties
    }
    /// <summary>
    /// Represents the font style of a cell in a worksheet.
    /// </summary>
    public class FontStyle
    {
        #region Public Fields
        /// <summary>
        /// Gets or sets the color of the font. default is accent1
        /// </summary>
        public string Color = "accent1";
        /// <summary>
        /// Gets or sets the font family of the font.
        /// </summary>
        public string Family = "2";
        /// <summary>
        /// Font name default is Calibri
        /// </summary>
        public string name = "Calibri";
        /// <summary>
        /// Gets or sets the size of the font. default is 11
        /// </summary>
        public string Size = "11";

        #endregion Public Fields

        #region Public Properties
        /// <summary>
        /// Font style ID
        /// </summary>
        public int Id { get; set; }

        #endregion Public Properties
    }
}