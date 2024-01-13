/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Excel
{
    public class BorderBase
    {
        #region Public Fields

        public string Color = "64";

        public StyleValues Style = StyleValues.THIN;

        #endregion Public Fields

        #region Public Enums

        public enum StyleValues
        {
            THIN
        }

        #endregion Public Enums
    }

    public class BorderStyle
    {
        #region Public Fields

        public BottomBorder Bottom = new();

        public LeftBorder Left = new();

        public RightBorder Right = new();

        public TopBorder Top = new();

        #endregion Public Fields

        #region Public Properties

        public int Id { get; set; }

        #endregion Public Properties

        #region Public Classes

        public class BottomBorder : BorderBase
        {
        }

        public class LeftBorder : BorderBase
        {
        }

        public class RightBorder : BorderBase
        {
        }

        public class TopBorder : BorderBase
        {
        }

        #endregion Public Classes
    }

    public class CellStyleSetting
    {
        #region Public Fields

        public string? BackgroundColor;

        public string FontFamily = "Calibri";

        public int FontSize = 11;

        public HorizontalAlignmentValues HorizontalAlignment = HorizontalAlignmentValues.LEFT;

        public bool IsBold = false;

        public bool IsDoubleStrick = false;

        public bool IsItalic = false;

        public bool IsStrick = false;

        public bool IsWrapText = false;

        public string NumberFormat = "General";

        public string TextColor = "000000";

        public VerticalAlignmentValues VerticalAlignment = VerticalAlignmentValues.BOTTOM;

        #endregion Public Fields

        #region Public Enums

        public enum HorizontalAlignmentValues
        {
            LEFT,
            CENTER,
            RIGHT
        }

        public enum VerticalAlignmentValues
        {
            TOP,
            MIDDLE,
            BOTTOM
        }

        #endregion Public Enums
    }

    public class FillStyle
    {
        #region Public Properties

        public int Id { get; set; }

        #endregion Public Properties
    }

    public class FontStyle
    {
        #region Public Fields

        public string Color = "accent1";
        public string Family = "2";
        public string name = "Calibri";
        public string Size = "11";

        #endregion Public Fields

        #region Public Properties

        public int Id { get; set; }

        #endregion Public Properties
    }
}