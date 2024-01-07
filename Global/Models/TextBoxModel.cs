/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    public class TextBoxSetting
    {
        #region Public Fields

        public string FontFamily = "Calibri (Body)";
        public int FontSize = 18;
        public uint Height = 100;
        public bool IsBold = false;
        public bool IsItalic = false;
        public bool IsUnderline = false;
        public string? ShapeBackground;
        public string Text = "Text Box";
        public string? TextBackground;
        public string TextColor = "000000";
        public uint Width = 100;
        public uint X = 0;
        public uint Y = 0;

        #endregion Public Fields
    }
}