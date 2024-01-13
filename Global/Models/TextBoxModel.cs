/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the settings for a text box.
    /// </summary>
    public class TextBoxSetting
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets the font family of the text.
        /// </summary>
        public string FontFamily = "Calibri (Body)";

        /// <summary>
        /// Gets or sets the font size of the text.
        /// </summary>
        public int FontSize = 18;

        /// <summary>
        /// Gets or sets the height of the text box.
        /// </summary>
        public uint Height = 100;

        /// <summary>
        /// Gets or sets a value indicating whether the text is bold.
        /// </summary>
        public bool IsBold = false;

        /// <summary>
        /// Gets or sets a value indicating whether the text is italic.
        /// </summary>
        public bool IsItalic = false;

        /// <summary>
        /// Gets or sets a value indicating whether the text is underlined.
        /// </summary>
        public bool IsUnderline = false;

        /// <summary>
        /// Gets or sets the background color of the text box shape.
        /// </summary>
        public string? ShapeBackground;

        /// <summary>
        /// Gets or sets the text content of the text box.
        /// </summary>
        public string Text = "Text Box";

        /// <summary>
        /// Gets or sets the background color of the text.
        /// </summary>
        public string? TextBackground;

        /// <summary>
        /// Gets or sets the color of the text.
        /// </summary>
        public string TextColor = "000000";

        /// <summary>
        /// Gets or sets the width of the text box.
        /// </summary>
        public uint Width = 100;

        /// <summary>
        /// Gets or sets the X-coordinate of the text box.
        /// </summary>
        public uint X = 0;

        /// <summary>
        /// Gets or sets the Y-coordinate of the text box.
        /// </summary>
        public uint Y = 0;

        #endregion Public Fields
    }
}