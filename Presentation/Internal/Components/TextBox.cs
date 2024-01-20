// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Textbox Class
    /// </summary>
    public class TextBox : TextBoxBase
    {
        #region Public Constructors

        /// <summary>
        /// Create Textbox with provided settings
        /// </summary>
        /// <param name="TextBoxSetting">
        /// </param>
        public TextBox(TextBoxSetting TextBoxSetting) : base(TextBoxSetting) { }

        #endregion Public Constructors

        #region Internal Methods

        /// <summary>
        /// Return OpenXML Run
        /// </summary>
        /// <returns>
        /// </returns>
        internal A.Run GetTextBoxRun()
        {
            return GetTextBoxBaseRun();
        }

        /// <summary>
        /// Return OpenXML Shape
        /// </summary>
        /// <returns>
        /// </returns>
        internal P.Shape GetTextBoxShape()
        {
            return GetTextBoxBaseShape();
        }

        #endregion Internal Methods
    }
}