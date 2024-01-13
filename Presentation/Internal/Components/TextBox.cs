/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

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
        /// <returns></returns>
        internal A.Run GetTextBoxRun()
        {
            return base.GetTextBoxBaseRun();
        }
        /// <summary>
        /// Return OpenXML Shape
        /// </summary>
        /// <returns></returns>
        internal P.Shape GetTextBoxShape()
        {
            return base.GetTextBoxBaseShape();
        }

        #endregion Internal Methods
    }
}