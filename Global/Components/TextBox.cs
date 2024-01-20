// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents Textbox base class to build on
    /// </summary>
    public class TextBoxBase : CommonProperties
    {
        #region Private Fields

        private readonly TextBoxSetting TextBoxSetting;
        private P.Shape? OpenXMLShape;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Create Textbox with provided settings
        /// </summary>
        /// <param name="TextBoxSetting">
        /// </param>
        public TextBoxBase(TextBoxSetting TextBoxSetting)
        {
            this.TextBoxSetting = TextBoxSetting;
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Get Textbox Run
        /// </summary>
        /// <returns>
        /// </returns>
        public A.Run GetTextBoxBaseRun()
        {
            return CreateTextRun();
        }

        /// <summary>
        /// Get Textbox Shape
        /// </summary>
        /// <returns>
        /// </returns>
        public P.Shape GetTextBoxBaseShape()
        {
            return CreateTextBox();
        }

        /// <summary>
        /// Update Textbox Position
        /// </summary>
        /// <param name="X">
        /// </param>
        /// <param name="Y">
        /// </param>
        public void UpdatePosition(uint X, uint Y)
        {
            TextBoxSetting.X = X;
            TextBoxSetting.Y = Y;
            if (OpenXMLShape != null)
            {
                OpenXMLShape.ShapeProperties!.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = TextBoxSetting.X, Y = TextBoxSetting.Y },
                    Extents = new A.Extents { Cx = TextBoxSetting.Width, Cy = TextBoxSetting.Height }
                };
            }
        }

        /// <summary>
        /// Update Textbox Size
        /// </summary>
        /// <param name="Width">
        /// </param>
        /// <param name="Height">
        /// </param>
        public void UpdateSize(uint Width, uint Height)
        {
            TextBoxSetting.Width = Width;
            TextBoxSetting.Height = Height;
            if (OpenXMLShape != null)
            {
                OpenXMLShape.ShapeProperties!.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = TextBoxSetting.X, Y = TextBoxSetting.Y },
                    Extents = new A.Extents { Cx = TextBoxSetting.Width, Cy = TextBoxSetting.Height }
                };
            }
        }

        #endregion Public Methods

        #region Private Methods

        private P.Shape CreateTextBox()
        {
            OpenXMLShape = new()
            {
                NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = 10,
                    Name = "Text Box"
                },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
                ShapeProperties = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = TextBoxSetting.X, Y = TextBoxSetting.Y },
                    new A.Extents { Cx = TextBoxSetting.Width, Cy = TextBoxSetting.Height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                TextBoxSetting.ShapeBackground != null ? CreateSolidFill(new List<string>() { TextBoxSetting.ShapeBackground }, 0) : new A.NoFill()),
                TextBody = new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(CreateTextRun()))
            };
            return OpenXMLShape;
        }

        private A.Run CreateTextRun()
        {
            A.Run Run = new(new A.RunProperties(CreateSolidFill(new List<string>() { TextBoxSetting.TextColor }, 0),
                        new A.LatinFont { Typeface = TextBoxSetting.FontFamily },
                        new A.EastAsianFont { Typeface = TextBoxSetting.FontFamily },
                        new A.ComplexScriptFont { Typeface = TextBoxSetting.FontFamily })
            {
                FontSize = TextBoxSetting.FontSize * 100,
                Bold = TextBoxSetting.IsBold,
                Italic = TextBoxSetting.IsItalic,
                Underline = TextBoxSetting.IsUnderline ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None,
                Dirty = false
            }, new A.Text(TextBoxSetting.Text));
            if (TextBoxSetting.TextBackground != null)
            {
                Run.Append(new A.Highlight(new A.RgbColorModelHex { Val = TextBoxSetting.TextBackground }));
            }
            return Run;
        }

        #endregion Private Methods
    }
}