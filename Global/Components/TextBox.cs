// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global {
    /// <summary>
    /// Represents Textbox base class to build on
    /// </summary>
    public class TextBoxBase : CommonProperties {
        #region Private Fields

        private readonly TextBoxSetting textBoxSetting;
        private P.Shape? openXMLShape;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Create Textbox with provided settings
        /// </summary>
        /// <param name="TextBoxSetting">
        /// </param>
        public TextBoxBase(TextBoxSetting TextBoxSetting) {
            this.textBoxSetting = TextBoxSetting;
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Get Textbox Run
        /// </summary>
        /// <returns>
        /// </returns>
        public A.Run GetTextBoxBaseRun() {
            return CreateTextRun();
        }

        /// <summary>
        /// Get Textbox Shape
        /// </summary>
        /// <returns>
        /// </returns>
        public P.Shape GetTextBoxBaseShape() {
            return CreateTextBox();
        }

        /// <summary>
        /// Update Textbox Position
        /// </summary>
        /// <param name="X">
        /// </param>
        /// <param name="Y">
        /// </param>
        public void UpdatePosition(uint X,uint Y) {
            textBoxSetting.x = X;
            textBoxSetting.y = Y;
            if(openXMLShape != null) {
                openXMLShape.ShapeProperties!.Transform2D = new A.Transform2D {
                    Offset = new A.Offset { X = textBoxSetting.x,Y = textBoxSetting.y },
                    Extents = new A.Extents { Cx = textBoxSetting.width,Cy = textBoxSetting.height }
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
        public void UpdateSize(uint Width,uint Height) {
            textBoxSetting.width = Width;
            textBoxSetting.height = Height;
            if(openXMLShape != null) {
                openXMLShape.ShapeProperties!.Transform2D = new A.Transform2D {
                    Offset = new A.Offset { X = textBoxSetting.x,Y = textBoxSetting.y },
                    Extents = new A.Extents { Cx = textBoxSetting.width,Cy = textBoxSetting.height }
                };
            }
        }

        #endregion Public Methods

        #region Private Methods

        private P.Shape CreateTextBox() {
            openXMLShape = new() {
                NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() {
                    Id = 10,
                    Name = "Text Box"
                },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
                ShapeProperties = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = textBoxSetting.x,Y = textBoxSetting.y },
                    new A.Extents { Cx = textBoxSetting.width,Cy = textBoxSetting.height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                textBoxSetting.shapeBackground != null ? CreateSolidFill(new List<string>() { textBoxSetting.shapeBackground },0) : new A.NoFill()),
                TextBody = new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(CreateTextRun()))
            };
            return openXMLShape;
        }

        private A.Run CreateTextRun() {
            A.Run Run = new(new A.RunProperties(CreateSolidFill(new List<string>() { textBoxSetting.textColor },0),
                        new A.LatinFont { Typeface = textBoxSetting.fontFamily },
                        new A.EastAsianFont { Typeface = textBoxSetting.fontFamily },
                        new A.ComplexScriptFont { Typeface = textBoxSetting.fontFamily }) {
                FontSize = textBoxSetting.fontSize * 100,
                Bold = textBoxSetting.isBold,
                Italic = textBoxSetting.isItalic,
                Underline = textBoxSetting.isUnderline ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None,
                Dirty = false
            },new A.Text(textBoxSetting.text));
            if(textBoxSetting.textBackground != null) {
                Run.Append(new A.Highlight(new A.RgbColorModelHex { Val = textBoxSetting.textBackground }));
            }
            return Run;
        }

        #endregion Private Methods
    }
}