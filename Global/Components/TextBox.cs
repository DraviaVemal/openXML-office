// Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License. See License in
// the project root for license information.
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global
{
    public class TextBoxBase : CommonProperties
    {
        #region Private Fields

        private readonly TextBoxSetting TextBoxSetting;
        private P.Shape? OpenXMLShape;

        #endregion Private Fields

        #region Public Constructors

        public TextBoxBase(TextBoxSetting TextBoxSetting)
        {
            this.TextBoxSetting = TextBoxSetting;
        }

        #endregion Public Constructors

        #region Public Methods

        public A.Run GetTextBoxRun()
        {
            return CreateTextRun();
        }

        public P.Shape GetTextBoxShape()
        {
            return CreateTextBox();
        }

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
                CreateSolidFill(new List<string>() { TextBoxSetting.ShapeBackground }, 0)),
                TextBody = new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(CreateTextRun()))
            };
            return OpenXMLShape;
        }

        private A.Run CreateTextRun()
        {
            return new(new A.RunProperties(CreateSolidFill(new List<string>() { TextBoxSetting.TextColor }, 0),
                        new A.Highlight(new A.RgbColorModelHex { Val = TextBoxSetting.TextBackground ?? "FFFFFF" }),
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
        }

        #endregion Private Methods
    }
}