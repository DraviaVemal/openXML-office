using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global
{
    public class TextBox : CommonProperties
    {
        #region Public Fields

        private int Height = 100;
        private int Width = 100;
        private int X = 0;
        private int Y = 0;
        private readonly TextBoxSetting TextBoxSetting;

        #endregion Public Fields

        #region Private Fields

        private P.Shape? OpenXMLShape;

        #endregion Private Fields

        public TextBox(TextBoxSetting TextBoxSetting)
        {
            this.TextBoxSetting = TextBoxSetting;
        }

        #region Public Methods

        public P.Shape GetTextBoxShape()
        {
            return CreateTextBox();
        }

        public A.Run GetTextBoxRun()
        {
            return CreateTextRun();
        }

        public void UpdatePosition(int X, int Y)
        {
            this.X = X;
            this.Y = Y;
            if (OpenXMLShape != null)
            {
                OpenXMLShape.ShapeProperties!.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        public void UpdateSize(int Width, int Height)
        {
            this.Width = Width;
            this.Height = Height;
            if (OpenXMLShape != null)
            {
                OpenXMLShape.ShapeProperties!.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
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
                    new A.Offset { X = X, Y = Y },
                    new A.Extents { Cx = Width, Cy = Height }),
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