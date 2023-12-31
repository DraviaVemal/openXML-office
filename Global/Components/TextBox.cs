using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global
{
    public class TextBox
    {
        #region Public Fields

        public int Height = 100;
        public int Width = 100;
        public int X = 0;
        public int Y = 0;
        private P.Shape? OpenXMLShape;

        #endregion Public Fields

        #region Public Methods

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

        public A.Run CreateTextRun(TextBoxSetting TextBoxSetting)
        {
            return new(new A.RunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = TextBoxSetting.TextColor }),
                        new A.Highlight(new A.RgbColorModelHex { Val = TextBoxSetting.TextBackground }),
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

        public P.Shape CreateTextBox(uint Id, TextBoxSetting TextBoxSetting)
        {
            OpenXMLShape = new()
            {
                NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = Id,
                    Name = "Text Box"
                },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
                ShapeProperties = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = X, Y = Y },
                    new A.Extents { Cx = Width, Cy = Height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.SolidFill(new A.RgbColorModelHex { Val = TextBoxSetting.ShapeBackground })),
                TextBody = new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(CreateTextRun(TextBoxSetting)))
            };
            return OpenXMLShape;
        }

        #endregion Public Methods
    }
}