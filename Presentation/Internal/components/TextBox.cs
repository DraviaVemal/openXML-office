using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class TextBox
    {
        public string Text = "Text Box";
        public int FontSize = 18;
        public bool IsBold = false;
        public bool IsItalic = false;
        public bool IsUnderline = false;
        public string TextColor = "000000";
        public string TextBackground = "FFFFFF";
        public string ShapeBackground = "FFFFFF";
        public string FontFamily = "Calibri (Body)";
        public int X = 0;
        public int Y = 0;
        public int Height = 100;
        public int Width = 100;

        public P.Shape CreateTextBox(uint Id = 100, string Name = "Text Box")
        {
            A.RunProperties runProperties = new(new A.SolidFill(new A.RgbColorModelHex() { Val = TextColor ?? "000000" }),
             new A.Highlight(new A.RgbColorModelHex() { Val = TextBackground }),
             new A.LatinFont() { Typeface = FontFamily },
             new A.EastAsianFont() { Typeface = FontFamily },
             new A.ComplexScriptFont() { Typeface = FontFamily })
            {
                FontSize = FontSize * 100,
                Bold = IsBold,
                Italic = IsItalic,
                Underline = IsUnderline ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None,
                Dirty = false
            };
            return new()
            {
                NonVisualShapeProperties = new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = Id,
                    Name = Name
                },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
                ShapeProperties = new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = X, Y = Y },
                    new A.Extents() { Cx = Width, Cy = Height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.SolidFill(new A.RgbColorModelHex() { Val = ShapeBackground })),
                TextBody = new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(runProperties, new A.Text(Text))))
            };
        }
    }
}