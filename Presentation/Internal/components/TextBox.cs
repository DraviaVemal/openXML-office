using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class TextBox
{
    public string Text = "Text Box";
    public int FontSize = 18;
    public bool IsBold = false;
    public bool IsItalic = false;
    public bool IsUnderline = false;
    public string TextColor = "FFFFFF";
    public string TextBackground = "000000";
    public int X = 0;
    public int Y = 0;
    public int Height = 100;
    public int Width = 100;

    private P.TextBody CreateTextBody(string text)
    {
        return new P.TextBody(new A.BodyProperties(),
                 new A.ListStyle(),
                 new A.Paragraph(new A.Run(new A.Text() { Text = text })));
    }

    public void UpdateTextInShape(Shape RefShape)
    {
        RefShape.GetShape().RemoveAllChildren<P.TextBody>();
        RefShape.GetShape().InsertAt(CreateTextBody(Text), RefShape.GetShape().Count());
    }
}