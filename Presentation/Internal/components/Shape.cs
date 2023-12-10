using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class Shape
{
    private P.Shape OpenXMLShape = new();
    internal Shape(P.Shape? shape = null)
    {
        if (shape != null)
        {
            OpenXMLShape = shape;
        }
    }
    public void RemoveShape()
    {
        OpenXMLShape.Remove();
    }
    internal P.Shape GetShape()
    {
        return OpenXMLShape;
    }
}