using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLOffice.Presentation;
public class Slide
{
    private readonly P.Slide OpenXMLSlide = new();
    internal Slide(P.Slide? OpenXMLSlide = null)
    {
        if (OpenXMLSlide != null)
        {
            this.OpenXMLSlide = OpenXMLSlide;
        }
        else
        {
            CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
            this.OpenXMLSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
            this.OpenXMLSlide.ColorMapOverride = new P.ColorMapOverride()
            {
                MasterColorMapping = new A.MasterColorMapping()
            };
            this.OpenXMLSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            this.OpenXMLSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        }
    }
    private P.CommonSlideData GetCommonSlideData()
    {
        return OpenXMLSlide.CommonSlideData!;
    }
    internal SlidePart GetSlidePart()
    {
        return OpenXMLSlide.SlidePart!;
    }
    internal string GetNextSlideRelationId()
    {
        return string.Format("rId{0}", GetSlidePart().Parts.Count() + 1);
    }
    public IEnumerable<Shape> FindShapeByText(string searchText)
    {
        IEnumerable<P.Shape> searchResults = GetCommonSlideData().ShapeTree!.Elements<P.Shape>().Where(shape =>
        {
            return shape.InnerText == searchText;
        });
        return searchResults.Select(shape =>
        {
            return new Shape(shape);
        });
    }
    internal P.Slide GetSlide()
    {
        return OpenXMLSlide;
    }
}
