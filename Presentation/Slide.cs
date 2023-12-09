using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OpenXMLOffice.Presentation;
internal class Slide
{
    private readonly P.Slide OpenXMLSlide = new();
    private readonly CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE, PresentationConstants.SlideLayoutType.BLANK);
    public Slide()
    {
        OpenXMLSlide.CommonSlideData = commonSlideData.GetCommonSlideData();
        OpenXMLSlide.ColorMapOverride = new P.ColorMapOverride()
        {
            MasterColorMapping = new A.MasterColorMapping()
        };
        OpenXMLSlide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        OpenXMLSlide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    }
    public P.Slide GetSlide()
    {
        return OpenXMLSlide;
    }
}
