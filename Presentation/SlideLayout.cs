using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class SlideLayout : SlideMaster
{

    protected P.SlideLayout CreateSlideLayout()
    {
        P.SlideLayout slideLayout = new(CreateCommonSlideData());
        slideLayout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        slideLayout.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        slideLayout.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        return slideLayout;
    }

}
