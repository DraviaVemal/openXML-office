using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class SlideMaster : Theme
{
    protected SlideLayoutIdList? slideLayoutIdList;
    protected P.SlideMaster CreateSlideMaster()
    {
        P.SlideMaster slideMaster = new(CreateCommonSlideData());
        slideMaster.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        slideMaster.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        slideMaster.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        slideMaster.AppendChild(new ColorMap()
        {
            Accent1 = A.ColorSchemeIndexValues.Accent1,
            Accent2 = A.ColorSchemeIndexValues.Accent2,
            Accent3 = A.ColorSchemeIndexValues.Accent3,
            Accent4 = A.ColorSchemeIndexValues.Accent4,
            Accent5 = A.ColorSchemeIndexValues.Accent5,
            Accent6 = A.ColorSchemeIndexValues.Accent6,
            Background1 = A.ColorSchemeIndexValues.Light1,
            Text1 = A.ColorSchemeIndexValues.Dark1,
            Background2 = A.ColorSchemeIndexValues.Light2,
            Text2 = A.ColorSchemeIndexValues.Dark2,
            Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
        });
        slideLayoutIdList = new();
        slideMaster.AppendChild(slideLayoutIdList);
        return slideMaster;
    }

    protected void AddSlideLayoutIdToList(string relationshipId)
    {
        slideLayoutIdList!.AppendChild(new SlideLayoutId()
        {
            Id = (uint)(2147483649 + slideLayoutIdList.Count() + 1),
            RelationshipId = relationshipId
        });
    }
}
