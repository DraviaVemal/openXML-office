using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    internal class SlideLayout
    {
        #region Private Fields

        private readonly CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE_LAYOUT, PresentationConstants.SlideLayoutType.BLANK);

        private readonly P.SlideLayout OpenXMLSlideLayout = new()
        {
            Type = P.SlideLayoutValues.Text
        };

        #endregion Private Fields

        #region Public Constructors

        public SlideLayout()
        {
            CreateSlideLayout();
        }

        #endregion Public Constructors

        #region Public Methods

        public P.SlideLayout GetSlideLayout()
        {
            return OpenXMLSlideLayout;
        }

        public string UpdateRelationship(OpenXmlPart openXmlPart, string RelationshipId)
        {
            return OpenXMLSlideLayout.SlideLayoutPart!.CreateRelationshipToPart(openXmlPart, RelationshipId);
        }

        #endregion Public Methods

        #region Private Methods

        private void CreateSlideLayout()
        {
            OpenXMLSlideLayout.AppendChild(commonSlideData.GetCommonSlideData());
            OpenXMLSlideLayout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            OpenXMLSlideLayout.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            OpenXMLSlideLayout.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            OpenXMLSlideLayout.AppendChild(new P.ColorMapOverride()
            {
                MasterColorMapping = new A.MasterColorMapping()
            });
        }

        #endregion Private Methods
    }
}