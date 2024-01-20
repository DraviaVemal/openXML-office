// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    internal class SlideLayout
    {
        #region Private Fields

        private readonly CommonSlideData commonSlideData = new(PresentationConstants.CommonSlideDataType.SLIDE_LAYOUT, PresentationConstants.SlideLayoutType.BLANK);

        private readonly P.SlideLayout openXMLSlideLayout = new()
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
            return openXMLSlideLayout;
        }

        public string UpdateRelationship(OpenXmlPart openXmlPart, string RelationshipId)
        {
            return openXMLSlideLayout.SlideLayoutPart!.CreateRelationshipToPart(openXmlPart, RelationshipId);
        }

        #endregion Public Methods

        #region Private Methods

        private void CreateSlideLayout()
        {
            openXMLSlideLayout.AppendChild(commonSlideData.GetCommonSlideData());
            openXMLSlideLayout.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            openXMLSlideLayout.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            openXMLSlideLayout.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            openXMLSlideLayout.AppendChild(new P.ColorMapOverride()
            {
                MasterColorMapping = new A.MasterColorMapping()
            });
        }

        #endregion Private Methods
    }
}