using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    internal class Presentation : PresentationCore
    {
        public Presentation(string filePath, bool isEditable, PresentationProperties? presentationProperties = null, bool autosave = true)
        : base(filePath, isEditable, presentationProperties, autosave) { }
        public Presentation(string filePath, PresentationProperties? presentationProperties = null, PresentationDocumentType presentationDocumentType = PresentationDocumentType.Presentation, bool autosave = true)
        : base(filePath, presentationProperties, presentationDocumentType, autosave) { }

        public Presentation(Stream stream, PresentationProperties? presentationProperties = null, PresentationDocumentType presentationDocumentType = PresentationDocumentType.Presentation, bool autosave = true)
        : base(stream, presentationProperties, presentationDocumentType) { }

        public Slide AddSlide(PresentationConstants.SlideLayoutType slideLayoutType)
        {
            SlidePart slidePart = GetPresentationPart().AddNewPart<SlidePart>(GetNextPresentationRelationId());
            Slide slide = new();
            slidePart.Slide = slide.GetSlide();
            slidePart.AddPart(GetSlideLayoutPart(slideLayoutType));
            P.SlideIdList slideIdList = GetSlideIdList();
            P.SlideId slideId = new() { Id = GetNextSlideId(), RelationshipId = GetPresentationPart().GetIdOfPart(slidePart) };
            slideIdList.Append(slideId);
            return slide;
        }

        // public Slide GetSlideByIndex(int SlideIndex)
        // {

        // }

        public void Save()
        {
            if (presentationInfo.FilePath == null)
            {
                throw new FieldAccessException("Data Is in File Stream Use SaveAs to Target save file");
            }
            if (presentationInfo.IsEditable)
            {
                presentationDocument.Clone(presentationInfo.FilePath).Dispose();
            }
            presentationDocument.Dispose();
        }

        public void SaveAs(string filePath)
        {
            presentationDocument.Clone(filePath).Dispose();
            presentationDocument.Dispose();
        }
    }
}