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

        public void MoveSlideByIndex(int SourceIndex, int TargetIndex)
        {
            if (SourceIndex >= 0 && GetSlideIdList().Count() > SourceIndex &&
            TargetIndex >= 0 && GetSlideIdList().Count() > TargetIndex)
            {
                P.SlideId SourceSlideId = (P.SlideId)GetSlideIdList().ElementAt(SourceIndex);
                P.SlideId TargetSlideId = (P.SlideId)GetSlideIdList().ElementAt(TargetIndex);
                GetSlideIdList().RemoveChild(SourceSlideId);
                GetSlideIdList().InsertBefore(SourceSlideId, TargetSlideId);
                presentationDocument.Save();
            }
            else
            {
                throw new IndexOutOfRangeException("The specified slide index is out of range.");
            }
        }

        public void RemoveSlideByIndex(int SlideIndex)
        {
            if (SlideIndex >= 0 && GetSlideIdList().Count() > SlideIndex)
            {
                P.SlideId SlideId = (P.SlideId)GetSlideIdList().ElementAt(SlideIndex);
                SlidePart SlidePart = (SlidePart)GetPresentationPart().GetPartById(SlideId.RelationshipId!.Value!);
                GetSlideIdList().RemoveChild(SlideId);
                GetPresentationPart().DeleteReferenceRelationship(SlideId.RelationshipId.Value!);
                GetPresentationPart().DeletePart(SlidePart);

            }
            else
            {
                throw new IndexOutOfRangeException("The specified slide index is out of range.");
            }
        }

        public Slide GetSlideByIndex(int SlideIndex)
        {
            if (SlideIndex >= 0 && GetSlideIdList().Count() > SlideIndex)
            {
                P.SlideId SlideId = (P.SlideId)GetSlideIdList().ElementAt(SlideIndex);
                SlidePart SlidePart = (SlidePart)GetPresentationPart().GetPartById(SlideId.RelationshipId!.Value!);
                return new Slide(SlidePart.Slide);
            }
            else
            {
                throw new IndexOutOfRangeException("The specified slide index is out of range.");
            }
        }

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