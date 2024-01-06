using OpenXMLOffice.Global;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class Picture : PictureBase
    {
        private readonly Slide CurrentSlide;
        public Picture(Stream Stream, Slide Slide, PictureSetting PictureSetting) : base(PictureSetting)
        {
            CurrentSlide = Slide;
            EmbedId = CurrentSlide.GetNextSlideRelationId();
            ImagePart ImagePart = CurrentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>("image/jpeg", EmbedId);
            ImagePart.FeedData(Stream);
        }

        public Picture(string FilePath, Slide Slide, PictureSetting PictureSetting) : base(PictureSetting)
        {
            CurrentSlide = Slide;
            EmbedId = CurrentSlide.GetNextSlideRelationId();
            ImagePart ImagePart = CurrentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>("image/jpeg", EmbedId);
            ImagePart.FeedData(new FileStream(FilePath, FileMode.Open, FileAccess.Read));

        }
        internal P.Picture GetPicture()
        {
            return base.GetPicture();
        }
    }
}