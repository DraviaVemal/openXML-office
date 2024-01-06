using OpenXMLOffice.Global;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class Picture : CommonProperties
    {
        private readonly Slide CurrentSlide;
        public Picture(Stream Stream, Slide Slide, PictureSetting PictureSetting)
        {
            CurrentSlide = Slide;
            string EmbedId = CurrentSlide.GetNextSlideRelationId();
            this.PictureSetting = PictureSetting;
            OpenXMLPicture = new();
            CreatePicture();
            ImagePart ImagePart = CurrentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>(PictureSetting.ImageType switch
            {
                ImageType.PNG => "image/png",
                ImageType.GIF => "image/gif",
                ImageType.TIFF => "image/tiff",
                _ => "image/jpeg"
            }, EmbedId);
            ImagePart.FeedData(Stream);
        }

        public Picture(string FilePath, Slide Slide, PictureSetting PictureSetting)
        {
            CurrentSlide = Slide;
            string EmbedId = CurrentSlide.GetNextSlideRelationId();
            this.PictureSetting = PictureSetting;
            OpenXMLPicture = new();
            CreatePicture();
            ImagePart ImagePart = CurrentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>(PictureSetting.ImageType switch
            {
                ImageType.PNG => "image/png",
                ImageType.GIF => "image/gif",
                ImageType.TIFF => "image/tiff",
                _ => "image/jpeg"
            }, EmbedId);
            ImagePart.FeedData(new FileStream(FilePath, FileMode.Open, FileAccess.Read));

        }
        private readonly P.Picture OpenXMLPicture;
        private readonly PictureSetting PictureSetting;
        internal P.Picture GetPicture()
        {
            return OpenXMLPicture;
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// X,Y
        /// </returns>
        public (uint, uint) GetPosition()
        {
            return (PictureSetting.X, PictureSetting.Y);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// Width,Height
        /// </returns>
        public (uint, uint) GetSize()
        {
            return (PictureSetting.Width, PictureSetting.Height);
        }

        public void UpdatePosition(uint X, uint Y)
        {
            PictureSetting.X = X;
            PictureSetting.Y = Y;
            if (OpenXMLPicture != null)
            {
                OpenXMLPicture.ShapeProperties!.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = PictureSetting.X, Y = PictureSetting.Y },
                    Extents = new A.Extents { Cx = PictureSetting.Width, Cy = PictureSetting.Height }
                };
            }
        }

        public void UpdateSize(uint Width, uint Height)
        {
            PictureSetting.Width = Width;
            PictureSetting.Height = Height;
            if (OpenXMLPicture != null)
            {
                OpenXMLPicture.ShapeProperties!.Transform2D = new A.Transform2D
                {
                    Offset = new A.Offset { X = PictureSetting.X, Y = PictureSetting.Y },
                    Extents = new A.Extents { Cx = PictureSetting.Width, Cy = PictureSetting.Height }
                };
            }
        }
        private void CreatePicture()
        {
            GetPicture().NonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = 1,
                    Name = "Picture"
                },
                new P.NonVisualPictureDrawingProperties(
                    new A.PictureLocks()
                    {
                        NoChangeAspect = true
                    }
                ),
                new P.ApplicationNonVisualDrawingProperties()
            );
            GetPicture().ShapeProperties = new P.ShapeProperties(
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            )
            {
                Transform2D = new A.Transform2D()
                {
                    Offset = new A.Offset() { X = PictureSetting.X, Y = PictureSetting.Y },
                    Extents = new A.Extents() { Cx = PictureSetting.Width, Cy = PictureSetting.Height }
                }
            };
        }
    }
}