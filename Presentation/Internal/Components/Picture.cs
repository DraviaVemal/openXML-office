/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    public class Picture : CommonProperties
    {
        #region Private Fields

        private readonly Slide CurrentSlide;

        private readonly P.Picture OpenXMLPicture;

        private readonly PictureSetting PictureSetting;

        #endregion Private Fields

        #region Public Constructors

        public Picture(Stream Stream, Slide Slide, PictureSetting PictureSetting)
        {
            CurrentSlide = Slide;
            string EmbedId = CurrentSlide.GetNextSlideRelationId();
            this.PictureSetting = PictureSetting;
            OpenXMLPicture = new();
            CreatePicture(EmbedId);
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
            CreatePicture(EmbedId);
            ImagePart ImagePart = CurrentSlide.GetSlide().SlidePart!.AddNewPart<ImagePart>(PictureSetting.ImageType switch
            {
                ImageType.PNG => "image/png",
                ImageType.GIF => "image/gif",
                ImageType.TIFF => "image/tiff",
                _ => "image/jpeg"
            }, EmbedId);
            ImagePart.FeedData(new FileStream(FilePath, FileMode.Open, FileAccess.Read));
        }

        #endregion Public Constructors

        #region Public Methods

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

        #endregion Public Methods

        #region Internal Methods

        internal P.Picture GetPicture()
        {
            return OpenXMLPicture;
        }

        #endregion Internal Methods

        #region Private Methods

        private void CreatePicture(string EmbedId)
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
            GetPicture().BlipFill = new()
            {
                Blip = new A.Blip() { Embed = EmbedId }
            };
            GetPicture().BlipFill!.Append(new A.Stretch(new A.FillRectangle()));
        }

        #endregion Private Methods
    }
}