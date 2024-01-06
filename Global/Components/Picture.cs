using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Global
{
    public class PictureBase : CommonProperties
    {
        private string? _EmbedId;
        protected string? EmbedId
        {
            get
            {
                return _EmbedId;
            }
            set
            {
                _EmbedId = value;
                GetPicture().BlipFill = new P.BlipFill(
                new A.Blip()
                {
                    Embed = value
                },
                new A.Stretch(new A.FillRectangle()));
            }
        }
        private readonly P.Picture OpenXMLPicture;
        private readonly PictureSetting PictureSetting;
        protected PictureBase(PictureSetting PictureSetting)
        {
            this.PictureSetting = PictureSetting;
            OpenXMLPicture = new();
            CreatePicture();
        }
        protected P.Picture GetPicture()
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