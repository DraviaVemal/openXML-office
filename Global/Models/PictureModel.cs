
namespace OpenXMLOffice.Global
{
    public enum ImageType
    {
        JPEG,
        PNG,
        GIF,
        BMP,
        TIFF
    }
    public class PictureSetting
    {
        public ImageType ImageType = ImageType.JPEG;
        public uint X = 0;
        public uint Y = 0;
        public uint Height = 6858000;
        public uint Width = 12192000;
    }
}