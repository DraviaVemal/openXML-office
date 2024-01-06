// Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License. See License in
// the project root for license information.
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
        #region Public Fields

        public uint Height = 6858000;
        public ImageType ImageType = ImageType.JPEG;
        public uint Width = 12192000;
        public uint X = 0;
        public uint Y = 0;

        #endregion Public Fields
    }
}