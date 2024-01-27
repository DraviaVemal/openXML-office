// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global
{
    /// <summary>
    /// Represents the type of image.
    /// </summary>
    public enum ImageType
    {
        /// <summary>
        /// JPEG image Format.
        /// </summary>
        JPEG,

        /// <summary>
        /// PNG image Format.
        /// </summary>
        PNG,

        /// <summary>
        /// GIF image Format.
        /// </summary>
        GIF,

        /// <summary>
        /// BMP image Format.
        /// </summary>
        BMP,

        /// <summary>
        /// TIFF image Format.
        /// </summary>
        TIFF
    }

    /// <summary>
    /// Represents the settings for a picture.
    /// </summary>
    public class PictureSetting
    {        /// <summary>
             /// The height of the picture in EMUs (English Metric Units).
             /// </summary>
        public uint height = 6858000;

        /// <summary>
        /// The type of image.
        /// </summary>
        public ImageType imageType = ImageType.JPEG;

        /// <summary>
        /// The width of the picture in EMUs (English Metric Units).
        /// </summary>
        public uint width = 12192000;

        /// <summary>
        /// The X coordinate of the picture in EMUs (English Metric Units).
        /// </summary>
        public uint x = 0;

        /// <summary>
        /// The Y coordinate of the picture in EMUs (English Metric Units).
        /// </summary>
        public uint y = 0;
    }
}