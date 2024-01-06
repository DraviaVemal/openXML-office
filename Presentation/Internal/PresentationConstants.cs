/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Presentation
{
    public static class PresentationConstants
    {
        #region Public Enums

        public enum CommonSlideDataType
        {
            SLIDE_MASTER,
            SLIDE_LAYOUT,
            SLIDE
        }

        public enum SlideLayoutType
        {
            BLANK
        }

        #endregion Public Enums

        #region Public Methods

        public static string GetSlideLayoutType(SlideLayoutType value)
        {
            return value switch
            {
                _ => "Blank",
            };
        }

        #endregion Public Methods
    }
}