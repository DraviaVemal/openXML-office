/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Presentation
{
    public class PresentationProperties
    {
        #region Public Fields

        public PresentationSettings Settings = new();

        /// <summary>
        /// TODO : Multi Theme Slide Master Support
        /// </summary>
        public Dictionary<string, PresentationSlideMaster>? SlideMasters;

        public PresentationTheme Theme = new();

        #endregion Public Fields
    }

    public class PresentationSettings
    {
        #region Public Fields

        public bool IsMultiSlideMasterPartPresentation = false;
        public bool IsMultiThemePresentation = false;

        #endregion Public Fields
    }

    /// <summary>
    /// TODO : Multi Theme Slide Master Support
    /// </summary>
    public class PresentationSlideMaster
    {
        #region Public Fields

        public PresentationTheme Theme = new();

        #endregion Public Fields
    }

    public class PresentationTheme
    {
        #region Public Fields

        public string Accent1 = "4472C4";
        public string Accent2 = "ED7D31";
        public string Accent3 = "A5A5A5";
        public string Accent4 = "FFC000";
        public string Accent5 = "5B9BD5";
        public string Accent6 = "70AD47";
        public string Dark1 = "000000";
        public string Dark2 = "44546A";
        public string FollowedHyperlink = "954F72";
        public string Hyperlink = "0563C1";
        public string Light1 = "FFFFFF";
        public string Light2 = "E7E6E6";

        #endregion Public Fields
    }

    internal class PresentationInfo
    {
        #region Public Fields

        public string? FilePath;
        public bool IsEditable = true;
        public bool IsExistingFile = false;

        #endregion Public Fields
    }
}