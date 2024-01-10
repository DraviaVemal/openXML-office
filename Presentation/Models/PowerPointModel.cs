/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using OpenXMLOffice.Global;

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

        public ThemePallet Theme = new();

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

        public ThemePallet Theme = new();

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