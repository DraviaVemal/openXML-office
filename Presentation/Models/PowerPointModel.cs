/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using OpenXMLOffice.Global;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Represents the properties of a presentation.
    /// </summary>
    public class PresentationProperties
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets the presentation settings.
        /// </summary>
        public PresentationSettings Settings = new();

        /// <summary>
        /// Gets or sets the slide masters of the presentation.
        /// </summary>
        /// <remarks>
        /// TODO: Multi Theme Slide Master Support
        /// </remarks>
        public Dictionary<string, PresentationSlideMaster>? SlideMasters;

        /// <summary>
        /// Gets or sets the theme of the presentation.
        /// </summary>
        public ThemePallet Theme = new();

        #endregion Public Fields
    }

    /// <summary>
    /// Represents the settings of a presentation.
    /// </summary>
    public class PresentationSettings
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets a value indicating whether the presentation has multiple slide masters.
        /// </summary>
        public bool IsMultiSlideMasterPartPresentation = false;

        /// <summary>
        /// Gets or sets a value indicating whether the presentation has multiple themes.
        /// </summary>
        public bool IsMultiThemePresentation = false;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents a slide master in a presentation.
    /// </summary>
    public class PresentationSlideMaster
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets the theme of the slide master.
        /// </summary>
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