// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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
        public PresentationSettings settings = new();

        /// <summary>
        /// Gets or sets the slide masters of the presentation.
        /// </summary>
        /// <remarks>
        /// TODO: Multi Theme Slide Master Support
        /// </remarks>
        public Dictionary<string, PresentationSlideMaster>? slideMasters;

        /// <summary>
        /// Gets or sets the theme of the presentation.
        /// </summary>
        public ThemePallet theme = new();

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
        public bool isMultiSlideMasterPartPresentation = false;

        /// <summary>
        /// Gets or sets a value indicating whether the presentation has multiple themes.
        /// </summary>
        public bool isMultiThemePresentation = false;

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
        public ThemePallet theme = new();

        #endregion Public Fields
    }

    internal class PresentationInfo
    {
        #region Public Fields

        public string? filePath;
        public bool isEditable = true;
        public bool isExistingFile = false;

        #endregion Public Fields
    }
}