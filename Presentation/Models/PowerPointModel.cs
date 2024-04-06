// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Represents the properties of a presentation.
	/// </summary>
	public class PresentationProperties
	{
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
	}

	/// <summary>
	/// Represents the settings of a presentation.
	/// </summary>
	public class PresentationSettings
	{
		/// <summary>
		/// Gets or sets a value indicating whether the presentation has multiple slide masters.
		/// </summary>
		public bool isMultiSlideMasterPartPresentation = false;

		/// <summary>
		/// Gets or sets a value indicating whether the presentation has multiple themes.
		/// </summary>
		public bool isMultiThemePresentation = false;
	}

	/// <summary>
	/// Represents a slide master in a presentation.
	/// </summary>
	public class PresentationSlideMaster
	{
		/// <summary>
		/// Gets or sets the theme of the slide master.
		/// </summary>
		public ThemePallet theme = new();
	}

	internal class PresentationInfo
	{
		public string? filePath;
		public bool isEditable = true;
		public bool isExistingFile = false;
	}
}
