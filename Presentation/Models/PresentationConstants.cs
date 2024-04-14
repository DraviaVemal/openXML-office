// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Contains constants related to the presentation functionality.
	/// </summary>
	public static class PresentationConstants
	{
		/// <summary>
		/// Represents the common slide data types.
		/// </summary>
		public enum CommonSlideDataType
		{
			/// <summary>
			/// Generate Common Slide Data for a Slide Master Specification
			/// </summary>
			SLIDE_MASTER,
			/// <summary>
			/// Generate Common Slide Data for a Slide Layout Specification
			/// </summary>
			SLIDE_LAYOUT,
			/// <summary>
			/// Generate Common Slide Data for a Slide Specification
			/// </summary>
			SLIDE
		}
		/// <summary>
		/// Represents the slide layout types.
		/// </summary>
		public enum SlideLayoutType
		{
			/// <summary>
			/// Slide Layout Blank option
			/// </summary>
			BLANK
		}
		/// <summary>
		/// Gets the string representation of the specified slide layout type.
		/// </summary>
		/// <param name="value">
		/// The slide layout type.
		/// </param>
		/// <returns>
		/// The string representation of the slide layout type.
		/// </returns>
		public static string GetSlideLayoutType(SlideLayoutType value)
		{
			return value switch
			{
				_ => "Blank",
			};
		}
	}
}
