// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Text Horizontal Alignment
	/// </summary>
	public enum HorizontalAlignmentValues
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		/// Align Left
		/// </summary>
		LEFT,
		/// <summary>
		/// Align Center
		/// </summary>
		CENTER,
		/// <summary>
		/// Align Right
		/// </summary>
		RIGHT,
		/// <summary>
		/// Align Justify
		/// </summary>
		JUSTIFY
	}
	/// <summary>
	/// Text Vertical Alignment
	/// </summary>
	public enum VerticalAlignmentValues
	{
		/// <summary>
		///
		/// </summary>
		NONE,
		/// <summary>
		/// Align Top
		/// </summary>
		TOP,
		/// <summary>
		/// Align Middle
		/// </summary>
		MIDDLE,
		/// <summary>
		/// Align Bottom
		/// </summary>
		BOTTOM
	}
	/// <summary>
	///
	/// </summary>
	public class TextBlock : TextOptions
	{
		/// <summary>
		///
		/// </summary>
		public HyperlinkProperties hyperlinkProperties;
		/// <summary>
		/// Gets or sets a value indicating whether the text is underlined.
		/// </summary>
		public bool isUnderline;
		/// <summary>
		/// Gets or sets the background color of the text.
		/// </summary>
		public string textBackground;
		/// <summary>
		/// Gets or sets the color of the text.
		/// </summary>
		public string textColor = "000000";
	}

	/// <summary>
	/// Represents the settings for a text box.
	/// </summary>
	public class TextBoxSetting
	{
		/// <summary>
		/// Define Each section of string and its property that goes in same Text box
		/// </summary>
		public TextBlock[] textBlocks;
		/// <summary>
		/// Gets or sets the background color of the text box shape.
		/// </summary>
		public string shapeBackground;
		/// <summary>
		/// Cell Alignment Option
		/// </summary>
		public HorizontalAlignmentValues? horizontalAlignment;
		/// <summary>
		/// Gets or sets the height of the text box.
		/// </summary>
		public uint height = 100;
		/// <summary>
		/// Gets or sets the width of the text box.
		/// </summary>
		public uint width = 100;
		/// <summary>
		/// Gets or sets the X-coordinate of the text box in EMUs (English Metric Units).
		/// </summary>
		public uint x = 0;
		/// <summary>
		/// Gets or sets the Y-coordinate of the text box in EMUs (English Metric Units).
		/// </summary>
		public uint y = 0;
	}
}
