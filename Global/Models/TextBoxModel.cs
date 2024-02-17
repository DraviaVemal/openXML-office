// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
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
	/// Represents the settings for a text box.
	/// </summary>
	public class TextBoxSetting
	{
		/// <summary>
		/// Cell Alignment Option
		/// </summary>
		public HorizontalAlignmentValues? horizontalAlignment;

		/// <summary>
		/// Gets or sets the font family of the text.
		/// </summary>
		public string fontFamily = "Calibri (Body)";

		/// <summary>
		/// Gets or sets the font size of the text.
		/// </summary>
		public int fontSize = 18;

		/// <summary>
		/// Gets or sets the height of the text box.
		/// </summary>
		public uint height = 100;

		/// <summary>
		/// Gets or sets a value indicating whether the text is bold.
		/// </summary>
		public bool isBold = false;

		/// <summary>
		/// Gets or sets a value indicating whether the text is italic.
		/// </summary>
		public bool isItalic = false;

		/// <summary>
		/// Gets or sets a value indicating whether the text is underlined.
		/// </summary>
		public bool isUnderline = false;

		/// <summary>
		/// Gets or sets the background color of the text box shape.
		/// </summary>
		public string? shapeBackground;

		/// <summary>
		/// Gets or sets the text content of the text box.
		/// </summary>
		public string text = "Text Box";

		/// <summary>
		/// Gets or sets the background color of the text.
		/// </summary>
		public string? textBackground;

		/// <summary>
		/// Gets or sets the color of the text.
		/// </summary>
		public string textColor = "000000";

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
