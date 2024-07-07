// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// 
	/// </summary>
	public enum BulletsAndNumberingValues
	{
		/// <summary>
		/// 
		/// </summary>
		NONE,
		/// <summary>
		/// 
		/// </summary>
		FILLED_ROUND,
		/// <summary>
		/// 
		/// </summary>
		HOLLOW_ROUND,
		/// <summary>
		/// 
		/// </summary>
		FILLED_SQUARE,
		/// <summary>
		/// 
		/// </summary>
		HOLLOW_SQUARE,
		/// <summary>
		/// 
		/// </summary>
		STAR_BULLET,
		/// <summary>
		/// 
		/// </summary>
		ARROW_BULLET,
		/// <summary>
		/// 
		/// </summary>
		CHECK_BULLET,
		/// <summary>
		/// 
		/// </summary>
		NUMERIC_DOT,
		/// <summary>
		/// 
		/// </summary>
		NUMERIC_BRACKET,
		/// <summary>
		/// 
		/// </summary>
		ROMAN_CAPS,
		/// <summary>
		/// 
		/// </summary>
		ROMAN_SMALL,
		/// <summary>
		/// 
		/// </summary>
		ALPHABET_CAPS,
		/// <summary>
		/// 
		/// </summary>
		ALPHABET_SMALL,
		/// <summary>
		/// 
		/// </summary>
		ALPHABET_SMALL_BRACKET
	}
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
		/// Used to end a paragraph in text block list group. Useful in mentioning when using numbering and bullets sequence
		/// </summary>
		public bool isEndParagraph = false;
		/// <summary>
		/// 
		/// </summary>
		public BulletsAndNumberingValues? bulletsAndNumbering = BulletsAndNumberingValues.NONE;
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
		/// Define Each section of string and its property that goes in same paragraph
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
		public int height = 100;
		/// <summary>
		/// Gets or sets the width of the text box.
		/// </summary>
		public int width = 100;
		/// <summary>
		/// Gets or sets the X-coordinate of the text box in EMUs (English Metric Units).
		/// </summary>
		public int x = 0;
		/// <summary>
		/// Gets or sets the Y-coordinate of the text box in EMUs (English Metric Units).
		/// </summary>
		public int y = 0;
	}
}
