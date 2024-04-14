// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Represents the properties of a column in a worksheet.
	/// </summary>
	public class SpreadsheetProperties
	{
		/// <summary>
		/// Spreadsheet settings
		/// </summary>
		public SpreadsheetSettings settings = new SpreadsheetSettings();
		/// <summary>
		/// Spreadsheet theme settings
		/// </summary>
		public ThemePallet theme = new ThemePallet();
	}
	/// <summary>
	/// Represents the settings of a spreadsheet.
	/// </summary>
	public class SpreadsheetSettings
	{
	}
	internal class SpreadsheetInfo
	{
		public bool isEditable = true;
		public bool isExistingFile = false;
	}
}
