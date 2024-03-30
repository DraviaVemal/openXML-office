// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	/// Represents the properties of a column in a worksheet.
	/// </summary>
	public class SpreadsheetProperties
	{
		/// <summary>
		/// Spreadsheet settings
		/// </summary>
		public SpreadsheetSettings settings = new();

		/// <summary>
		/// Spreadsheet theme settings
		/// </summary>
		public ThemePallet theme = new();
	}

	/// <summary>
	/// Represents the settings of a spreadsheet.
	/// </summary>
	public class SpreadsheetSettings
	{
	}

	internal class SpreadsheetInfo
	{
		public string? filePath;
		public bool isEditable = true;
		public bool isExistingFile = false;
	}
}
