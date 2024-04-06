// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.


using OpenXMLOffice.Global_2007;

namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	///
	/// </summary>
	public class ExcelPictureSetting
	{
		/// <summary>
		/// The type of image.
		/// </summary>
		public ImageType imageType = ImageType.JPEG;

		/// <summary>
		///
		/// </summary>
		public uint fromCol = 1;
		/// <summary>
		///
		/// </summary>
		public uint fromColOff = 0;
		/// <summary>
		///
		/// </summary>
		public uint fromRow = 2;
		/// <summary>
		///
		/// </summary>
		public uint fromRowOff = 0;
		/// <summary>
		///
		/// </summary>
		public uint toCol = 1;
		/// <summary>
		///
		/// </summary>
		public uint toColOff = 0;
		/// <summary>
		///
		/// </summary>
		public uint toRow = 3;
		/// <summary>
		///
		/// </summary>
		public uint toRowOff = 0;
	}
}
