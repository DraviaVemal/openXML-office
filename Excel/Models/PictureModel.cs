// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.


using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Excel_2013
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
		public int fromCol = 1;
		/// <summary>
		///
		/// </summary>
		public int fromColOff = 0;
		/// <summary>
		///
		/// </summary>
		public int fromRow = 2;
		/// <summary>
		///
		/// </summary>
		public int fromRowOff = 0;
		/// <summary>
		///
		/// </summary>
		public int toCol = 1;
		/// <summary>
		///
		/// </summary>
		public int toColOff = 0;
		/// <summary>
		///
		/// </summary>
		public int toRow = 3;
		/// <summary>
		///
		/// </summary>
		public int toRowOff = 0;
	}
}
