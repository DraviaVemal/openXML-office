// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the Data type of the chart data.
	/// </summary>
	public enum DataType
	{
		/// <summary>
		/// Date Data Type
		/// </summary>
		DATE,

		/// <summary>
		/// Number Data Type
		/// </summary>
		NUMBER,

		/// <summary>
		/// String Data Type
		/// </summary>
		STRING
	}

	/// <summary>
	/// Represents the settings for a chart data.
	/// </summary>
	public class ChartData
	{
		/// <summary>
		/// The data type of the chart data.
		/// </summary>
		public DataType dataType = DataType.STRING;

		/// <summary>
		/// Number Format for Chart Data (Default: General)
		/// </summary>
		public string numberFormat = "General";

		/// <summary>
		/// The value of the chart data.
		/// </summary>
		public string? value;
	}

	/// <summary>
	/// Set the target data present with in the excel workbook
	/// </summary>
	public class DataRange
	{
		/// <summary>
		/// Sheet name
		/// Default : Current Sheet the chart is targetted for
		/// </summary>
		public string? sheetName { get; set; }
		/// <summary>
		/// Cell Id Range start. Ex: A1
		/// </summary>
		public required string cellIdStart { get; set; }
		/// <summary>
		/// Cell Id Range End. Ex: D4
		/// </summary>
		public required string cellIdEnd { get; set; }
	}
}
