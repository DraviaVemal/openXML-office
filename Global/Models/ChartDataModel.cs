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
}
