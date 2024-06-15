// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties
	{
		/// <summary>
		///
		/// </summary>
		internal readonly ChartSetting<ExcelSetting> chartSetting;
		/// <summary>
		///
		/// </summary>
		internal readonly Worksheet currentWorksheet;
		internal ChartProperties(Worksheet worksheet, ChartSetting<ExcelSetting> chartSetting)
		{
			this.chartSetting = chartSetting;
			currentWorksheet = worksheet;
		}
		/// <summary>
		/// Save Chart Part
		/// </summary>
		internal void Save()
		{
			currentWorksheet.GetWorksheetPart().Worksheet.Save();
		}
	}
}
