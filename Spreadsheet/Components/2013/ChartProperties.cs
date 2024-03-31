// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties<ApplicationSpecificSetting> where ApplicationSpecificSetting : ExcelSetting
	{
		/// <summary>
		///
		/// </summary>
		internal readonly ChartSetting<ApplicationSpecificSetting> chartSetting;

		/// <summary>
		///
		/// </summary>
		internal readonly Worksheet currentWorksheet;

		internal ChartProperties(Worksheet worksheet, ChartSetting<ApplicationSpecificSetting> chartSetting)
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
