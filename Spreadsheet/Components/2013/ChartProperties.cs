// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	///
	/// </summary>
	public class ChartProperties
	{
		/// <summary>
		///
		/// </summary>
		internal readonly ChartSetting chartSetting;

		/// <summary>
		///
		/// </summary>
		internal readonly Worksheet currentWorksheet;

		internal ChartProperties(Worksheet worksheet, ChartSetting chartSetting)
		{
			this.chartSetting = chartSetting;
			currentWorksheet = worksheet;
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// X,Y
		/// </returns>
		internal (uint, uint) GetPosition()
		{
			return (chartSetting.x, chartSetting.y);
		}

		/// <summary>
		/// </summary>
		/// <returns>
		/// Width,Height
		/// </returns>
		internal (uint, uint) GetSize()
		{
			return (chartSetting.width, chartSetting.height);
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
