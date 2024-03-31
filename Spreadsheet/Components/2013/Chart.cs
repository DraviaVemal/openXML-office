// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.


using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Spreadsheet_2013
{
	/// <summary>
	/// Chart Class Exported out of Excel importing from Global
	/// </summary>
	public class Chart<ApplicationSpecificSetting> : ChartProperties<ApplicationSpecificSetting> where ApplicationSpecificSetting : class, ISizeAndPosition
	{
		private readonly ChartPart openXMLChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) : base(worksheet, areaChartSetting)
		{
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, areaChartSetting);
		}

		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting) : base(worksheet, barChartSetting)
		{
			openXMLChartPart = worksheet.GetWorksheetPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, barChartSetting);
		}

		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) : base(worksheet, columnChartSetting)
		{
			openXMLChartPart = worksheet.GetWorksheetPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, columnChartSetting);
		}

		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) : base(worksheet, lineChartSetting)
		{
			openXMLChartPart = worksheet.GetWorksheetPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, lineChartSetting);
		}

		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) : base(worksheet, pieChartSetting)
		{
			openXMLChartPart = worksheet.GetWorksheetPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, pieChartSetting);
		}

		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) : base(worksheet, scatterChartSetting)
		{
			openXMLChartPart = worksheet.GetWorksheetPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting) : base(worksheet, comboChartSetting)
		{
			openXMLChartPart = worksheet.GetWorksheetPart().AddNewPart<ChartPart>(worksheet.GetNextSheetPartRelationId());
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet);
			CreateChart(chartDatas, dataRange, comboChartSetting);
		}

		internal string GetNextChartRelationId()
		{
			return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
		}

		private static void ConnectDrawingToChart(Worksheet worksheet)
		{
			// Add anchor to drawing for chart grapics
			worksheet.CreateTwoCellAnchor(new()
			{

			});
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting)
		{
			AreaChart<ApplicationSpecificSetting> areaChart = new(areaChartSetting, chartDatas, dataRange);
			SaveChanges(areaChart);
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting)
		{
			BarChart<ApplicationSpecificSetting> barChart = new(barChartSetting, chartDatas, dataRange);
			SaveChanges(barChart);
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting)
		{
			ColumnChart<ApplicationSpecificSetting> columnChart = new(columnChartSetting, chartDatas, dataRange);
			SaveChanges(columnChart);
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting)
		{
			LineChart<ApplicationSpecificSetting> lineChart = new(lineChartSetting, chartDatas, dataRange);
			SaveChanges(lineChart);
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting)
		{
			PieChart<ApplicationSpecificSetting> pieChart = new(pieChartSetting, chartDatas, dataRange);
			SaveChanges(pieChart);
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting)
		{
			ScatterChart<ApplicationSpecificSetting> scatterChart = new(scatterChartSetting, chartDatas, dataRange);
			SaveChanges(scatterChart);
		}

		private void CreateChart(ChartData[][] chartDatas, DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting)
		{
			ComboChart<ApplicationSpecificSetting> comboChart = new(comboChartSetting, chartDatas, dataRange);
			SaveChanges(comboChart);
		}

		private void SaveChanges(ChartBase<ApplicationSpecificSetting> chart)
		{
			GetChartPart().ChartSpace = chart.GetChartSpace();
			GetChartStylePart().ChartStyle = ChartStyle.CreateChartStyles();
			GetChartColorStylePart().ColorStyle = ChartColor.CreateColorStyles();
			// Save All Changes
			GetChartPart().ChartSpace.Save();
			GetChartStylePart().ChartStyle.Save();
			GetChartColorStylePart().ColorStyle.Save();
		}

		private ChartColorStylePart GetChartColorStylePart()
		{
			return openXMLChartPart.ChartColorStyleParts.FirstOrDefault()!;
		}

		private ChartPart GetChartPart()
		{
			return openXMLChartPart;
		}

		private ChartStylePart GetChartStylePart()
		{
			return openXMLChartPart.ChartStyleParts.FirstOrDefault()!;
		}

		private void InitialiseChartParts()
		{
			GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
		}
	}

}