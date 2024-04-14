// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Global_2013;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Chart Class Exported out of Excel importing from Global
	/// </summary>
	public class Chart<ApplicationSpecificSetting> : ChartProperties<ApplicationSpecificSetting> where ApplicationSpecificSetting : ExcelSetting
	{
		private readonly ChartPart openXMLChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) : base(worksheet, areaChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, areaChartSetting);
		}
		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting) : base(worksheet, barChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, barChartSetting);
		}
		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) : base(worksheet, columnChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, columnChartSetting);
		}
		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) : base(worksheet, lineChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, lineChartSetting);
		}
		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) : base(worksheet, pieChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, pieChartSetting);
		}
		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) : base(worksheet, scatterChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting) : base(worksheet, comboChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitialiseChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, comboChartSetting);
		}
		internal string GetNextChartRelationId()
		{
			return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
		}
		private void ConnectDrawingToChart(Worksheet worksheet, string chartId)
		{
			// Add anchor to drawing for chart grapics
			XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new()
			{
				anchorEditType = AnchorEditType.NONE,
				from = new()
				{
					row = chartSetting.applicationSpecificSetting.from.row,
					rowOffset = chartSetting.applicationSpecificSetting.from.rowOffset,
					column = chartSetting.applicationSpecificSetting.from.column,
					columnOffset = chartSetting.applicationSpecificSetting.from.columnOffset,
				},
				to = new()
				{
					row = chartSetting.applicationSpecificSetting.to.row,
					rowOffset = chartSetting.applicationSpecificSetting.to.rowOffset,
					column = chartSetting.applicationSpecificSetting.to.column,
					columnOffset = chartSetting.applicationSpecificSetting.to.columnOffset,
				},
				drawingGraphicFrame = new()
				{
					id = (uint)worksheet.GetDrawingsPart().Parts.Count(),
					name = string.Format("Chart {0}", (uint)worksheet.GetDrawingsPart().Parts.Count()),
					chartId = chartId
				}
			});
			worksheet.GetDrawing().AppendChild(twoCellAnchor);
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
			// TODO : Ignore Till 2013 color or style implementation use
			GetChartStylePart().ChartStyle = ChartStyle.CreateChartStyles();
			GetChartColorStylePart().ColorStyle = ChartColor.CreateColorStyles();
			GetChartStylePart().ChartStyle.Save();
			GetChartColorStylePart().ColorStyle.Save();
			GetChartPart().ChartSpace.Save();
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
