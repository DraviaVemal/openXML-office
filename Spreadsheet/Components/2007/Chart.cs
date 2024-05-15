// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Global_2013;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Chart Class Exported out of Excel importing from Global
	/// </summary>
	public class Chart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> : ChartProperties<ApplicationSpecificSetting>
		where ApplicationSpecificSetting : ExcelSetting, new()
		where XAxisType : class, IAxisTypeOptions, new()
	 	where YAxisType : class, IAxisTypeOptions, new()
	  	where ZAxisType : class, IAxisTypeOptions, new()
	{
		private readonly ChartPart openXMLChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartDatas, DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) : base(worksheet, areaChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartDatas, dataRange, areaChartSetting);
		}
		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting) : base(worksheet, barChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, barChartSetting);
		}
		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) : base(worksheet, columnChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, columnChartSetting);
		}
		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) : base(worksheet, lineChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, lineChartSetting);
		}
		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) : base(worksheet, pieChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, pieChartSetting);
		}
		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) : base(worksheet, scatterChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting) : base(worksheet, comboChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, comboChartSetting);
		}
		internal string GetNextChartRelationId()
		{
			return string.Format("rId{0}", GetChartPart().Parts.Count() + GetChartPart().ExternalRelationships.Count() + GetChartPart().HyperlinkRelationships.Count() + GetChartPart().DataPartReferenceRelationships.Count() + 1);
		}
		private void ConnectDrawingToChart(Worksheet worksheet, string chartId)
		{
			// Add anchor to drawing for chart graphics
			XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel()
			{
				anchorEditType = AnchorEditType.NONE,
				from = new AnchorPosition()
				{
					row = chartSetting.applicationSpecificSetting.from.row,
					rowOffset = chartSetting.applicationSpecificSetting.from.rowOffset,
					column = chartSetting.applicationSpecificSetting.from.column,
					columnOffset = chartSetting.applicationSpecificSetting.from.columnOffset,
				},
				to = new AnchorPosition()
				{
					row = chartSetting.applicationSpecificSetting.to.row,
					rowOffset = chartSetting.applicationSpecificSetting.to.rowOffset,
					column = chartSetting.applicationSpecificSetting.to.column,
					columnOffset = chartSetting.applicationSpecificSetting.to.columnOffset,
				},
				drawingGraphicFrame = new DrawingGraphicFrame()
				{
					id = (uint)worksheet.GetDrawingsPart().Parts.Count(),
					name = string.Format("Chart {0}", (uint)worksheet.GetDrawingsPart().Parts.Count()),
					chartId = chartId
				}
			});
			worksheet.GetDrawing().AppendChild(twoCellAnchor);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting)
		{
			AreaChart<ApplicationSpecificSetting> areaChart = new AreaChart<ApplicationSpecificSetting>(areaChartSetting, chartData, dataRange);
			SaveChanges(areaChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, BarChartSetting<ApplicationSpecificSetting> barChartSetting)
		{
			BarChart<ApplicationSpecificSetting> barChart = new BarChart<ApplicationSpecificSetting>(barChartSetting, chartData, dataRange);
			SaveChanges(barChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting)
		{
			ColumnChart<ApplicationSpecificSetting> columnChart = new ColumnChart<ApplicationSpecificSetting>(columnChartSetting, chartData, dataRange);
			SaveChanges(columnChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, LineChartSetting<ApplicationSpecificSetting> lineChartSetting)
		{
			LineChart<ApplicationSpecificSetting> lineChart = new LineChart<ApplicationSpecificSetting>(lineChartSetting, chartData, dataRange);
			SaveChanges(lineChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, PieChartSetting<ApplicationSpecificSetting> pieChartSetting)
		{
			PieChart<ApplicationSpecificSetting> pieChart = new PieChart<ApplicationSpecificSetting>(pieChartSetting, chartData, dataRange);
			SaveChanges(pieChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting)
		{
			ScatterChart<ApplicationSpecificSetting> scatterChart = new ScatterChart<ApplicationSpecificSetting>(scatterChartSetting, chartData, dataRange);
			SaveChanges(scatterChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, ComboChartSetting<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting)
		{
			ComboChart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType> comboChart = new ComboChart<ApplicationSpecificSetting, XAxisType, YAxisType, ZAxisType>(comboChartSetting, chartData, dataRange);
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
			return openXMLChartPart.ChartColorStyleParts.FirstOrDefault();
		}
		private ChartPart GetChartPart()
		{
			return openXMLChartPart;
		}
		private ChartStylePart GetChartStylePart()
		{
			return openXMLChartPart.ChartStyleParts.FirstOrDefault();
		}
		private void InitializeChartParts()
		{
			GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
		}
	}
}
