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
	public class Chart<XAxisType, YAxisType, ZAxisType> : ChartProperties
		where XAxisType : class, IAxisTypeOptions, new()
	 	where YAxisType : class, IAxisTypeOptions, new()
	  	where ZAxisType : class, IAxisTypeOptions, new()
	{
		private readonly ChartPart openXMLChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, AreaChartSetting<ExcelSetting> areaChartSetting) : base(worksheet, areaChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, areaChartSetting);
		}
		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, BarChartSetting<ExcelSetting> barChartSetting) : base(worksheet, barChartSetting)
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
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, ColumnChartSetting<ExcelSetting> columnChartSetting) : base(worksheet, columnChartSetting)
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
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, LineChartSetting<ExcelSetting> lineChartSetting) : base(worksheet, lineChartSetting)
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
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, PieChartSetting<ExcelSetting> pieChartSetting) : base(worksheet, pieChartSetting)
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
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, ScatterChartSetting<ExcelSetting> scatterChartSetting) : base(worksheet, scatterChartSetting)
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
		internal Chart(Worksheet worksheet, ChartData[][] chartData, DataRange dataRange, ComboChartSetting<ExcelSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting) : base(worksheet, comboChartSetting)
		{
			string chartId = worksheet.GetNextDrawingPartRelationId();
			openXMLChartPart = worksheet.GetDrawingsPart().AddNewPart<ChartPart>(chartId);
			InitializeChartParts();
			ConnectDrawingToChart(worksheet, chartId);
			CreateChart(chartData, dataRange, comboChartSetting);
		}
		internal string GetNextChartRelationId()
		{
			int nextId = GetChartPart().Parts.Count() + GetChartPart().ExternalRelationships.Count() + GetChartPart().HyperlinkRelationships.Count() + GetChartPart().DataPartReferenceRelationships.Count();
			do
			{
				++nextId;
			} while (GetChartPart().Parts.Any(item => item.RelationshipId == string.Format("rId{0}", nextId)) ||
			GetChartPart().ExternalRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
			GetChartPart().HyperlinkRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)) ||
			GetChartPart().DataPartReferenceRelationships.Any(item => item.Id == string.Format("rId{0}", nextId)));
			return string.Format("rId{0}", nextId);
		}
		private void ConnectDrawingToChart(Worksheet worksheet, string chartId)
		{
			// Add anchor to drawing for chart graphics
			XDR.TwoCellAnchor twoCellAnchor = worksheet.CreateTwoCellAnchor(new TwoCellAnchorModel<NoOptions>()
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
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, AreaChartSetting<ExcelSetting> areaChartSetting)
		{
			AreaChart<ExcelSetting> areaChart = new AreaChart<ExcelSetting>(areaChartSetting, chartData, dataRange);
			SaveChanges(areaChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, BarChartSetting<ExcelSetting> barChartSetting)
		{
			BarChart<ExcelSetting> barChart = new BarChart<ExcelSetting>(barChartSetting, chartData, dataRange);
			SaveChanges(barChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, ColumnChartSetting<ExcelSetting> columnChartSetting)
		{
			ColumnChart<ExcelSetting> columnChart = new ColumnChart<ExcelSetting>(columnChartSetting, chartData, dataRange);
			SaveChanges(columnChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, LineChartSetting<ExcelSetting> lineChartSetting)
		{
			LineChart<ExcelSetting> lineChart = new LineChart<ExcelSetting>(lineChartSetting, chartData, dataRange);
			SaveChanges(lineChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, PieChartSetting<ExcelSetting> pieChartSetting)
		{
			PieChart<ExcelSetting> pieChart = new PieChart<ExcelSetting>(pieChartSetting, chartData, dataRange);
			SaveChanges(pieChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, ScatterChartSetting<ExcelSetting> scatterChartSetting)
		{
			ScatterChart<ExcelSetting> scatterChart = new ScatterChart<ExcelSetting>(scatterChartSetting, chartData, dataRange);
			SaveChanges(scatterChart);
		}
		private void CreateChart(ChartData[][] chartData, DataRange dataRange, ComboChartSetting<ExcelSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting)
		{
			ComboChart<ExcelSetting, XAxisType, YAxisType, ZAxisType> comboChart = new ComboChart<ExcelSetting, XAxisType, YAxisType, ZAxisType>(comboChartSetting, chartData, dataRange);
			SaveChanges(comboChart);
		}
		private void SaveChanges(ChartBase<ExcelSetting> chart)
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
