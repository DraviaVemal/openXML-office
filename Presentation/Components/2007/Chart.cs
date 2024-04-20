// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Spreadsheet_2007;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Global_2013;
using System.IO;
using System.Linq;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Chart Class Exported out of PPT importing from Global
	/// </summary>
	public class Chart<ApplicationSpecificSetting> : ChartProperties<ApplicationSpecificSetting> where ApplicationSpecificSetting : PresentationSetting, new()
	{
		private readonly ChartPart openXMLChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting) : base(slide, areaChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, areaChartSetting);
		}
		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, BarChartSetting<ApplicationSpecificSetting> barChartSetting) : base(slide, barChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, barChartSetting);
		}
		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting) : base(slide, columnChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, columnChartSetting);
		}
		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, LineChartSetting<ApplicationSpecificSetting> lineChartSetting) : base(slide, lineChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, lineChartSetting);
		}
		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, PieChartSetting<ApplicationSpecificSetting> pieChartSetting) : base(slide, pieChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, pieChartSetting);
		}
		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting) : base(slide, scatterChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting) : base(slide, comboChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, comboChartSetting);
		}
		/// <summary>
		/// Get Workbook control for the chart embedded object.
		/// use OpenXML-Office.SpreadSheet Excel to load the stream and update the excel if furthere data addition needed other than actual chart data
		/// </summary>
		/// <returns> Chart attached workbook scheme
		/// </returns>
		public Stream GetWorkBookStream()
		{
			return GetChartPart().EmbeddedPackagePart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite);
		}
		internal string GetNextChartRelationId()
		{
			return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
		}
		private void CreateChart(DataCell[][] dataRows, AreaChartSetting<ApplicationSpecificSetting> areaChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			AreaChart<ApplicationSpecificSetting> areaChart = new AreaChart<ApplicationSpecificSetting>(areaChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(areaChart);
		}
		private void CreateChart(DataCell[][] dataRows, BarChartSetting<ApplicationSpecificSetting> barChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			BarChart<ApplicationSpecificSetting> barChart = new BarChart<ApplicationSpecificSetting>(barChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(barChart);
		}
		private void CreateChart(DataCell[][] dataRows, ColumnChartSetting<ApplicationSpecificSetting> columnChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			ColumnChart<ApplicationSpecificSetting> columnChart = new ColumnChart<ApplicationSpecificSetting>(columnChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(columnChart);
		}
		private void CreateChart(DataCell[][] dataRows, LineChartSetting<ApplicationSpecificSetting> lineChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			LineChart<ApplicationSpecificSetting> lineChart = new LineChart<ApplicationSpecificSetting>(lineChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(lineChart);
		}
		private void CreateChart(DataCell[][] dataRows, PieChartSetting<ApplicationSpecificSetting> pieChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			PieChart<ApplicationSpecificSetting> pieChart = new PieChart<ApplicationSpecificSetting>(pieChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(pieChart);
		}
		private void CreateChart(DataCell[][] dataRows, ScatterChartSetting<ApplicationSpecificSetting> scatterChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			ScatterChart<ApplicationSpecificSetting> scatterChart = new ScatterChart<ApplicationSpecificSetting>(scatterChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(scatterChart);
		}
		private void CreateChart(DataCell[][] dataRows, ComboChartSetting<ApplicationSpecificSetting> comboChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart.GetStream();
			WriteDataToExcel(dataRows, stream);
			ComboChart<ApplicationSpecificSetting> comboChart = new ComboChart<ApplicationSpecificSetting>(comboChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(comboChart);
		}
		private void SaveChanges(ChartBase<ApplicationSpecificSetting> chart)
		{
			chart.GetChartSpace().Append(new C.ExternalData(
				new C.AutoUpdate() { Val = false })
			{ Id = "rId1" });
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
		private void InitialiseChartParts()
		{
			GetChartPart().AddNewPart<EmbeddedPackagePart>(EmbeddedPackagePartType.Xlsx.ContentType, GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
		}
	}
}
