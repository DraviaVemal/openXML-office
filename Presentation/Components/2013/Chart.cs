// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Excel_2013;
using OpenXMLOffice.Global_2013;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2013
{
	/// <summary>
	/// Chart Class Exported out of PPT importing from Global
	/// </summary>
	public class Chart : ChartProperties
	{
		private readonly ChartPart openXMLChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, AreaChartSetting areaChartSetting) : base(slide, areaChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, areaChartSetting);
		}

		/// <summary>
		/// Create Bar Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, BarChartSetting barChartSetting) : base(slide, barChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, barChartSetting);
		}

		/// <summary>
		/// Create Column Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ColumnChartSetting columnChartSetting) : base(slide, columnChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, columnChartSetting);
		}

		/// <summary>
		/// Create Line Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, LineChartSetting lineChartSetting) : base(slide, lineChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, lineChartSetting);
		}

		/// <summary>
		/// Create Pie Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, PieChartSetting pieChartSetting) : base(slide, pieChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, pieChartSetting);
		}

		/// <summary>
		/// Create Scatter Chart with provided settings
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ScatterChartSetting scatterChartSetting) : base(slide, scatterChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ComboChartSetting comboChartSetting) : base(slide, comboChartSetting)
		{
			openXMLChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, comboChartSetting);
		}

		/// <summary>
		/// Get Worksheet control for the chart embedded object
		/// </summary>
		/// <returns>
		/// </returns>
		public Spreadsheet GetChartWorkBook()
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			return new(stream);
		}

		internal string GetNextChartRelationId()
		{
			return string.Format("rId{0}", GetChartPart().Parts.Count() + 1);
		}

		private void CreateChart(DataCell[][] dataRows, AreaChartSetting areaChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			AreaChart areaChart = new(areaChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(areaChart);
		}

		private void CreateChart(DataCell[][] dataRows, BarChartSetting barChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			BarChart barChart = new(barChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(barChart);
		}

		private void CreateChart(DataCell[][] dataRows, ColumnChartSetting columnChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			ColumnChart columnChart = new(columnChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(columnChart);
		}

		private void CreateChart(DataCell[][] dataRows, LineChartSetting lineChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			LineChart lineChart = new(lineChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(lineChart);
		}

		private void CreateChart(DataCell[][] dataRows, PieChartSetting pieChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			PieChart pieChart = new(pieChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(pieChart);
		}

		private void CreateChart(DataCell[][] dataRows, ScatterChartSetting scatterChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			ScatterChart scatterChart = new(scatterChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(scatterChart);
		}

		private void CreateChart(DataCell[][] dataRows, ComboChartSetting comboChartSetting)
		{
			Stream stream = GetChartPart().EmbeddedPackagePart!.GetStream();
			LoadDataToExcel(dataRows, stream);
			// Prepare Excel Data for PPT Cache
			ComboChart comboChart = new(comboChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count());
			SaveChanges(comboChart);
		}

		private void SaveChanges(ChartBase chart)
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
			GetChartPart().AddNewPart<EmbeddedPackagePart>(EmbeddedPackagePartType.Xlsx.ContentType, GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
		}


	}
}
