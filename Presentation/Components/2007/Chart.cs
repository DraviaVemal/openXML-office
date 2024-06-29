// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.IO;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Spreadsheet_2007;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Global_2013;

namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Chart Class Exported out of PPT importing from Global
	/// </summary>
	public class Chart<XAxisType, YAxisType, ZAxisType> : ChartProperties
		where XAxisType : class, IAxisTypeOptions, new()
	 	where YAxisType : class, IAxisTypeOptions, new()
	  	where ZAxisType : class, IAxisTypeOptions, new()
	{
		private readonly ChartPart documentChartPart;
		/// <summary>
		/// Create Area Chart with provided settings
		/// Not Required Generic
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, AreaChartSetting<PresentationSetting> areaChartSetting) : base(slide, areaChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, areaChartSetting);
		}
		/// <summary>
		/// Create Bar Chart with provided settings
		/// Not Required Generic
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, BarChartSetting<PresentationSetting> barChartSetting) : base(slide, barChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, barChartSetting);
		}
		/// <summary>
		/// Create Column Chart with provided settings
		/// Not Required Generic
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ColumnChartSetting<PresentationSetting> columnChartSetting) : base(slide, columnChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, columnChartSetting);
		}
		/// <summary>
		/// Create Line Chart with provided settings
		/// Not Required Generic
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, LineChartSetting<PresentationSetting> lineChartSetting) : base(slide, lineChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, lineChartSetting);
		}
		/// <summary>
		/// Create Pie Chart with provided settings
		/// Not Required Generic
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, PieChartSetting<PresentationSetting> pieChartSetting) : base(slide, pieChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, pieChartSetting);
		}
		/// <summary>
		/// Create Scatter Chart with provided settings
		/// Not Required Generic
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ScatterChartSetting<PresentationSetting> scatterChartSetting) : base(slide, scatterChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, scatterChartSetting);
		}
		/// <summary>
		///
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, ComboChartSetting<PresentationSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting) : base(slide, comboChartSetting)
		{
			documentChartPart = slide.GetSlidePart().AddNewPart<ChartPart>(slide.GetNextSlideRelationId());
			InitializeChartParts();
			CreateChart(dataRows, comboChartSetting);
		}
		/// <summary>
		/// Get Workbook control for the chart embedded object.
		/// use OpenXML-Office.SpreadSheet Excel to load the stream and update the excel if further data addition needed other than actual chart data
		/// </summary>
		/// <returns> Chart attached workbook scheme
		/// </returns>
		public Stream GetWorkBookStream()
		{
			return GetChartPart().EmbeddedPackagePart.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite);
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
		private void CreateChart(DataCell[][] dataRows, AreaChartSetting<PresentationSetting> areaChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			AreaChart<PresentationSetting> areaChart = new AreaChart<PresentationSetting>(areaChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), areaChartSetting.hyperlinkProperties);
			SaveChanges(areaChart);
		}
		private void CreateChart(DataCell[][] dataRows, BarChartSetting<PresentationSetting> barChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			BarChart<PresentationSetting> barChart = new BarChart<PresentationSetting>(barChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), barChartSetting.hyperlinkProperties);
			SaveChanges(barChart);
		}
		private void CreateChart(DataCell[][] dataRows, ColumnChartSetting<PresentationSetting> columnChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			ColumnChart<PresentationSetting> columnChart = new ColumnChart<PresentationSetting>(columnChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), columnChartSetting.hyperlinkProperties);
			SaveChanges(columnChart);
		}
		private void CreateChart(DataCell[][] dataRows, LineChartSetting<PresentationSetting> lineChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			LineChart<PresentationSetting> lineChart = new LineChart<PresentationSetting>(lineChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), lineChartSetting.hyperlinkProperties);
			SaveChanges(lineChart);
		}
		private void CreateChart(DataCell[][] dataRows, PieChartSetting<PresentationSetting> pieChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			PieChart<PresentationSetting> pieChart = new PieChart<PresentationSetting>(pieChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), pieChartSetting.hyperlinkProperties);
			SaveChanges(pieChart);
		}
		private void CreateChart(DataCell[][] dataRows, ScatterChartSetting<PresentationSetting> scatterChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			ScatterChart<PresentationSetting> scatterChart = new ScatterChart<PresentationSetting>(scatterChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), scatterChartSetting.hyperlinkProperties);
			SaveChanges(scatterChart);
		}
		private void CreateChart(DataCell[][] dataRows, ComboChartSetting<PresentationSetting, XAxisType, YAxisType, ZAxisType> comboChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			ComboChart<PresentationSetting, XAxisType, YAxisType, ZAxisType> comboChart = new ComboChart<PresentationSetting, XAxisType, YAxisType, ZAxisType>(comboChartSetting, ExcelToPPTdata(dataRows));
			CreateChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), comboChartSetting.hyperlinkProperties);
			SaveChanges(comboChart);
		}
		private void SaveChanges(ChartBase<PresentationSetting> chart)
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
			return documentChartPart.ChartColorStyleParts.FirstOrDefault();
		}
		private ChartPart GetChartPart()
		{
			return documentChartPart;
		}
		private ChartStylePart GetChartStylePart()
		{
			return documentChartPart.ChartStyleParts.FirstOrDefault();
		}
		private void InitializeChartParts()
		{
			GetChartPart().AddNewPart<EmbeddedPackagePart>(EmbeddedPackagePartType.Xlsx.ContentType, GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
		}
	}
}
