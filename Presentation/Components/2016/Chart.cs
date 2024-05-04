// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Packaging;
using OpenXMLOffice.Presentation_2007;
using OpenXMLOffice.Spreadsheet_2007;
using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Global_2013;
using OpenXMLOffice.Global_2016;
using System.IO;
using System.Linq;
namespace OpenXMLOffice.Presentation_2016
{
	/// <summary>
	///
	/// </summary>
	public class Chart<ApplicationSpecificSetting> : AdvancedChartProperties<ApplicationSpecificSetting> where ApplicationSpecificSetting : PresentationSetting, new()
	{
		/// <summary>
		///
		/// </summary>
		private ExtendedChartPart OpenXMLChartPart { get; set; }
		/// <summary>
		///
		/// </summary>
		public Chart(Slide slide, DataCell[][] dataRows, WaterfallChartSetting<ApplicationSpecificSetting> waterfallChartSetting) : base(slide, waterfallChartSetting)
		{
			OpenXMLChartPart = slide.GetSlidePart().AddNewPart<ExtendedChartPart>(slide.GetNextSlideRelationId());
			InitialiseChartParts();
			CreateChart(dataRows, waterfallChartSetting);
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
			return string.Format("rId{0}", GetChartPart().Parts.Count() + GetChartPart().ExternalRelationships.Count() + GetChartPart().HyperlinkRelationships.Count() + GetChartPart().DataPartReferenceRelationships.Count() + 1);
		}
		private ExtendedChartPart GetChartPart()
		{
			return OpenXMLChartPart;
		}
		private void CreateChart(DataCell[][] dataRows, WaterfallChartSetting<ApplicationSpecificSetting> waterfallChartSetting)
		{
			using (Stream stream = GetWorkBookStream())
			{
				WriteDataToExcel(dataRows, stream);
			};
			WaterfallChart<ApplicationSpecificSetting> waterfallChart = new WaterfallChart<ApplicationSpecificSetting>(waterfallChartSetting, ExcelToPPTdata(dataRows));
			CreateExtendedChartGraphicFrame(currentSlide.GetSlidePart().GetIdOfPart(GetChartPart()), (uint)currentSlide.GetSlidePart().GetPartsOfType<ChartPart>().Count(), waterfallChartSetting.hyperlinkProperties);
			SaveChanges(waterfallChart);
		}
		private void SaveChanges(AdvanceCharts<ApplicationSpecificSetting> chart)
		{
			GetChartPart().ChartSpace = chart.GetExtendedChartSpace();
			GetChartStylePart().ChartStyle = ChartStyle.CreateChartStyles();
			GetChartColorStylePart().ColorStyle = ChartColor.CreateColorStyles();
			// Save All Changes
			GetChartPart().ChartSpace.Save();
			GetChartStylePart().ChartStyle.Save();
			GetChartColorStylePart().ColorStyle.Save();
		}
		private ChartColorStylePart GetChartColorStylePart()
		{
			return OpenXMLChartPart.ChartColorStyleParts.FirstOrDefault();
		}
		private ChartStylePart GetChartStylePart()
		{
			return OpenXMLChartPart.ChartStyleParts.FirstOrDefault();
		}
		private void InitialiseChartParts()
		{
			GetChartPart().AddNewPart<EmbeddedPackagePart>(EmbeddedPackagePartType.Xlsx.ContentType, GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartColorStylePart>(GetNextChartRelationId());
			GetChartPart().AddNewPart<ChartStylePart>(GetNextChartRelationId());
		}
	}
}
