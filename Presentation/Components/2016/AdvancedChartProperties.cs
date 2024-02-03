// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Runtime;
using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2013;
using OpenXMLOffice.Presentation_2013;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OpenXMLOffice.Presentation_2016
{
	/// <summary>
	///
	/// </summary>
	public class AdvancedChartProperties : ChartProperties
	{
		private AlternateContent? alternateContent;

		private readonly TextBox errorMessage;

		/// <summary>
		///
		/// </summary>
		public AdvancedChartProperties(Slide slide, ChartSetting chartSetting) : base(slide, chartSetting)
		{
			errorMessage = new TextBox(new()
			{
				text = "This chart is not supported in this version of PowerPoint. Requires PowerPoint 2016 or later.",
				x = chartSetting.x,
				y = chartSetting.y,
				width = chartSetting.width,
				height = chartSetting.height,
			});
		}

		/// <summary>
		///
		/// </summary>
		internal void CreateExtendedChartGraphicFrame(string relationshipId, uint id)
		{
			// Load Chart Part To Graphics Frame For Export
			P.NonVisualGraphicFrameProperties nonVisualProperties = new()
			{
				NonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = id, Name = "Chart" },
				NonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties(),
				ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
			};
			graphicFrame = new()
			{
				NonVisualGraphicFrameProperties = nonVisualProperties,
				Transform = new P.Transform(
				   new A.Offset
				   {
					   X = chartSetting.x,
					   Y = chartSetting.y
				   },
				   new A.Extents
				   {
					   Cx = chartSetting.width,
					   Cy = chartSetting.height
				   }),
				Graphic = new A.Graphic(
				   new A.GraphicData(
					new CX.RelId() { Id = relationshipId }
				   )
				   {
					   Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex"
				   })
			};
			CreateAlternateContent();
		}

		/// <summary>
		///
		/// </summary>
		public override void UpdateSize(uint width, uint height)
		{
			base.UpdateSize(width, height);
			errorMessage.UpdateSize(width, height);
		}

		/// <summary>
		///
		/// </summary>
		public override void UpdatePosition(uint x, uint y)
		{
			base.UpdatePosition(x, y);
			errorMessage.UpdatePosition(x, y);
		}

		private void CreateAlternateContent()
		{
			alternateContent = new(
				new AlternateContentChoice(
					(OpenXmlElement)graphicFrame!.Clone()
				)
				{ Requires = "cx1" },
				new AlternateContentFallback(
					errorMessage.GetTextBoxShape()
				)
			);
			alternateContent.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
		}

		internal AlternateContent GetAlternateContent()
		{
			return alternateContent!;
		}

		new internal void GetChartGraphicFrame()
		{
			throw new AmbiguousImplementationException("Use GetAlternateContent() instead.");
		}
	}
}
