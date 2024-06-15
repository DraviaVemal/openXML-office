// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Common Slide Data Class used to create the base components of a slide, slideMaster.
	/// </summary>
	public class CommonSlideData : PresentationCommonProperties
	{
		private readonly P.CommonSlideData openXMLCommonSlideData;
		internal CommonSlideData(PresentationConstants.CommonSlideDataType commonSlideDataType, PresentationConstants.SlideLayoutType layoutType)
		{
			openXMLCommonSlideData = new P.CommonSlideData()
			{
				Name = PresentationConstants.GetSlideLayoutType(layoutType)
			};
			CreateCommonSlideData(commonSlideDataType);
		}
		internal CommonSlideData(P.CommonSlideData commonSlideData)
		{
			openXMLCommonSlideData = commonSlideData;
		}
		// Return OpenXML CommonSlideData Object
		internal P.CommonSlideData GetCommonSlideData()
		{
			return openXMLCommonSlideData;
		}
		private void CreateCommonSlideData(PresentationConstants.CommonSlideDataType commonSlideDataType)
		{
			P.Background background = new P.Background()
			{
				BackgroundStyleReference = new P.BackgroundStyleReference(new A.SchemeColor()
				{
					Val = A.SchemeColorValues.Background1
				})
				{
					Index = 1001
				}
			};
			P.ShapeTree shapeTree = new P.ShapeTree()
			{
				GroupShapeProperties = new P.GroupShapeProperties()
				{
					TransformGroup = new A.TransformGroup()
					{
						Offset = new A.Offset()
						{
							X = 0,
							Y = 0
						},
						Extents = new A.Extents()
						{
							Cx = 0,
							Cy = 0
						},
						ChildOffset = new A.ChildOffset()
						{
							X = 0,
							Y = 0
						},
						ChildExtents = new A.ChildExtents()
						{
							Cx = 0,
							Cy = 0
						}
					}
				},
				NonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties(
								new P.NonVisualDrawingProperties { Id = 1, Name = "" },
								new P.NonVisualGroupShapeDrawingProperties(),
								new P.ApplicationNonVisualDrawingProperties()
							)
			};
			switch (commonSlideDataType)
			{
				case PresentationConstants.CommonSlideDataType.SLIDE_MASTER:
					var unused5 = openXMLCommonSlideData.AppendChild(background);
					var unused4 = openXMLCommonSlideData.AppendChild(shapeTree);
					break;
				case PresentationConstants.CommonSlideDataType.SLIDE_LAYOUT:
					var unused3 = shapeTree.AppendChild(CreateShape1());
					var unused2 = shapeTree.AppendChild(CreateShape2());
					var unused1 = openXMLCommonSlideData.AppendChild(shapeTree);
					break;
				default: // slide
					var unused = openXMLCommonSlideData.AppendChild(shapeTree);
					break;
			}
		}
		private static P.Shape CreateShape1()
		{
			P.Shape shape = new P.Shape();
			P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties(
				new P.NonVisualDrawingProperties { Id = 2, Name = "Title 1" },
				new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
				new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })
			);
			P.ShapeProperties shapeProperties = new P.ShapeProperties(
				new A.Transform2D(
					new A.Offset { X = 838200L, Y = 365125L },
					new A.Extents { Cx = 10515600L, Cy = 1325563L }
				),
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
			);
			P.TextBody textBody = new P.TextBody(
				new A.BodyProperties(),
				new A.ListStyle(),
				new A.Paragraph(
					new A.Run(
						new A.RunProperties { Language = "en-IN" },
						new A.Text { Text = "Click to edit Master title style" }
					),
					new A.EndParagraphRunProperties { Language = "en-IN" }
				)
			);
			shape.Append(nonVisualShapeProperties);
			shape.Append(shapeProperties);
			shape.Append(textBody);
			return shape;
		}
		private static P.Shape CreateShape2()
		{
			P.Shape shape = new P.Shape();
			P.NonVisualShapeProperties nonVisualShapeProperties = new P.NonVisualShapeProperties(
				new P.NonVisualDrawingProperties { Id = 3U, Name = "Text Placeholder 2" },
				new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
				new P.ApplicationNonVisualDrawingProperties(
					new P.PlaceholderShape { Index = 1U, Type = P.PlaceholderValues.Body })
			);
			P.ShapeProperties shapeProperties = new P.ShapeProperties(
				new A.Transform2D(
					new A.Offset { X = 838200L, Y = 1825625L },
					new A.Extents { Cx = 10515600L, Cy = 4351338L }
				),
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
			);
			P.TextBody textBody = new P.TextBody(
				new A.BodyProperties(),
				new A.ListStyle(),
				new A.Paragraph(
					new A.ParagraphProperties { Level = 0 },
					new A.Run(
						new A.RunProperties { Language = "en-IN" },
						new A.Text("Click to edit Master text styles")
					)
				),
				new A.Paragraph(
					new A.ParagraphProperties { Level = 1 },
					new A.Run(
						new A.RunProperties { Language = "en-IN" },
						new A.Text("Second Level")
					)
				),
				new A.Paragraph(
					new A.ParagraphProperties { Level = 2 },
					new A.Run(
						new A.RunProperties { Language = "en-IN" },
						new A.Text("Third Level")
					)
				),
				new A.Paragraph(
					new A.ParagraphProperties { Level = 3 },
					new A.Run(
						new A.RunProperties { Language = "en-IN" },
						new A.Text("Fourth Level")
					)
				),
				new A.Paragraph(
					new A.ParagraphProperties { Level = 4 },
					new A.Run(
						new A.RunProperties { Language = "en-IN" },
						new A.Text("Fifth Level")
					),
					new A.EndParagraphRunProperties()
					{
						Language = "en-IN"
					}
				)
			);
			shape.Append(nonVisualShapeProperties);
			shape.Append(shapeProperties);
			shape.Append(textBody);
			return shape;
		}
	}
}
