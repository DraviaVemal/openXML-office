// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// Common Slide Data Class used to create the base components of a slide, slidemaster.
	/// </summary>
	public class CommonSlideData
	{
		private readonly P.CommonSlideData openXMLCommonSlideData;

		internal CommonSlideData(PresentationConstants.CommonSlideDataType commonSlideDataType, PresentationConstants.SlideLayoutType layoutType)
		{
			openXMLCommonSlideData = new()
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
			Background background = new()
			{
				BackgroundStyleReference = new BackgroundStyleReference(new A.SchemeColor()
				{
					Val = A.SchemeColorValues.Background1
				})
				{
					Index = 1001
				}
			};
			ShapeTree shapeTree = new()
			{
				GroupShapeProperties = new GroupShapeProperties()
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
				NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
								new NonVisualDrawingProperties { Id = 1, Name = "" },
								new NonVisualGroupShapeDrawingProperties(),
								new ApplicationNonVisualDrawingProperties()
							)
			};

			switch (commonSlideDataType)
			{
				case PresentationConstants.CommonSlideDataType.SLIDE_MASTER:
					openXMLCommonSlideData.AppendChild(background);
					openXMLCommonSlideData.AppendChild(shapeTree);
					break;

				case PresentationConstants.CommonSlideDataType.SLIDE_LAYOUT:
					shapeTree.AppendChild(CreateShape1());
					shapeTree.AppendChild(CreateShape2());
					openXMLCommonSlideData.AppendChild(shapeTree);
					break;

				default: // slide
					openXMLCommonSlideData.AppendChild(shapeTree);
					break;
			}
		}

		private static P.Shape CreateShape1()
		{
			P.Shape shape = new();
			NonVisualShapeProperties nonVisualShapeProperties = new(
				new NonVisualDrawingProperties { Id = 2, Name = "Title 1" },
				new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
				new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Title })
			);
			ShapeProperties shapeProperties = new(
				new A.Transform2D(
					new A.Offset { X = 838200L, Y = 365125L },
					new A.Extents { Cx = 10515600L, Cy = 1325563L }
				),
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
			);
			TextBody textBody = new(
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
			P.Shape shape = new();
			NonVisualShapeProperties nonVisualShapeProperties = new(
				new NonVisualDrawingProperties { Id = 3U, Name = "Text Placeholder 2" },
				new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
				new ApplicationNonVisualDrawingProperties(
					new PlaceholderShape { Index = 1U, Type = PlaceholderValues.Body })
			);
			ShapeProperties shapeProperties = new(
				new A.Transform2D(
					new A.Offset { X = 838200L, Y = 1825625L },
					new A.Extents { Cx = 10515600L, Cy = 4351338L }
				),
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
			);
			TextBody textBody = new(
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
