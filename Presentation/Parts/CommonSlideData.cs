// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Collections.Generic;
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
					openXMLCommonSlideData.AppendChild(background);
					openXMLCommonSlideData.AppendChild(shapeTree);
					break;
				case PresentationConstants.CommonSlideDataType.SLIDE_LAYOUT:
					shapeTree.AppendChild(CreateShape(new ShapeModel<SolidOptions, ShapeRectangleModel<PresentationSetting, SolidOptions, NoFillOptions>>()
					{
						id = (uint)shapeTree.ChildElements.Count + 1,
						name = "Title 1",
						shapeTypeOptions = new ShapeRectangleModel<PresentationSetting, SolidOptions, NoFillOptions>()
						{
							rectangleType = ShapeRectangleTypes.RECTANGLE,
							lineColorOption = new SolidOptions()
							{
								hexColor = "FFFFFF",
							}
						},
						shapePropertiesModel = new ShapePropertiesModel()
						{
							x = 838200L,
							y = 365125L,
							cx = 10515600L,
							cy = 1325563L
						},
						drawingParagraph = new DrawingParagraphModel<SolidOptions>()
						{
							drawingRuns = new List<DrawingRunModel<SolidOptions>>()
							{
								new DrawingRunModel<SolidOptions>(){
									text = "Click to edit Master title style",
									drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
								}
							}.ToArray()
						}
					}));
					shapeTree.AppendChild(CreateShape(new ShapeModel<SolidOptions, ShapeRectangleModel<PresentationSetting, SolidOptions, NoFillOptions>>()
					{
						id = (uint)shapeTree.ChildElements.Count + 1,
						name = "Text Placeholder 1",
						shapeTypeOptions = new ShapeRectangleModel<PresentationSetting, SolidOptions, NoFillOptions>()
						{
							rectangleType = ShapeRectangleTypes.RECTANGLE,
							lineColorOption = new SolidOptions()
							{
								hexColor = "FFFFFF",
							}
						},
						shapePropertiesModel = new ShapePropertiesModel()
						{
							x = 838200L,
							y = 1825625L,
							cx = 10515600L,
							cy = 4351338L
						},
						drawingParagraph = new DrawingParagraphModel<SolidOptions>()
						{
							drawingRuns = new List<DrawingRunModel<SolidOptions>>()
							{
								new DrawingRunModel<SolidOptions>(){
									text = "Click to edit Master title style",
									drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
								},
								new DrawingRunModel<SolidOptions>(){
									text = "Second Level",
									drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
								},
								new DrawingRunModel<SolidOptions>(){
									text = "Third Level",
									drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
								},
								new DrawingRunModel<SolidOptions>(){
									text = "Fourth Level",
									drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
								},
								new DrawingRunModel<SolidOptions>(){
									text = "Fifth Level",
									drawingRunProperties = new DrawingRunPropertiesModel<SolidOptions>()
								}
							}.ToArray()
						}
					}));
					openXMLCommonSlideData.AppendChild(shapeTree);
					break;
				default: // slide
					openXMLCommonSlideData.AppendChild(shapeTree);
					break;
			}
		}
	}
}
