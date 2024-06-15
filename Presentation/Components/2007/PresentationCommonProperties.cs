// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Common Properties organized in one place to get inherited by child classes
	/// </summary>
	public class PresentationCommonProperties : CommonProperties
	{
		/// <summary>
		///
		/// </summary>
		protected P.Shape CreateShape<TextColorOption, ShapeTypeOptions>(ShapeModel<TextColorOption, ShapeTypeOptions> shapeModel)
		where TextColorOption : class, IColorOptions, new()
		where ShapeTypeOptions : class, IShapeTypeDetailsModel, new()
		{
			P.Shape shape = new P.Shape()
			{
				NonVisualShapeProperties = new P.NonVisualShapeProperties(
				new P.NonVisualDrawingProperties { Id = 2, Name = shapeModel.name },
				new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
				new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
				ShapeProperties = new P.ShapeProperties(
				new A.Transform2D(
					new A.Offset
					{
						X = (Int64Value)shapeModel.shapePropertiesModel.x,
						Y = (Int64Value)shapeModel.shapePropertiesModel.y
					},
					new A.Extents
					{
						Cx = (Int64Value)shapeModel.shapePropertiesModel.cx,
						Cy = (Int64Value)shapeModel.shapePropertiesModel.cy
					}
				),
				new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
				TextBody = new P.TextBody(new A.BodyProperties(),
				new A.ListStyle(),
				CreateDrawingParagraph(shapeModel.drawingParagraph))
			};
			return shape;
		}
	}
}
