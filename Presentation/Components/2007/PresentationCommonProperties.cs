// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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
		protected P.Shape CreateShape<TextColorOption>(ShapeModel<TextColorOption> shapeModel)
		where TextColorOption : class, IColorOptions, new()
		{
			P.Shape shape = new P.Shape()
			{
				NonVisualShapeProperties = new P.NonVisualShapeProperties(
				new P.NonVisualDrawingProperties { Id = 2, Name = shapeModel.Name },
				new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
				new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
				ShapeProperties = new P.ShapeProperties(
				new A.Transform2D(
					new A.Offset
					{
						X = shapeModel.shapePropertiesModel.X,
						Y = shapeModel.shapePropertiesModel.Y
					},
					new A.Extents
					{
						Cx = shapeModel.shapePropertiesModel.Cx,
						Cy = shapeModel.shapePropertiesModel.Cy
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
