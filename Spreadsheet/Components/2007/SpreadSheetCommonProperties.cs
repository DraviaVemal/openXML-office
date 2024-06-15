// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Common Properties organized in one place to get inherited by child classes
	/// </summary>
	public class SpreadSheetCommonProperties : CommonProperties
	{
		/// <summary>
		///
		/// </summary>
		protected XDR.Shape CreateShape<TextColorOption>(ShapeModel<TextColorOption> shapeModel)
		where TextColorOption : class, IColorOptions, new()
		{
			XDR.Shape shape = new XDR.Shape()
			{
				NonVisualShapeProperties = new XDR.NonVisualShapeProperties(
				new XDR.NonVisualDrawingProperties { Id = 2, Name = shapeModel.Name },
				new XDR.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
				ShapeProperties = new XDR.ShapeProperties(
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
				TextBody = new XDR.TextBody(new A.BodyProperties(),
				new A.ListStyle(),
				CreateDrawingParagraph(shapeModel.drawingParagraph))
			};
			return shape;
		}
	}
}
