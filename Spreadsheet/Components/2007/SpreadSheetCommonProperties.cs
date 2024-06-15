// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
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
		protected XDR.Shape CreateShape<TextColorOption, ShapeTypeOptions>(ShapeModel<TextColorOption, ShapeTypeOptions> shapeModel)
		where TextColorOption : class, IColorOptions, new()
		where ShapeTypeOptions : class, IShapeTypeDetailsModel, new()
		{
			XDR.Shape shape = new XDR.Shape()
			{
				NonVisualShapeProperties = new XDR.NonVisualShapeProperties(
				new XDR.NonVisualDrawingProperties { Id = 2, Name = shapeModel.name },
				new XDR.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
				ShapeProperties = new XDR.ShapeProperties(
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
				TextBody = new XDR.TextBody(new A.BodyProperties(),
				new A.ListStyle(),
				CreateDrawingParagraph(shapeModel.drawingParagraph))
			};
			return shape;
		}
	}
}
