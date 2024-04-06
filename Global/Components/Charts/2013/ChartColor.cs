// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using A = DocumentFormat.OpenXml.Drawing;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;

namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents a class for creating color styles for charts.
	/// </summary>
	public class ChartColor
	{

		/// <summary>
		/// Creates the color styles for charts.
		/// </summary>
		/// <returns>
		/// The color style object.
		/// </returns>
		public static CS.ColorStyle CreateColorStyles()
		{
			CS.ColorStyle colorStyle = new() { Method = "cycle", Id = 10 };
			colorStyle.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			for (int i = 1; i < 7; i++)
			{
				colorStyle.Append(new A.SchemeColor()
				{
					Val = new A.SchemeColorValues($"accent{i}")
				});
			}
			colorStyle.Append(new CS.ColorStyleVariation());
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 60000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 80000
			}, new A.LuminanceOffset()
			{
				Val = 20000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 80000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 60000
			}, new A.LuminanceOffset()
			{
				Val = 40000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 50000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 70000
			}, new A.LuminanceOffset()
			{
				Val = 30000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 70000
			}));
			colorStyle.Append(new CS.ColorStyleVariation(new A.LuminanceModulation()
			{
				Val = 50000
			}, new A.LuminanceOffset()
			{
				Val = 50000
			}));
			return colorStyle;
		}


	}
}
