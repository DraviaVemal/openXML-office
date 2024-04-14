// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using CS = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
namespace OpenXMLOffice.Global_2013
{
	/// <summary>
	/// Represents the chart style used for creating various chart elements.
	/// </summary>
	public class ChartStyle
	{
		/// <summary>
		///
		/// </summary>
		/// <returns></returns>
		public static CS.ChartStyle CreateChartStyles()
		{
			CS.ChartStyle ChartStyle = new CS.ChartStyle()
			{
				Id = 395,
				AxisTitle = CreateAxisTitle(),
				CategoryAxis = CreateCategoryAxis(),
				ChartArea = CreateChartArea(),
				DataLabel = CreateDataLabel(),
				DataLabelCallout = CreateDataLabelCallout(),
				DataPoint = CreateDataPoint(),
				DataPoint3D = CreateDataPoint3D(),
				DataPointLine = CreateDataPointLine(),
				DataPointMarker = CreateDataPointMarker(),
				MarkerLayoutProperties = CreateMarkerLayoutProperties(),
				DataPointWireframe = CreateDataPointWireframe(),
				DataTableStyle = CreateDataTableStyle(),
				DownBar = CreateDownBar(),
				DropLine = CreateDropLine(),
				ErrorBar = CreateErrorBar(),
				Floor = CreateFloor(),
				GridlineMajor = CreateGridlineMajor(),
				GridlineMinor = CreateGridlineMinor(),
				HiLoLine = CreateHiLoLine(),
				LeaderLine = CreateLeaderLine(),
				LegendStyle = CreateLegendStyle(),
				PlotArea = CreatePlotArea(),
				PlotArea3D = CreatePlotArea3D(),
				SeriesAxis = CreateSeriesAxis(),
				SeriesLine = CreateSeriesLine(),
				TitleStyle = CreateTitleStyle(),
				TrendlineStyle = CreateTrendlineStyle(),
				TrendlineLabel = CreateTrendlineLabel(),
				UpBar = CreateUpBar(),
				ValueAxis = CreateValueAxis(),
				Wall = CreateWall()
			};
			ChartStyle.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			return ChartStyle;
		}
		private static CS.AxisTitle CreateAxisTitle()
		{
			CS.AxisTitle axisTitle = new CS.AxisTitle();
			axisTitle.Append(new CS.LineReference { Index = (UInt32Value)0 });
			axisTitle.Append(new CS.FillReference { Index = (UInt32Value)0 });
			axisTitle.Append(new CS.EffectReference { Index = (UInt32Value)0 });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			axisTitle.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1330, Kerning = 1200 };
			axisTitle.Append(defRPr);
			return axisTitle;
		}
		private static CS.CategoryAxis CreateCategoryAxis()
		{
			CS.CategoryAxis categoryAxis = new CS.CategoryAxis();
			categoryAxis.Append(new CS.LineReference { Index = (UInt32Value)0U });
			categoryAxis.Append(new CS.FillReference { Index = (UInt32Value)0U });
			categoryAxis.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClrFont = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClrFont.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClrFont.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClrFont);
			categoryAxis.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill solidFill = new A.SolidFill();
			A.SchemeColor schemeClrLn = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClrLn.Append(new A.LuminanceModulation { Val = 15000 });
			schemeClrLn.Append(new A.LuminanceOffset { Val = 85000 });
			solidFill.Append(schemeClrLn);
			ln.Append(solidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			categoryAxis.Append(spPr);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };
			categoryAxis.Append(defRPr);
			return categoryAxis;
		}
		private static CS.ChartArea CreateChartArea()
		{
			CS.ChartArea chartArea = new CS.ChartArea();
			chartArea.Append(new CS.LineReference { Index = (UInt32Value)0U });
			chartArea.Append(new CS.FillReference { Index = (UInt32Value)0U });
			chartArea.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			chartArea.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.SolidFill solidFill = new A.SolidFill();
			solidFill.Append(new A.SchemeColor { Val = A.SchemeColorValues.Background1 });
			spPr.Append(solidFill);
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill();
			A.SchemeColor lnSchemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			lnSchemeClr.Append(new A.LuminanceModulation { Val = 15000 });
			lnSchemeClr.Append(new A.LuminanceOffset { Val = 85000 });
			lnSolidFill.Append(lnSchemeClr);
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			chartArea.Append(spPr);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1330, Kerning = 1200 };
			chartArea.Append(defRPr);
			return chartArea;
		}
		private static CS.DataLabel CreateDataLabel()
		{
			CS.DataLabel dataLabel = new CS.DataLabel();
			dataLabel.Append(new CS.LineReference { Index = (UInt32Value)0U });
			dataLabel.Append(new CS.FillReference { Index = (UInt32Value)0U });
			dataLabel.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 75000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 25000 });
			fontRef.Append(schemeClr);
			dataLabel.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };
			dataLabel.Append(defRPr);
			return dataLabel;
		}
		private static CS.DataLabelCallout CreateDataLabelCallout()
		{
			CS.DataLabelCallout dataLabelCallout = new CS.DataLabelCallout();
			dataLabelCallout.Append(new CS.LineReference { Index = (UInt32Value)0U });
			dataLabelCallout.Append(new CS.FillReference { Index = (UInt32Value)0U });
			dataLabelCallout.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			dataLabelCallout.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.SolidFill solidFill = new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 });
			spPr.Append(solidFill);
			A.Outline ln = new A.Outline();
			A.SolidFill lnSolidFill = new A.SolidFill();
			A.SchemeColor lnSchemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };
			lnSchemeClr.Append(new A.LuminanceModulation { Val = 25000 });
			lnSchemeClr.Append(new A.LuminanceOffset { Val = 75000 });
			lnSolidFill.Append(lnSchemeClr);
			ln.Append(lnSolidFill);
			spPr.Append(ln);
			dataLabelCallout.Append(spPr);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };
			dataLabelCallout.Append(defRPr);
			CS.TextBodyProperties bodyPr = new CS.TextBodyProperties()
			{
				Rotation = 0,
				UseParagraphSpacing = true,
				VerticalOverflow = A.TextVerticalOverflowValues.Clip,
				HorizontalOverflow = A.TextHorizontalOverflowValues.Clip,
				Vertical = A.TextVerticalValues.Horizontal,
				Wrap = A.TextWrappingValues.Square,
				LeftInset = 36576,
				TopInset = 18288,
				RightInset = 36576,
				BottomInset = 18288,
				Anchor = A.TextAnchoringTypeValues.Center,
				AnchorCenter = true
			};
			bodyPr.Append(new A.ShapeAutoFit());
			dataLabelCallout.Append(bodyPr);
			return dataLabelCallout;
		}
		private static CS.DataPoint CreateDataPoint()
		{
			CS.DataPoint dataPoint = new CS.DataPoint();
			dataPoint.Append(new CS.LineReference { Index = (UInt32Value)0U });
			CS.FillReference fillRef = new CS.FillReference() { Index = (UInt32Value)1U };
			fillRef.Append(new CS.StyleColor { Val = "auto" });
			dataPoint.Append(fillRef);
			dataPoint.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			dataPoint.Append(fontRef);
			return dataPoint;
		}
		private static CS.DataPoint3D CreateDataPoint3D()
		{
			CS.DataPoint3D dataPoint3D = new CS.DataPoint3D();
			dataPoint3D.Append(new CS.LineReference { Index = (UInt32Value)0U });
			CS.FillReference fillRef = new CS.FillReference() { Index = (UInt32Value)1U };
			fillRef.Append(new CS.StyleColor { Val = "auto" });
			dataPoint3D.Append(fillRef);
			dataPoint3D.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			dataPoint3D.Append(fontRef);
			return dataPoint3D;
		}
		private static CS.DataPointLine CreateDataPointLine()
		{
			CS.DataPointLine dataPointLine = new CS.DataPointLine();
			CS.LineReference lnRef = new CS.LineReference() { Index = (UInt32Value)0U };
			lnRef.Append(new CS.StyleColor { Val = "auto" });
			dataPointLine.Append(lnRef);
			dataPointLine.Append(new CS.FillReference { Index = (UInt32Value)1U });
			dataPointLine.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			dataPointLine.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };
			A.SolidFill solidFill = new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
			ln.Append(solidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			dataPointLine.Append(spPr);
			return dataPointLine;
		}
		private static CS.DataPointMarker CreateDataPointMarker()
		{
			CS.DataPointMarker dataPointMarker = new CS.DataPointMarker();
			CS.LineReference lnRef = new CS.LineReference() { Index = (UInt32Value)0U };
			lnRef.Append(new CS.StyleColor { Val = "auto" });
			dataPointMarker.Append(lnRef);
			CS.FillReference fillRef = new CS.FillReference() { Index = (UInt32Value)1U };
			fillRef.Append(new CS.StyleColor { Val = "auto" });
			dataPointMarker.Append(fillRef);
			dataPointMarker.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			dataPointMarker.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525 };
			A.SolidFill solidFill = new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
			ln.Append(solidFill);
			spPr.Append(ln);
			dataPointMarker.Append(spPr);
			return dataPointMarker;
		}
		private static CS.DataPointWireframe CreateDataPointWireframe()
		{
			return new CS.DataPointWireframe(new CS.LineReference(
				new CS.StyleColor()
				{
					Val = "auto"
				})
			{ Index = 0 },
			new CS.FillReference()
			{
				Index = 1
			}, new CS.EffectReference()
			{
				Index = 1
			}, new CS.FontReference(
				new A.SchemeColor()
				{
					Val = A.SchemeColorValues.Text1
				}
			)
			{
				Index = A.FontCollectionIndexValues.Minor
			}, new CS.ShapeProperties(
				new A.Outline(
					new A.SolidFill(new A.SchemeColor()
					{
						Val = A.SchemeColorValues.PhColor
					}),
					new A.Round()
				)
				{
					Width = 9525,
					CapType = A.LineCapValues.Round
				}
			));
		}
		private static CS.DataTableStyle CreateDataTableStyle()
		{
			CS.DataTableStyle dataTableStyle = new CS.DataTableStyle();
			dataTableStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			dataTableStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			dataTableStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClrFont = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClrFont.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClrFont.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClrFont);
			dataTableStyle.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill();
			A.SchemeColor lnSchemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			lnSchemeClr.Append(new A.LuminanceModulation { Val = 15000 });
			lnSchemeClr.Append(new A.LuminanceOffset { Val = 85000 });
			lnSolidFill.Append(lnSchemeClr);
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(new A.NoFill());
			spPr.Append(ln);
			dataTableStyle.Append(spPr);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };
			dataTableStyle.Append(defRPr);
			return dataTableStyle;
		}
		private static CS.DownBar CreateDownBar()
		{
			CS.DownBar downBar = new CS.DownBar();
			downBar.Append(new CS.LineReference { Index = (UInt32Value)0U });
			downBar.Append(new CS.FillReference { Index = (UInt32Value)0U });
			downBar.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Dark1 });
			downBar.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.SolidFill solidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 65000 },
				new A.LuminanceOffset { Val = 35000 })
			{
				Val = A.SchemeColorValues.Dark1
			});
			spPr.Append(solidFill);
			A.Outline ln = new A.Outline() { Width = 9525 };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 65000 },
				new A.LuminanceOffset { Val = 35000 })
			{
				Val = A.SchemeColorValues.Text1
			});
			ln.Append(lnSolidFill);
			spPr.Append(ln);
			downBar.Append(spPr);
			return downBar;
		}
		private static CS.DropLine CreateDropLine()
		{
			CS.DropLine dropLine = new CS.DropLine();
			dropLine.Append(new CS.LineReference { Index = (UInt32Value)0U });
			dropLine.Append(new CS.FillReference { Index = (UInt32Value)0U });
			dropLine.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			dropLine.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 35000 },
				new A.LuminanceOffset { Val = 65000 })
			{
				Val = A.SchemeColorValues.Text1
			});
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			dropLine.Append(spPr);
			return dropLine;
		}
		private static CS.ErrorBar CreateErrorBar()
		{
			CS.ErrorBar errorBar = new CS.ErrorBar();
			errorBar.Append(new CS.LineReference { Index = (UInt32Value)0U });
			errorBar.Append(new CS.FillReference { Index = (UInt32Value)0U });
			errorBar.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			errorBar.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 65000 },
				new A.LuminanceOffset { Val = 35000 })
			{
				Val = A.SchemeColorValues.Text1
			});
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			errorBar.Append(spPr);
			return errorBar;
		}
		private static CS.Floor CreateFloor()
		{
			CS.Floor floor = new CS.Floor();
			floor.Append(new CS.LineReference { Index = (UInt32Value)0U });
			floor.Append(new CS.FillReference { Index = (UInt32Value)0U });
			floor.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			floor.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			spPr.Append(new A.NoFill());
			A.Outline ln = new A.Outline();
			ln.Append(new A.NoFill());
			spPr.Append(ln);
			floor.Append(spPr);
			return floor;
		}
		private static CS.GridlineMajor CreateGridlineMajor()
		{
			CS.GridlineMajor gridlineMajor = new CS.GridlineMajor();
			gridlineMajor.Append(new CS.LineReference { Index = (UInt32Value)0U });
			gridlineMajor.Append(new CS.FillReference { Index = (UInt32Value)0U });
			gridlineMajor.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			gridlineMajor.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 15000 },
				new A.LuminanceOffset { Val = 85000 })
			{ Val = A.SchemeColorValues.Text1 });
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			gridlineMajor.Append(spPr);
			return gridlineMajor;
		}
		private static CS.GridlineMinor CreateGridlineMinor()
		{
			CS.GridlineMinor gridlineMinor = new CS.GridlineMinor();
			gridlineMinor.Append(new CS.LineReference { Index = (UInt32Value)0U });
			gridlineMinor.Append(new CS.FillReference { Index = (UInt32Value)0U });
			gridlineMinor.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			gridlineMinor.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 5000 },
				new A.LuminanceOffset { Val = 95000 })
			{ Val = A.SchemeColorValues.Text1 });
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			gridlineMinor.Append(spPr);
			return gridlineMinor;
		}
		private static CS.HiLoLine CreateHiLoLine()
		{
			CS.HiLoLine hiLoLine = new CS.HiLoLine();
			hiLoLine.Append(new CS.LineReference { Index = (UInt32Value)0U });
			hiLoLine.Append(new CS.FillReference { Index = (UInt32Value)0U });
			hiLoLine.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			hiLoLine.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(
				new A.SchemeColor(new A.LuminanceModulation { Val = 75000 },
				new A.LuminanceOffset { Val = 25000 })
				{ Val = A.SchemeColorValues.Text1 });
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			hiLoLine.Append(spPr);
			return hiLoLine;
		}
		private static CS.LeaderLine CreateLeaderLine()
		{
			CS.LeaderLine leaderLine = new CS.LeaderLine();
			leaderLine.Append(new CS.LineReference { Index = (UInt32Value)0U });
			leaderLine.Append(new CS.FillReference { Index = (UInt32Value)0U });
			leaderLine.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			leaderLine.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 35000 },
			new A.LuminanceOffset { Val = 65000 })
			{ Val = A.SchemeColorValues.Text1 });
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			leaderLine.Append(spPr);
			return leaderLine;
		}
		private static CS.LegendStyle CreateLegendStyle()
		{
			CS.LegendStyle legendStyle = new CS.LegendStyle();
			legendStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			legendStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			legendStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			legendStyle.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };
			legendStyle.Append(defRPr);
			return legendStyle;
		}
		private static CS.MarkerLayoutProperties CreateMarkerLayoutProperties()
		{
			return new CS.MarkerLayoutProperties()
			{
				Size = 5,
				Symbol = CS.MarkerStyle.Circle
			};
		}
		private static CS.PlotArea CreatePlotArea()
		{
			CS.PlotArea plotAreaStyle = new CS.PlotArea();
			plotAreaStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			plotAreaStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			plotAreaStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			plotAreaStyle.Append(fontRef);
			return plotAreaStyle;
		}
		private static CS.PlotArea3D CreatePlotArea3D()
		{
			CS.PlotArea3D plotArea3DStyle = new CS.PlotArea3D();
			plotArea3DStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			plotArea3DStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			plotArea3DStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			plotArea3DStyle.Append(fontRef);
			return plotArea3DStyle;
		}
		private static CS.SeriesAxis CreateSeriesAxis()
		{
			CS.SeriesAxis seriesAxisStyle = new CS.SeriesAxis();
			seriesAxisStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			seriesAxisStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			seriesAxisStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			seriesAxisStyle.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType() { FontSize = 1197, Kerning = 1200 };
			seriesAxisStyle.Append(defRPr);
			return seriesAxisStyle;
		}
		private static CS.SeriesLine CreateSeriesLine()
		{
			CS.SeriesLine seriesLineStyle = new CS.SeriesLine();
			seriesLineStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			seriesLineStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			seriesLineStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			seriesLineStyle.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 35000 },
			new A.LuminanceOffset { Val = 65000 })
			{
				Val = A.SchemeColorValues.Text1
			});
			ln.Append(lnSolidFill);
			ln.Append(new A.Round());
			spPr.Append(ln);
			seriesLineStyle.Append(spPr);
			return seriesLineStyle;
		}
		private static CS.TitleStyle CreateTitleStyle()
		{
			CS.TitleStyle titleStyle = new CS.TitleStyle();
			titleStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			titleStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			titleStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			titleStyle.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType()
			{
				FontSize = 1862,
				Bold = false,
				Kerning = 1200,
				Spacing = 0,
				Baseline = 0
			};
			titleStyle.Append(defRPr);
			return titleStyle;
		}
		private static CS.TrendlineLabel CreateTrendlineLabel()
		{
			CS.TrendlineLabel trendlineLabelStyle = new CS.TrendlineLabel();
			trendlineLabelStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			trendlineLabelStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			trendlineLabelStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			trendlineLabelStyle.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType()
			{
				FontSize = 1197,
				Kerning = 1200
			};
			trendlineLabelStyle.Append(defRPr);
			return trendlineLabelStyle;
		}
		private static CS.TrendlineStyle CreateTrendlineStyle()
		{
			CS.TrendlineStyle trendlineStyle = new CS.TrendlineStyle();
			CS.LineReference lnRef = new CS.LineReference() { Index = (UInt32Value)0U };
			lnRef.Append(new CS.StyleColor { Val = "auto" });
			trendlineStyle.Append(lnRef);
			trendlineStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			trendlineStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			trendlineStyle.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.Outline ln = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Round };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
			ln.Append(lnSolidFill);
			ln.Append(new A.PresetDash { Val = A.PresetLineDashValues.SystemDot });
			spPr.Append(ln);
			trendlineStyle.Append(spPr);
			return trendlineStyle;
		}
		private static CS.UpBar CreateUpBar()
		{
			CS.UpBar upBarStyle = new CS.UpBar();
			upBarStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			upBarStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			upBarStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Dark1 });
			upBarStyle.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			A.SolidFill solidFill = new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 });
			spPr.Append(solidFill);
			A.Outline ln = new A.Outline() { Width = 9525 };
			A.SolidFill lnSolidFill = new A.SolidFill(new A.SchemeColor(
				new A.LuminanceModulation { Val = 15000 },
				new A.LuminanceOffset { Val = 85000 })
			{ Val = A.SchemeColorValues.Text1 });
			ln.Append(lnSolidFill);
			spPr.Append(ln);
			upBarStyle.Append(spPr);
			return upBarStyle;
		}
		private static CS.ValueAxis CreateValueAxis()
		{
			CS.ValueAxis valueAxisStyle = new CS.ValueAxis();
			valueAxisStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			valueAxisStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			valueAxisStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			A.SchemeColor schemeClr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
			schemeClr.Append(new A.LuminanceModulation { Val = 65000 });
			schemeClr.Append(new A.LuminanceOffset { Val = 35000 });
			fontRef.Append(schemeClr);
			valueAxisStyle.Append(fontRef);
			CS.TextCharacterPropertiesType defRPr = new CS.TextCharacterPropertiesType()
			{
				FontSize = 1197,
				Kerning = 1200
			};
			valueAxisStyle.Append(defRPr);
			return valueAxisStyle;
		}
		private static CS.Wall CreateWall()
		{
			CS.Wall wallStyle = new CS.Wall();
			wallStyle.Append(new CS.LineReference { Index = (UInt32Value)0U });
			wallStyle.Append(new CS.FillReference { Index = (UInt32Value)0U });
			wallStyle.Append(new CS.EffectReference { Index = (UInt32Value)0U });
			CS.FontReference fontRef = new CS.FontReference() { Index = A.FontCollectionIndexValues.Minor };
			fontRef.Append(new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
			wallStyle.Append(fontRef);
			CS.ShapeProperties spPr = new CS.ShapeProperties();
			spPr.Append(new A.NoFill());
			A.Outline ln = new A.Outline();
			ln.Append(new A.NoFill());
			spPr.Append(ln);
			wallStyle.Append(spPr);
			return wallStyle;
		}
	}
}
