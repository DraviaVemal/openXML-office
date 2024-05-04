// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	/// Common Theme Class used by both presentation and spreadsheet.
	/// </summary>
	public class Theme
	{        // TODO : Understand the purpose and migrate it to right place Postition, sat, lum, shade, tint
		private readonly int?[][] gsLst1 = new int?[][]{
			new int?[]{0,105000,110000,null, 67000},
			new int?[]{50000, 103000,105000,null, 73000},
			new int?[]{100000,109000,105000,null, 81000}
		};
		private readonly int?[][] gsLst2 = new int?[][]{
			new int?[]{0, 103000,102000,null, 94000},
			new int?[]{50000,110000, 100000, 100000,null},
		   new int?[]{100000,120000,99000, 78000,null}
		   };
		private readonly int?[][] gsLst3 = new int?[][]{
			new int?[]{0,150000,102000, 98000,93000},
			new int?[]{50000, 130000 ,103000,90000,98000},
			new int?[]{100000, 120000 ,null,63000,null}
		};
		private readonly A.Theme openXMLTheme = new A.Theme();
		/// <summary>
		/// Creates a new theme with the given pallet.
		/// </summary>
		/// <param name="ThemePallet">
		/// </param>
		public Theme(ThemePallet ThemePallet = null)
		{
			CreateTheme(ThemePallet);
		}
		/// <summary>
		/// Returns the OpenXML Theme Object.
		/// </summary>
		/// <returns>
		/// </returns>
		public A.Theme GetTheme()
		{
			return openXMLTheme;
		}
		private static A.Outline AppendNodes(int width)
		{
			A.Outline outline = new A.Outline(
				new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
				new A.PresetDash { Val = A.PresetLineDashValues.Solid },
				new A.Miter { Limit = 800000 })
			{
				Width = width,
				CapType = A.LineCapValues.Flat,
				CompoundLineType = A.CompoundLineValues.Single,
				Alignment = A.PenAlignmentValues.Center
			};
			return outline;
		}
		private static A.Outline[] GenerateOutlines()
		{
			return new A.Outline[]{
				AppendNodes(6350),
				AppendNodes(12700),
				AppendNodes(19050)};
		}
		private void CreateTheme(ThemePallet ThemePallet)
		{
			openXMLTheme.Name = "Office Theme";
			openXMLTheme.ObjectDefaults = new A.ObjectDefaults();
			openXMLTheme.ThemeElements = new A.ThemeElements()
			{
				FontScheme = GenerateFontScheme(),
				FormatScheme = GenerateFormatScheme(),
				ColorScheme = new A.ColorScheme(
				   new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = ThemePallet.dark1 }),
				   new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = ThemePallet.light1 }),
				   new A.Dark2Color(new A.RgbColorModelHex { Val = ThemePallet.dark2 }),
				   new A.Light2Color(new A.RgbColorModelHex { Val = ThemePallet.light2 }),
				   new A.Accent1Color(new A.RgbColorModelHex { Val = ThemePallet.accent1 }),
				   new A.Accent2Color(new A.RgbColorModelHex { Val = ThemePallet.accent2 }),
				   new A.Accent3Color(new A.RgbColorModelHex { Val = ThemePallet.accent3 }),
				   new A.Accent4Color(new A.RgbColorModelHex { Val = ThemePallet.accent4 }),
				   new A.Accent5Color(new A.RgbColorModelHex { Val = ThemePallet.accent5 }),
				   new A.Accent6Color(new A.RgbColorModelHex { Val = ThemePallet.accent6 }),
				   new A.Hyperlink(new A.RgbColorModelHex { Val = ThemePallet.hyperlink }),
				   new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = ThemePallet.followedHyperlink })
				   )
				{
					Name = "Office"
				}
			};
		}
		private A.BackgroundFillStyleList GenerateBackgroundFillStyleList()
		{
			A.BackgroundFillStyleList backgroundFillStyleList = new A.BackgroundFillStyleList(new A.SolidFill()
			{
				SchemeColor = new A.SchemeColor { Val = A.SchemeColorValues.PhColor }
			}, new A.SolidFill()
			{
				SchemeColor = new A.SchemeColor(new A.Tint { Val = 95000 }, new A.SaturationModulation { Val = 170000 }) { Val = A.SchemeColorValues.PhColor }
			}, new A.GradientFill(
				new A.GradientStopList(gsLst3.Select(v => GetGradientStop(v[0], v[1], v[2], v[3], v[4])).ToList()),
				new A.LinearGradientFill()
				{
					Angle = 5400000,
					Scaled = false
				})
			{
				RotateWithShape = true,
			});
			return backgroundFillStyleList;
		}
		private static A.FontScheme GenerateFontScheme()
		{
			return new A.FontScheme()
			{
				Name = "OpenXMLOffice Fonts",
				MajorFont = new A.MajorFont(
					new A.LatinFont { Typeface = "Calibri Light", Panose = "020F0302020204030204" },
					new A.EastAsianFont { Typeface = "" },
					new A.ComplexScriptFont { Typeface = "" },
					new A.SupplementalFont { Script = "Jpan", Typeface = "游ゴシック Light" },
					new A.SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" },
					new A.SupplementalFont { Script = "Hans", Typeface = "等线 Light" },
					new A.SupplementalFont { Script = "Hant", Typeface = "新細明體" },
					new A.SupplementalFont { Script = "Arab", Typeface = "Times New Roman" },
					new A.SupplementalFont { Script = "Hebr", Typeface = "Times New Roman" },
					new A.SupplementalFont { Script = "Thai", Typeface = "Angsana New" },
					new A.SupplementalFont { Script = "Ethi", Typeface = "Nyala" },
					new A.SupplementalFont { Script = "Beng", Typeface = "Vrinda" },
					new A.SupplementalFont { Script = "Gujr", Typeface = "Shruti" },
					new A.SupplementalFont { Script = "Khmr", Typeface = "MoolBoran" },
					new A.SupplementalFont { Script = "Knda", Typeface = "Tunga" },
					new A.SupplementalFont { Script = "Guru", Typeface = "Raavi" },
					new A.SupplementalFont { Script = "Cans", Typeface = "Euphemia" },
					new A.SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" },
					new A.SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" },
					new A.SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" },
					new A.SupplementalFont { Script = "Thaa", Typeface = "MV Boli" },
					new A.SupplementalFont { Script = "Deva", Typeface = "Mangal" },
					new A.SupplementalFont { Script = "Telu", Typeface = "Gautami" },
					new A.SupplementalFont { Script = "Taml", Typeface = "Latha" },
					new A.SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Orya", Typeface = "Kalinga" },
					new A.SupplementalFont { Script = "Mlym", Typeface = "Kartika" },
					new A.SupplementalFont { Script = "Laoo", Typeface = "DokChampa" },
					new A.SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" },
					new A.SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" },
					new A.SupplementalFont { Script = "Viet", Typeface = "Times New Roman" },
					new A.SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" },
					new A.SupplementalFont { Script = "Geor", Typeface = "Sylfaen" },
					new A.SupplementalFont { Script = "Armn", Typeface = "Arial" },
					new A.SupplementalFont { Script = "Bugi", Typeface = "Leelawadee UI" },
					new A.SupplementalFont { Script = "Bopo", Typeface = "Microsoft JhengHei" },
					new A.SupplementalFont { Script = "Java", Typeface = "Javanese Text" },
					new A.SupplementalFont { Script = "Lisu", Typeface = "Segoe UI" },
					new A.SupplementalFont { Script = "Mymr", Typeface = "Myanmar Text" },
					new A.SupplementalFont { Script = "Nkoo", Typeface = "Ebrima" },
					new A.SupplementalFont { Script = "Olck", Typeface = "Nirmala UI" },
					new A.SupplementalFont { Script = "Osma", Typeface = "Ebrima" },
					new A.SupplementalFont { Script = "Phag", Typeface = "Phagspa" },
					new A.SupplementalFont { Script = "Syrn", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Syrj", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Syre", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Sora", Typeface = "Nirmala UI" },
					new A.SupplementalFont { Script = "Tale", Typeface = "Microsoft Tai Le" },
					new A.SupplementalFont { Script = "Talu", Typeface = "Microsoft New Tai Lue" },
					new A.SupplementalFont { Script = "Tfng", Typeface = "Ebrima" }
				),
				MinorFont = new A.MinorFont(
					new A.LatinFont { Typeface = "Calibri", Panose = "020F0502020204030204" },
					new A.EastAsianFont { Typeface = "" },
					new A.ComplexScriptFont { Typeface = "" },
					new A.SupplementalFont { Script = "Jpan", Typeface = "游ゴシック" },
					new A.SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" },
					new A.SupplementalFont { Script = "Hans", Typeface = "等线" },
					new A.SupplementalFont { Script = "Hant", Typeface = "新細明體" },
					new A.SupplementalFont { Script = "Arab", Typeface = "Arial" },
					new A.SupplementalFont { Script = "Hebr", Typeface = "Arial" },
					new A.SupplementalFont { Script = "Thai", Typeface = "Cordia New" },
					new A.SupplementalFont { Script = "Ethi", Typeface = "Nyala" },
					new A.SupplementalFont { Script = "Beng", Typeface = "Vrinda" },
					new A.SupplementalFont { Script = "Gujr", Typeface = "Shruti" },
					new A.SupplementalFont { Script = "Khmr", Typeface = "DaunPenh" },
					new A.SupplementalFont { Script = "Knda", Typeface = "Tunga" },
					new A.SupplementalFont { Script = "Guru", Typeface = "Raavi" },
					new A.SupplementalFont { Script = "Cans", Typeface = "Euphemia" },
					new A.SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" },
					new A.SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" },
					new A.SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" },
					new A.SupplementalFont { Script = "Thaa", Typeface = "MV Boli" },
					new A.SupplementalFont { Script = "Deva", Typeface = "Mangal" },
					new A.SupplementalFont { Script = "Telu", Typeface = "Gautami" },
					new A.SupplementalFont { Script = "Taml", Typeface = "Latha" },
					new A.SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Orya", Typeface = "Kalinga" },
					new A.SupplementalFont { Script = "Mlym", Typeface = "Kartika" },
					new A.SupplementalFont { Script = "Laoo", Typeface = "DokChampa" },
					new A.SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" },
					new A.SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" },
					new A.SupplementalFont { Script = "Viet", Typeface = "Arial" },
					new A.SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" },
					new A.SupplementalFont { Script = "Geor", Typeface = "Sylfaen" },
					new A.SupplementalFont { Script = "Armn", Typeface = "Arial" },
					new A.SupplementalFont { Script = "Bugi", Typeface = "Leelawadee UI" },
					new A.SupplementalFont { Script = "Bopo", Typeface = "Microsoft JhengHei" },
					new A.SupplementalFont { Script = "Java", Typeface = "Javanese Text" },
					new A.SupplementalFont { Script = "Lisu", Typeface = "Segoe UI" },
					new A.SupplementalFont { Script = "Mymr", Typeface = "Myanmar Text" },
					new A.SupplementalFont { Script = "Nkoo", Typeface = "Ebrima" },
					new A.SupplementalFont { Script = "Olck", Typeface = "Nirmala UI" },
					new A.SupplementalFont { Script = "Osma", Typeface = "Ebrima" },
					new A.SupplementalFont { Script = "Phag", Typeface = "Phagspa" },
					new A.SupplementalFont { Script = "Syrn", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Syrj", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Syre", Typeface = "Estrangelo Edessa" },
					new A.SupplementalFont { Script = "Sora", Typeface = "Nirmala UI" },
					new A.SupplementalFont { Script = "Tale", Typeface = "Microsoft Tai Le" },
					new A.SupplementalFont { Script = "Talu", Typeface = "Microsoft New Tai Lue" },
					new A.SupplementalFont { Script = "Tfng", Typeface = "Ebrima" }
				)
			};
		}
		private A.GradientFill GetGradientFill(int?[][] gsLst)
		{
			A.GradientFill gradientFill = new A.GradientFill(new A.LinearGradientFill()
			{
				Angle = 5400000,
				Scaled = false
			})
			{
				RotateWithShape = true,
				GradientStopList = new A.GradientStopList(
					gsLst.Select(v => GetGradientStop(v[0], v[1], v[2], v[3], v[4])).ToList()
				)
			};
			return gradientFill;
		}
		private A.FormatScheme GenerateFormatScheme()
		{
			return new A.FormatScheme()
			{
				Name = "Office",
				FillStyleList = new A.FillStyleList(
				new A.SolidFill()
				{
					SchemeColor = new A.SchemeColor { Val = A.SchemeColorValues.PhColor }
				},
				GetGradientFill(gsLst1),
				GetGradientFill(gsLst2)),
				LineStyleList = new A.LineStyleList(GenerateOutlines()),
				EffectStyleList = new A.EffectStyleList(new A.EffectStyle[]
					{
						new A.EffectStyle(new A.EffectList()),
						new A.EffectStyle(new A.EffectList()),
						new A.EffectStyle(
							new A.EffectList(
								new A.OuterShadow()
								{
									BlurRadius = 57150,
									Distance = 19050,
									Direction = 5400000,
									Alignment = A.RectangleAlignmentValues.Center,
									RotateWithShape = false,
									RgbColorModelHex = new A.RgbColorModelHex(new A.Alpha(){Val=63000}) { Val = "000000" }
								}
							)
						)
					}),
				BackgroundFillStyleList = GenerateBackgroundFillStyleList()
			};
		}
		private static A.GradientStop GetGradientStop(int? position, int? saturationModulation, int? luminanceModulation, int? shade, int? tint)
		{
			A.SchemeColor schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
			if (luminanceModulation != null)
			{
				schemeColor.AppendChild(new A.LuminanceModulation()
				{
					Val = luminanceModulation
				});
			}
			if (saturationModulation != null)
			{
				schemeColor.AppendChild(new A.SaturationModulation()
				{
					Val = saturationModulation
				});
			}
			if (shade != null)
			{
				schemeColor.AppendChild(new A.Shade()
				{
					Val = shade
				});
			}
			if (tint != null)
			{
				schemeColor.AppendChild(new A.Tint()
				{
					Val = tint
				});
			}
			return new A.GradientStop()
			{
				Position = position,
				SchemeColor = schemeColor
			};
		}
	}
}
