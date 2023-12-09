using DocumentFormat.OpenXml.Drawing;

namespace OpenXMLOffice.Presentation;
public class Theme : Global
{
    // TODO : Understand the purpose and migrate it to right place
    private int[][] gsLst1 = new int[][]{
        new int[] { 0,110000, 105000, 67000 },
        new int[] {50000, 105000, 103000, 73000 },
        new int[] { 100000,105000, 109000, 81000 }
    };
    private int[][] gsLst2 = new int[][]{
        new int[] { 0,103000, 102000, 94000 },
        new int[] { 50000,110000, 100000, 100000 },
        new int[] { 100000,99000, 120000, 78000 }
    };
    private int[][] gsLst3 = new int[][]{
        new int[] { 0,93000, 150000, 98000,102000 },
        new int[] { 50000,98000, 130000, 90000 ,103000}
    };
    private FontScheme GenerateFontScheme()
    {
        return new FontScheme()
        {
            Name = "OpenXMLOffice Fonts",
            MajorFont = new MajorFont(
                new LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" },
                new EastAsianFont() { Typeface = "" },
                new ComplexScriptFont() { Typeface = "" },
                new SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" },
                new SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" },
                new SupplementalFont() { Script = "Hans", Typeface = "等线 Light" },
                new SupplementalFont() { Script = "Hant", Typeface = "新細明體" },
                new SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" },
                new SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" },
                new SupplementalFont() { Script = "Thai", Typeface = "Angsana New" },
                new SupplementalFont() { Script = "Ethi", Typeface = "Nyala" },
                new SupplementalFont() { Script = "Beng", Typeface = "Vrinda" },
                new SupplementalFont() { Script = "Gujr", Typeface = "Shruti" },
                new SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" },
                new SupplementalFont() { Script = "Knda", Typeface = "Tunga" },
                new SupplementalFont() { Script = "Guru", Typeface = "Raavi" },
                new SupplementalFont() { Script = "Cans", Typeface = "Euphemia" },
                new SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" },
                new SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" },
                new SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" },
                new SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" },
                new SupplementalFont() { Script = "Deva", Typeface = "Mangal" },
                new SupplementalFont() { Script = "Telu", Typeface = "Gautami" },
                new SupplementalFont() { Script = "Taml", Typeface = "Latha" },
                new SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Orya", Typeface = "Kalinga" },
                new SupplementalFont() { Script = "Mlym", Typeface = "Kartika" },
                new SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" },
                new SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" },
                new SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" },
                new SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" },
                new SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" },
                new SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" },
                new SupplementalFont() { Script = "Armn", Typeface = "Arial" },
                new SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" },
                new SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" },
                new SupplementalFont() { Script = "Java", Typeface = "Javanese Text" },
                new SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" },
                new SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" },
                new SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" },
                new SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" },
                new SupplementalFont() { Script = "Osma", Typeface = "Ebrima" },
                new SupplementalFont() { Script = "Phag", Typeface = "Phagspa" },
                new SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" },
                new SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" },
                new SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" },
                new SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" }
            ),
            MinorFont = new MinorFont(
                new LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" },
                new EastAsianFont() { Typeface = "" },
                new ComplexScriptFont() { Typeface = "" },
                new SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" },
                new SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" },
                new SupplementalFont() { Script = "Hans", Typeface = "等线" },
                new SupplementalFont() { Script = "Hant", Typeface = "新細明體" },
                new SupplementalFont() { Script = "Arab", Typeface = "Arial" },
                new SupplementalFont() { Script = "Hebr", Typeface = "Arial" },
                new SupplementalFont() { Script = "Thai", Typeface = "Cordia New" },
                new SupplementalFont() { Script = "Ethi", Typeface = "Nyala" },
                new SupplementalFont() { Script = "Beng", Typeface = "Vrinda" },
                new SupplementalFont() { Script = "Gujr", Typeface = "Shruti" },
                new SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" },
                new SupplementalFont() { Script = "Knda", Typeface = "Tunga" },
                new SupplementalFont() { Script = "Guru", Typeface = "Raavi" },
                new SupplementalFont() { Script = "Cans", Typeface = "Euphemia" },
                new SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" },
                new SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" },
                new SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" },
                new SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" },
                new SupplementalFont() { Script = "Deva", Typeface = "Mangal" },
                new SupplementalFont() { Script = "Telu", Typeface = "Gautami" },
                new SupplementalFont() { Script = "Taml", Typeface = "Latha" },
                new SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Orya", Typeface = "Kalinga" },
                new SupplementalFont() { Script = "Mlym", Typeface = "Kartika" },
                new SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" },
                new SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" },
                new SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" },
                new SupplementalFont() { Script = "Viet", Typeface = "Arial" },
                new SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" },
                new SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" },
                new SupplementalFont() { Script = "Armn", Typeface = "Arial" },
                new SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" },
                new SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" },
                new SupplementalFont() { Script = "Java", Typeface = "Javanese Text" },
                new SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" },
                new SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" },
                new SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" },
                new SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" },
                new SupplementalFont() { Script = "Osma", Typeface = "Ebrima" },
                new SupplementalFont() { Script = "Phag", Typeface = "Phagspa" },
                new SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" },
                new SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" },
                new SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" },
                new SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" },
                new SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" }
            )
        };
    }
    private FormatScheme GenerateFormatScheme()
    {
        return new FormatScheme()
        {
            Name = "OpenXMLOffice Formats",
            FillStyleList = new FillStyleList(
            new SolidFill()
            {
                SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
            },
            GetGradientFill(gsLst1),
            GetGradientFill(gsLst2)),
            LineStyleList = new LineStyleList(GenerateOutlines()),
            EffectStyleList = new EffectStyleList(new EffectStyle[]
                {
                    new(new EffectList()),
                    new(new EffectList()),
                    new(
                        new EffectList(
                            new OuterShadow()
                            {
                                BlurRadius = 57150,
                                Distance = 19050,
                                Direction = 5400000,
                                Alignment = RectangleAlignmentValues.Center,
                                RotateWithShape = false,
                                RgbColorModelHex = new RgbColorModelHex(new Alpha(){Val=63000}) { Val = "000000" }
                            }
                        )
                    )
                }),
            BackgroundFillStyleList = GenerateBackgroundFillStyleList()
        };
        GradientFill GetGradientFill(int[][] gsLst)
        {
            GradientStop GetGradientStop(int position, int luminanceModulation, int saturationModulation, int tint)
            {
                return new()
                {
                    Position = position,
                    SchemeColor = new(new LuminanceModulation()
                    {
                        Val = luminanceModulation
                    }, new SaturationModulation()
                    {
                        Val = saturationModulation
                    }, new Tint()
                    {
                        Val = tint
                    })
                    { Val = SchemeColorValues.PhColor }
                }; ;
            }
            GradientFill gradientFill = new(new LinearGradientFill()
            {
                Angle = 5400000,
                Scaled = false
            })
            {
                RotateWithShape = true,
                GradientStopList = new GradientStopList(
                    gsLst.Select(v => GetGradientStop(v[0], v[1], v[2], v[3])).ToList()
                )
            };
            return gradientFill;
        }
    }

    private BackgroundFillStyleList GenerateBackgroundFillStyleList()
    {
        GradientStop GetGradientStop(int position, int tint, int saturationModulation, int shade, int luminanceModulation)
        {
            return new()
            {
                Position = position,
                SchemeColor = new(new Tint()
                {
                    Val = tint
                }, new SaturationModulation()
                {
                    Val = saturationModulation
                }, new Shade()
                {
                    Val = shade
                }, new LuminanceModulation()
                {
                    Val = luminanceModulation
                })
                { Val = SchemeColorValues.PhColor }
            };
        };
        GradientStopList gradientStopList = new(gsLst3.Select(v => GetGradientStop(v[0], v[1], v[2], v[3], v[4])).ToList());
        gradientStopList.AppendChild(new GradientStop(new SchemeColor(
            new Shade()
            {
                Val = 63000
            },
            new SaturationModulation()
            {
                Val = 120000
            }
        )
        {
            Val = SchemeColorValues.PhColor
        })
        {
            Position = 100000
        });
        BackgroundFillStyleList backgroundFillStyleList = new(new SolidFill()
        {
            SchemeColor = new SchemeColor() { Val = SchemeColorValues.PhColor }
        }, new SolidFill()
        {
            SchemeColor = new(new Tint() { Val = 95000 }, new SaturationModulation() { Val = 170000 }) { Val = SchemeColorValues.PhColor }
        }, new GradientFill(
            gradientStopList,
            new LinearGradientFill()
            {
                Angle = 5400000,
                Scaled = false
            })
        {
            RotateWithShape = true,
        });
        return backgroundFillStyleList;
    }

    private static Outline[] GenerateOutlines()
    {
        Outline AppendNodes(int width)
        {
            Outline outline = new(
                new SolidFill(new SchemeColor() { Val = SchemeColorValues.PhColor }),
                new PresetDash() { Val = PresetLineDashValues.Solid },
                new Miter() { Limit = 800000 })
            {
                Width = width,
                CapType = LineCapValues.Flat,
                CompoundLineType = CompoundLineValues.Single,
                Alignment = PenAlignmentValues.Center
            };
            return outline;
        }
        return new Outline[]{
            AppendNodes(6350),
            AppendNodes(12700),
            AppendNodes(19050)};
    }
    protected DocumentFormat.OpenXml.Drawing.Theme CreateTheme(PowerPointTheme? powerPointTheme)
    {
        return new()
        {
            Name = "OpenXMLOffice Theme",
            ObjectDefaults = new ObjectDefaults(),
            ThemeElements = new ThemeElements()
            {
                FontScheme = GenerateFontScheme(),
                FormatScheme = GenerateFormatScheme(),
                ColorScheme = new ColorScheme(
               new Dark1Color(new SystemColor() { Val = SystemColorValues.WindowText, LastColor = powerPointTheme?.Dark1 ?? "000000" }),
               new Light1Color(new SystemColor() { Val = SystemColorValues.Window, LastColor = powerPointTheme?.Dark1 ?? "FFFFFF" }),
               new Dark2Color(new RgbColorModelHex() { Val = powerPointTheme?.Dark2 ?? "44546A" }),
               new Light2Color(new RgbColorModelHex() { Val = powerPointTheme?.Light2 ?? "E7E6E6" }),
               new Accent1Color(new RgbColorModelHex() { Val = powerPointTheme?.Accent1 ?? "4472C4" }),
               new Accent2Color(new RgbColorModelHex() { Val = powerPointTheme?.Accent2 ?? "ED7D31" }),
               new Accent3Color(new RgbColorModelHex() { Val = powerPointTheme?.Accent3 ?? "A5A5A5" }),
               new Accent4Color(new RgbColorModelHex() { Val = powerPointTheme?.Accent4 ?? "FFC000" }),
               new Accent5Color(new RgbColorModelHex() { Val = powerPointTheme?.Accent5 ?? "5B9BD5" }),
               new Accent6Color(new RgbColorModelHex() { Val = powerPointTheme?.Accent6 ?? "70AD47" }),
               new Hyperlink(new RgbColorModelHex() { Val = powerPointTheme?.Hyperlink ?? "0563C1" }),
               new FollowedHyperlinkColor(new RgbColorModelHex() { Val = powerPointTheme?.FollowedHyperlink ?? "954F72" })
               )
                {
                    Name = "OpenXMLOffice Color Scheme"
                }
            }
        };
    }
}
