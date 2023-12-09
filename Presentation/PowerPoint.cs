using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class PowerPoint
{

    private readonly PresentationDocument presentationDocument;
    private PresentationPart? presentationPart;
    private SlideMasterPart? slideMasterPart;
    private SlideLayoutPart? slideLayoutPart;
    public PowerPoint(string filePath, PowerPointProperties? powerPointProperties = null, bool isEditable = true, bool autosave = true)
    {
        presentationDocument = PresentationDocument.Open(filePath, isEditable, new OpenSettings()
        {
            AutoSave = autosave
        });
        PreparePresentation(powerPointProperties);
    }
    public PowerPoint(string filePath, PresentationDocumentType presentationDocumentType, PowerPointProperties? powerPointProperties = null, bool autosave = true)
    {
        presentationDocument = PresentationDocument.Create(filePath, presentationDocumentType, autosave);
        PreparePresentation(powerPointProperties);
    }
    public PowerPoint(Stream stream, PresentationDocumentType presentationDocumentType, PowerPointProperties? powerPointProperties = null, bool autosave = true)
    {
        presentationDocument = PresentationDocument.Create(stream, presentationDocumentType, autosave);
        PreparePresentation(powerPointProperties);
    }
    /// <summary>
    /// Initializes a new instance of the PowerPoint class and creates a new PowerPoint presentation using a template file specified by the filePath parameter.
    /// </summary>
    /// <param name="filePath">The file path to the template file for creating the PowerPoint presentation.</param>
    public PowerPoint(string filePath, PowerPointProperties? powerPointProperties = null)
    {
        presentationDocument = PresentationDocument.CreateFromTemplate(filePath);
        PreparePresentation(powerPointProperties);
    }
    private void PreparePresentation(PowerPointProperties? powerPointProperties)
    {
        presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();
        presentationPart.Presentation ??= new P.Presentation();
        presentationPart.Presentation.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        if (presentationPart.Presentation.GetFirstChild<SlideMasterIdList>() == null)
        {
            presentationPart.Presentation.AddChild(new SlideMasterIdList());
        }
        if (presentationPart.Presentation.SlideIdList == null)
        {
            presentationPart.Presentation.AddChild(new SlideIdList());
        }
        if (presentationPart.Presentation.GetFirstChild<SlideSize>() == null)
        {
            presentationPart.Presentation.AddChild(new SlideSize() { Cx = 12192000, Cy = 6858000 });
        }
        if (presentationPart.Presentation.GetFirstChild<NotesSize>() == null)
        {
            presentationPart.Presentation.AddChild(new NotesSize() { Cx = 6858000, Cy = 9144000 });
        }
        if (presentationPart.ThemePart == null)
        {
            ThemePart themePart = presentationPart.AddNewPart<ThemePart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            themePart.Theme = CreateTheme(powerPointProperties?.Theme);
        }
        else
        {
            presentationPart.ThemePart.Theme = CreateTheme(powerPointProperties?.Theme);
        }
        if (!presentationPart.SlideMasterParts.Any())
        {
            slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            slideMasterPart.SlideMaster = new SlideMaster(CreateCommonSlideData());
            SlideMasterIdList slideMasterIdList = presentationPart.Presentation.SlideMasterIdList!;
            SlideMasterId slideMasterId = new() { Id = (uint)(2147483647 + slideMasterIdList.Count() + 1), RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) };
            slideMasterIdList.Append(slideMasterId);
            slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            slideLayoutPart.SlideLayout = new SlideLayout(CreateCommonSlideData());
            slideMasterPart.SlideMaster.AddChild(new ColorMap()
            {
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            });
            slideLayoutPart.SlideLayout.Save();
            slideMasterPart.SlideMaster.Save();
        }
        presentationPart.Presentation.Save();
    }

    private A.Theme CreateTheme(PowerPointTheme? powerPointTheme)
    {
        return new A.Theme
        {
            Name = "OpenXMLOffice Theme",
            ThemeElements = new A.ThemeElements()
            {
                ColorScheme = new A.ColorScheme(
               new A.Dark1Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Dark1 ?? "000000" }),
               new A.Light1Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Light1 ?? "FFFFFF" }),
               new A.Dark2Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Dark2 ?? "44546A" }),
               new A.Light2Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Light2 ?? "E7E6E6" }),
               new A.Accent1Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Accent1 ?? "4472C4" }),
               new A.Accent2Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Accent2 ?? "ED7D31" }),
               new A.Accent3Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Accent3 ?? "A5A5A5" }),
               new A.Accent4Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Accent4 ?? "FFC000" }),
               new A.Accent5Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Accent5 ?? "5B9BD5" }),
               new A.Accent6Color(new A.RgbColorModelHex() { Val = powerPointTheme?.Accent6 ?? "70AD47" }),
               new A.Hyperlink(new A.RgbColorModelHex() { Val = powerPointTheme?.Hyperlink ?? "0563C1" }),
               new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = powerPointTheme?.FollowedHyperlink ?? "954F72" })
               )
                {
                    Name = "OpenXMLOffice Color Scheme"
                }
            }
        };
    }

    private CommonSlideData CreateCommonSlideData()
    {
        return new CommonSlideData(new ShapeTree()
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
                            new NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()
                        )
        });
    }

    public void AddBlankSlide()
    {
        SlidePart slidePart = presentationPart!.AddNewPart<SlidePart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
        Slide slide = new()
        {
            CommonSlideData = CreateCommonSlideData()
        };
        slide.Save(slidePart);
        SlideIdList slideIdList = presentationPart.Presentation.SlideIdList!;
        SlideId slideId = new() { Id = (uint)(255 + slideIdList.Count() + 1), RelationshipId = presentationPart.GetIdOfPart(slidePart) };
        slideIdList.Append(slideId);
    }

    public void Save()
    {
        presentationDocument.Save();
        presentationDocument.Dispose();
    }

    public void SaveAs(string filePath)
    {
        presentationDocument.Clone(filePath).Dispose();
    }
}
