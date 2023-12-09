using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class PowerPoint
{

    private readonly PresentationDocument presentationDocument;
    private PresentationPart? presentationPart;
    private SlideMasterPart? slideMasterPart;
    private SlideLayoutPart? slideLayoutPart;
    public PowerPoint(string filePath, bool isEditable = true, bool autosave = true)
    {
        presentationDocument = PresentationDocument.Open(filePath, isEditable, new OpenSettings()
        {
            AutoSave = autosave
        });
    }
    public PowerPoint(string filePath, PresentationDocumentType presentationDocumentType, bool autosave = true)
    {
        presentationDocument = PresentationDocument.Create(filePath, presentationDocumentType, autosave);
        PreparePresentation();
    }
    public PowerPoint(Stream stream, PresentationDocumentType presentationDocumentType, bool autosave = true)
    {
        presentationDocument = PresentationDocument.Create(stream, presentationDocumentType, autosave);
        PreparePresentation();
    }
    /// <summary>
    /// Initializes a new instance of the PowerPoint class and creates a new PowerPoint presentation using a template file specified by the filePath parameter.
    /// </summary>
    /// <param name="filePath">The file path to the template file for creating the PowerPoint presentation.</param>
    public PowerPoint(string filePath)
    {
        presentationDocument = PresentationDocument.CreateFromTemplate(filePath);
        PreparePresentation();
    }

    private void PreparePresentation()
    {
        presentationPart = presentationDocument.PresentationPart ?? presentationDocument.AddPresentationPart();
        presentationPart.Presentation ??= new P.Presentation();
        presentationPart.Presentation.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        if (presentationPart.Presentation.GetFirstChild<SlideMasterIdList>() == null)
        {
            presentationPart.Presentation.AddChild(new SlideMasterIdList());
        }
        if (presentationPart.Presentation.GetFirstChild<SlideIdList>() == null)
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
        if (!presentationPart.SlideMasterParts.Any())
        {
            slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new P.SlideMaster(new P.CommonSlideData(new P.ShapeTree()
            {
                GroupShapeProperties = new P.GroupShapeProperties(),
                NonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()
                )
            }));
            slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new P.SlideLayout(new P.CommonSlideData(new P.ShapeTree()
            {
                GroupShapeProperties = new P.GroupShapeProperties(),
                NonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()
                )
            }));
            slideMasterPart.SlideMaster.AddChild(new P.ColorMap()
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

    public void AddBlankSlide()
    {
        SlidePart slidePart = presentationPart!.AddNewPart<SlidePart>();
        P.Slide slide = new();
        slidePart.Slide = slide;
        slideMasterPart!.SlideMaster.AppendChild(slide);
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
