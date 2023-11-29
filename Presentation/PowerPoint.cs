using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class PowerPoint
{

    private readonly PresentationDocument presentationDocument;
    private PresentationPart? presentationPart;
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
        if (!presentationPart.SlideMasterParts.Any())
        {
            SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new();
            SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new();
            slideLayoutPart.SlideLayout.Save();
            slideMasterPart.SlideMaster.Save();
        }
        presentationPart.Presentation ??= new DocumentFormat.OpenXml.Presentation.Presentation();
        presentationPart.Presentation.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        presentationPart.Presentation.Save();
    }

    public void AddBlankSlide()
    {
        SlidePart slidePart = presentationPart!.AddNewPart<SlidePart>();
        Slide slide = new();
        slidePart.Slide = slide;
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
