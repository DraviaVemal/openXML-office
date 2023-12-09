using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation;
public class PowerPoint : SlideLayout
{

    private readonly PresentationDocument presentationDocument;
    private ExtendedFilePropertiesPart? extendedFilePropertiesPart;
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
        presentationPart.Presentation.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        presentationPart.Presentation.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        if (presentationPart.Presentation.GetFirstChild<SlideMasterIdList>() == null)
        {
            presentationPart.Presentation.AppendChild(new SlideMasterIdList());
        }
        if (presentationPart.Presentation.SlideIdList == null)
        {
            presentationPart.Presentation.AppendChild(new SlideIdList());
        }
        if (presentationPart.Presentation.GetFirstChild<SlideSize>() == null)
        {
            presentationPart.Presentation.AppendChild(new SlideSize() { Cx = 12192000, Cy = 6858000 });
        }
        if (presentationPart.Presentation.GetFirstChild<NotesSize>() == null)
        {
            presentationPart.Presentation.AppendChild(new NotesSize() { Cx = 6858000, Cy = 9144000 });
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
        if (presentationPart.PresentationPropertiesPart == null)
        {
            PresentationPropertiesPart presentationPropertiesPart = presentationPart.AddNewPart<PresentationPropertiesPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            presentationPropertiesPart.PresentationProperties ??= new PresentationProperties();
            presentationPropertiesPart.PresentationProperties.Save();
        }
        if (presentationPart.ViewPropertiesPart == null)
        {
            ViewPropertiesPart viewPropertiesPart = presentationPart.AddNewPart<ViewPropertiesPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            viewPropertiesPart.ViewProperties ??= new ViewProperties();
            viewPropertiesPart.ViewProperties.Save();
        }
        if (presentationDocument.ExtendedFilePropertiesPart == null)
        {
            extendedFilePropertiesPart = presentationDocument.AddExtendedFilePropertiesPart();
            extendedFilePropertiesPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            extendedFilePropertiesPart.Properties.Save();
        }
        if (presentationPart.TableStylesPart == null)
        {
            TableStylesPart tableStylesPart = presentationPart.AddNewPart<TableStylesPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            tableStylesPart.TableStyleList ??= new A.TableStyleList()
            {
                Default = string.Format("{{{0}}}", Guid.NewGuid().ToString("D").ToUpper())
            };
            tableStylesPart.TableStyleList.Save();
        }
        if (!presentationPart.SlideMasterParts.Any())
        {
            slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            slideMasterPart.SlideMaster = CreateSlideMaster();
            SlideMasterIdList slideMasterIdList = presentationPart.Presentation.SlideMasterIdList!;
            SlideMasterId slideMasterId = new() { Id = (uint)(2147483647 + slideMasterIdList.Count() + 1), RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) };
            slideMasterIdList.Append(slideMasterId);
            slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
            AddSlideLayoutIdToList(slideMasterPart.GetIdOfPart(slideLayoutPart));
            slideLayoutPart.SlideLayout = CreateSlideLayout();
            slideLayoutPart.SlideLayout.Save();
            slideMasterPart.SlideMaster.Save();
        }
        presentationPart.Presentation.Save();
    }

    public void AddBlankSlide()
    {
        SlidePart slidePart = presentationPart!.AddNewPart<SlidePart>(string.Format("rId{0}", presentationPart.Parts.Count() + 1));
        Slide slide = new()
        {
            CommonSlideData = CreateCommonSlideData(),
            ColorMapOverride = new ColorMapOverride()
            {
                MasterColorMapping = new A.MasterColorMapping()
            }
        };
        slide.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        slide.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
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
