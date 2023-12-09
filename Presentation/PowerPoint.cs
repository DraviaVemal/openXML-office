using DocumentFormat.OpenXml;
using OpenXMLOffice.Global;

namespace OpenXMLOffice.Presentation;
public class PowerPoint
{
    private readonly Presentation presentation;

    public PowerPoint(string filePath, PowerPointProperties? powerPointProperties = null)
    {
        presentation = new(filePath, powerPointProperties);
    }

    public PowerPoint(string filePath, PowerPointProperties? powerPointProperties = null, PresentationDocumentType presentationDocumentType = PresentationDocumentType.Presentation)
    {
        presentation = new(filePath, powerPointProperties, presentationDocumentType);
    }

    public PowerPoint(Stream stream, PresentationDocumentType presentationDocumentType, PowerPointProperties? powerPointProperties = null)
    {
        presentation = new(stream, powerPointProperties, presentationDocumentType);
    }

    public void AddSlide(Constants.SlideLayoutType slideLayoutType)
    {
        switch (slideLayoutType)
        {
            default: //Blank
                presentation.AddBlankSlide();
                break;
        }
    }

    public void Save()
    {
        presentation.Save();
    }
    public void SaveAs(string filePath)
    {
        presentation.SaveAs(filePath);
    }
}
