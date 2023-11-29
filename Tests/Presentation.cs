using OpenXMLOffice.Presentation;

namespace OpenXMLOffice.Tests;

[TestClass]
public class Presentation
{
    private static PowerPoint powerPoint = new(new MemoryStream(), DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

    [ClassInitialize]
    public static void ClassInitialize(TestContext context)
    {
        powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
    }

    [ClassCleanup]
    public static void ClassCleanup()
    {
        powerPoint.Save();
    }

    [TestMethod]
    public void SheetConstructorFile()
    {
        PowerPoint powerPoint1 = new("../try.pptx", DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        Assert.IsNotNull(powerPoint1);
        File.Delete("../try.pptx");
    }
}
