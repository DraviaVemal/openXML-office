using OpenXMLOffice.Presentation;

namespace OpenXMLOffice.Tests;

[TestClass]
public class Presentation
{
    private static PowerPoint powerPoint = new(new MemoryStream(), DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

    [ClassInitialize]
    public static void ClassInitialize(TestContext context)
    {
        powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), null);
    }

    [ClassCleanup]
    public static void ClassCleanup()
    {
        powerPoint.Save();
    }

    [TestMethod]
    public void SheetConstructorFile()
    {
        PowerPoint powerPoint1 = new("../try.pptx", null);
        Assert.IsNotNull(powerPoint1);
        powerPoint1.Save();
        File.Delete("../try.pptx");
    }

    [TestMethod]
    public void AddBlankSlide()
    {
        powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
        powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
        powerPoint.Save();
        Assert.IsTrue(true);
    }

    [TestMethod]
    public void OpenExistingPresentationNonEdit()
    {
        PowerPoint powerPoint1 = new("C:\\Users\\draviavemal\\Projects\\OpenXMLOffice\\1.pptx", false);
        powerPoint1.Save();
        Assert.IsTrue(true);
    }

    [TestMethod]
    public void OpenExistingPresentationEdit()
    {
        PowerPoint powerPoint1 = new("C:\\Users\\draviavemal\\Projects\\OpenXMLOffice\\1.pptx", true);
        powerPoint1.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
        powerPoint1.Save();
        Assert.IsTrue(true);
    }
}
