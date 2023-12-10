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
        powerPoint.Save();
        Assert.IsTrue(true);
    }

    [TestMethod]
    public void OpenExistingPresentationNonEdit()
    {
        PowerPoint powerPoint1 = new("./TestFiles/basic_test.pptx", false);
        powerPoint1.Save();
        Assert.IsTrue(true);
    }

    [TestMethod]
    public void OpenExistingPresentationEdit()
    {
        PowerPoint powerPoint1 = new("./TestFiles/basic_test.pptx", true);
        powerPoint1.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
        powerPoint1.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
        Slide slide = powerPoint1.GetSlideByIndex(0);
        slide.FindShapeByText("");
        powerPoint1.MoveSlideByIndex(4, 0);
        powerPoint1.SaveAs(string.Format("../../edit-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
        Assert.IsTrue(true);
    }
}
