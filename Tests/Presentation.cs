using OpenXMLOffice.Excel;
using OpenXMLOffice.Presentation;

namespace OpenXMLOffice.Tests
{
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
            List<Shape> shapes = slide.FindShapeByText("Slide_1_Shape_1").ToList();
            List<Shape> shapes1 = slide.FindShapeByText("Slide_1_Shape_2").ToList();
            List<Shape> shapes2 = slide.FindShapeByText("Test Update").ToList();
            shapes[0].ReplaceShape(new TextBox()
            {
                Text = "Dravia Vemal",
                FontFamily = "Bernard MT Condensed"
            }.CreateTextBox());
            shapes1[0].ReplaceShape(new TextBox()
            {
                Text = "Vemal Dravia",
                TextBackground = "777777"
            }.CreateTextBox());
            shapes2[0].ReplaceShape(new TextBox()
            {
                Text = "This is text box",
                FontSize = 22,
                IsBold = true,
                TextColor = "AAAAAA"
            }.CreateTextBox());
            powerPoint1.MoveSlideByIndex(4, 0);
            powerPoint1.SaveAs(string.Format("../../edit-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void OpenExistingPresentationEditBarChart()
        {
            PowerPoint powerPoint1 = new("./TestFiles/basic_test.pptx", true);
            Slide Slide = powerPoint1.GetSlideByIndex(0);
            List<Shape> shapes = Slide.FindShapeByText("Slide_1_Shape_1").ToList();
            shapes[0].ReplaceShape(new Chart(Slide).CreateChart(Global.GlobalConstants.ChartTypes.BAR, CreateDataPayload()));
            powerPoint1.SaveAs(string.Format("../../chart-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
            Assert.IsTrue(true);
        }

        private DataCell[][] CreateDataPayload()
        {
            Random random = new();
            DataCell[][] data = new DataCell[5][];
            data[0] = new DataCell[5];
            for (int col = 0; col < 5; col++)
            {
                data[0][col] = new DataCell
                {
                    CellValue = $"Heading {col + 1}",
                    DataType = CellDataType.STRING
                };
            }
            for (int row = 1; row < 5; row++)
            {
                data[row] = new DataCell[5];
                for (int col = 0; col < 5; col++)
                {
                    data[row][col] = new DataCell
                    {
                        CellValue = random.Next(1, 100).ToString(),
                        DataType = CellDataType.NUMBER,
                        numberFormatting = "General",
                        styleId = 1
                    };
                }
            }
            return data;
        }
    }
}
