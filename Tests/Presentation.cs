using OpenXMLOffice.Excel;
using OpenXMLOffice.Global;
using OpenXMLOffice.Presentation;

namespace OpenXMLOffice.Tests
{
    [TestClass]
    public class Presentation
    {
        #region Private Fields

        private static PowerPoint powerPoint = new(new MemoryStream(), DocumentFormat.OpenXml.PresentationDocumentType.Presentation);

        #endregion Private Fields

        #region Public Methods

        [ClassCleanup]
        public static void ClassCleanup()
        {
            powerPoint.Save();
        }

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), null);
        }

        [TestMethod]
        public void AddAllChartTypesToSlide()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting());
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting()
            {
                AreaChartTypes = AreaChartTypes.STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting()
            {
                AreaChartTypes = AreaChartTypes.PERCENT_STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                BarChartSeriesSettings = new List<BarChartSeriesSetting>(){
                    new(),
                    new(){
                        BarChartDataLabel = new BarChartDataLabel(){
                            DataLabelPosition = BarChartDataLabel.eDataLabelPosition.OUTSIDE_END
                        }
                    }
                }
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                Title = "Change Data Layout",
                BarChartTypes = BarChartTypes.STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                BarChartTypes = BarChartTypes.PERCENT_STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                Title = "Color Change Chart",
                ChartLegendOptions = new ChartLegendOptions()
                {
                    legendPosition = ChartLegendOptions.eLegendPosition.TOP
                },
                ColumnChartSeriesSettings = new List<ColumnChartSeriesSetting>(){
                    new(){
                        FillColor= "AABBCC"
                    },
                    new(){
                        FillColor= "CCBBAA"
                    }
                }
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                ColumnChartTypes = ColumnChartTypes.STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                ColumnChartTypes = ColumnChartTypes.PERCENT_STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting());
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.PERCENT_STACKED
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.CLUSTERED_MARKER
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.STACKED_MARKER
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.PERCENT_STACKED_MARKER
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new PieChartSetting());
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new PieChartSetting()
            {
                PieChartTypes = PieChartTypes.DOUGHNUT
            });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddBlankSlide()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddTable()
        {
            Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            slide.AddTable(CreateTableRowPayload(5), new TableSetting()
            {
                Name = "New Table",
                WidthType = TableSetting.eWidthType.AUTO,
                TableColumnwidth = new() { 100, 100 }
            });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void OpenExistingPresentationEdit()
        {
            PowerPoint powerPoint1 = new("./TestFiles/basic_test.pptx", true);
            powerPoint1.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            powerPoint1.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            Slide slide = powerPoint1.GetSlideByIndex(0);
            List<Shape> shapes1 = slide.FindShapeByText("Slide_1_Shape_1").ToList();
            List<Shape> shapes2 = slide.FindShapeByText("Slide_1_Shape_2").ToList();
            List<Shape> shapes3 = slide.FindShapeByText("Test Update").ToList();
            shapes1[0].ReplaceShape(new TextBox().CreateTextBox(1, new TextBoxSetting()
            {
                Text = "Dravia Vemal",
                FontFamily = "Bernard MT Condensed"
            }));
            shapes2[0].ReplaceShape(new TextBox().CreateTextBox(2, new TextBoxSetting()
            {
                Text = "Vemal Dravia",
                TextBackground = "777777"
            }));
            shapes3[0].ReplaceShape(new TextBox().CreateTextBox(30, new TextBoxSetting()
            {
                Text = "This is text box",
                FontSize = 22,
                IsBold = true,
                TextColor = "AAAAAA"
            }));
            powerPoint1.MoveSlideByIndex(4, 0);
            powerPoint1.SaveAs(string.Format("../../edit-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void OpenExistingPresentationEditBarChart()
        {
            PowerPoint powerPoint1 = new("./TestFiles/basic_test.pptx", true);
            Slide Slide = powerPoint1.GetSlideByIndex(0);
            List<Shape> shape1 = Slide.FindShapeByText("Slide_1_Shape_1").ToList();
            List<Shape> shape2 = Slide.FindShapeByText("Slide_1_Shape_2").ToList();
            List<Shape> shape3 = Slide.FindShapeByText("Slide_1_Shape_3").ToList();
            List<Shape> shape4 = Slide.FindShapeByText("Slide_1_Shape_4").ToList();
            List<Shape> shape5 = Slide.FindShapeByText("Slide_1_Shape_5").ToList();
            List<Shape> shape6 = Slide.FindShapeByText("Slide_1_Shape_6").ToList();
            shape1[0].ReplaceShape(new Chart(Slide, CreateDataCellPayload(),
            new ColumnChartSetting()
            {
                Title = "Title",
                ChartLegendOptions = new ChartLegendOptions()
                {
                    IsEnableLegend = false
                },
            }).GetChartGraphicFrame());
            shape2[0].ReplaceShape(new Chart(Slide, CreateDataCellPayload(),
            new BarChartSetting()
            {
                ChartLegendOptions = new ChartLegendOptions()
                {
                    legendPosition = ChartLegendOptions.eLegendPosition.RIGHT
                }
            }).GetChartGraphicFrame());
            shape3[0].ReplaceShape(new Chart(Slide, CreateDataCellPayload(), new LineChartSetting()
            {
                ChartAxesOptions = new ChartAxesOptions()
                {
                    IsHorizontalAxesEnabled = false
                },
                ChartGridLinesOptions = new ChartGridLinesOptions()
                {
                    IsMajorCategoryLinesEnabled = true,
                    IsMajorValueLinesEnabled = true,
                    IsMinorCategoryLinesEnabled = true,
                    IsMinorValueLinesEnabled = true,
                }
            }).GetChartGraphicFrame());
            shape4[0].ReplaceShape(new Chart(Slide, CreateDataCellPayload(), new LineChartSetting()).GetChartGraphicFrame());
            shape5[0].ReplaceShape(new Chart(Slide, CreateDataCellPayload(), new AreaChartSetting()).GetChartGraphicFrame());
            shape6[0].ReplaceShape(new Chart(Slide, CreateDataCellPayload(), new PieChartSetting()).GetChartGraphicFrame());
            powerPoint1.SaveAs(string.Format("../../chart-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
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
        public void SheetConstructorFile()
        {
            PowerPoint powerPoint1 = new("../try.pptx", null);
            Assert.IsNotNull(powerPoint1);
            powerPoint1.Save();
            File.Delete("../try.pptx");
        }

        #endregion Public Methods

        #region Private Methods

        private DataCell[][] CreateDataCellPayload()
        {
            Random random = new();
            DataCell[][] data = new DataCell[5][];
            data[0] = new DataCell[5];
            for (int col = 1; col < 5; col++)
            {
                data[0][col] = new DataCell
                {
                    CellValue = $"Series {col}",
                    DataType = CellDataType.STRING
                };
            }
            for (int row = 1; row < 5; row++)
            {
                data[row] = new DataCell[5];
                data[row][0] = new DataCell
                {
                    CellValue = $"Category {row}",
                    DataType = CellDataType.STRING
                };
                for (int col = 1; col < 5; col++)
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

        private TableRow[] CreateTableRowPayload(int rowCount)
        {
            TableRow[] data = new TableRow[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                TableRow row = new()
                {
                    Height = 370840,
                    TableCells = new List<TableCell>
                {
                    new() {
                        Value = $"Row {i + 1}, Cell 1",
                        TextColor = "FF0000"
                    },
                    new() {
                        Value = $"Row {i + 1}, Cell 2",
                        TextColor = "00FF00"
                    },
                }
                };
                data[i] = row;
            }
            return data;
        }

        #endregion Private Methods
    }
}