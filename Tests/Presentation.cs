/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using OpenXMLOffice.Excel;
using OpenXMLOffice.Global;
using OpenXMLOffice.Presentation;

namespace OpenXMLOffice.Tests
{
    [TestClass]
    public class Presentation
    {
        #region Private Fields

        private static PowerPoint powerPoint = new(new MemoryStream());

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
            //1
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting());
            //2
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting()
            {
                AreaChartTypes = AreaChartTypes.STACKED
            });
            //3
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting()
            {
                AreaChartTypes = AreaChartTypes.PERCENT_STACKED
            });
            //4
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                BarChartDataLabel = new BarChartDataLabel()
                {
                    DataLabelPosition = BarChartDataLabel.DataLabelPositionValues.INSIDE_END
                },
                BarChartSeriesSettings = new List<BarChartSeriesSetting>(){
                    new(),
                    new(){
                        BarChartDataLabel = new BarChartDataLabel(){
                            DataLabelPosition = BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END
                        }
                    }
                }
            });
            //5
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                Title = "Change Data Layout",
                BarChartTypes = BarChartTypes.STACKED
            });
            //6
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                BarChartTypes = BarChartTypes.PERCENT_STACKED
            });
            //7
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                Title = "Color Change Chart",
                ChartLegendOptions = new ChartLegendOptions()
                {
                    LegendPosition = ChartLegendOptions.LegendPositionValues.TOP
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
            //8
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                ColumnChartTypes = ColumnChartTypes.STACKED
            });
            //9
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                ColumnChartTypes = ColumnChartTypes.PERCENT_STACKED
            });
            //10
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting());
            //11
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.STACKED
            });
            //12
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.PERCENT_STACKED
            });
            //13
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.CLUSTERED_MARKER
            });
            //14
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.STACKED_MARKER
            });
            //15
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartTypes = LineChartTypes.PERCENT_STACKED_MARKER
            });
            //16
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new PieChartSetting());
            //17
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new PieChartSetting()
            {
                PieChartTypes = PieChartTypes.DOUGHNUT
            });
            //18
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting());
            //19
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH
            });
            //20
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH_MARKER
            });
            //21
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT
            });
            //22
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT_MARKER
            });
            //23
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(3, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.BUBBLE
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
        public void AddDevChart()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                LineChartDataLabel = new LineChartDataLabel()
                {
                    DataLabelPosition = LineChartDataLabel.DataLabelPositionValues.LEFT,
                    ShowCategoryName = true,
                    ShowValue = true,
                    Separator = ". "
                },
                ChartDataSetting = new ChartDataSetting()
                {
                    ChartDataColumnEnd = 2,
                    ValueFromColumn = new Dictionary<uint, uint>(){
                        {2,4}
                    }
                },
                LineChartSeriesSettings = new List<LineChartSeriesSetting>(){
                    new(),
                    new(){
                        LineChartDataLabel = new LineChartDataLabel(){
                            DataLabelPosition = LineChartDataLabel.DataLabelPositionValues.RIGHT
                        }
                    }
                }
            });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddPicture()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new PictureSetting());
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new PictureSetting());
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddScatterPlot()
        {
            //1
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                Title = "Default"
            });
            //2
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH,
                Title = "Scatter Smooth"
            });
            //3
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH_MARKER,
                Title = "Scatter Smooth Market"
            });
            //4
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT,
                Title = "Scatter Stright"
            });
            //5
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT_MARKER,
                Title = "Scatter Straight Marker"
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(3, true), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.BUBBLE,
                Title = "Scatter  Bubble"
            });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddTable()
        {
            Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            slide.AddTable(CreateTableRowPayload(5), new TableSetting()
            {
                Name = "New Table",
                WidthType = TableSetting.WidthOptionValues.AUTO,
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
            shapes1[0].ReplaceTextBox(slide.AddTextBox(new TextBoxSetting()
            {
                Text = "Dravia Vemal",
                FontFamily = "Bernard MT Condensed"
            }));
            shapes2[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
            {
                Text = "Vemal Dravia",
                TextBackground = "777777"
            }));
            shapes3[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
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
            shape1[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(),
            new ColumnChartSetting()
            {
                Title = "Title",
                ChartLegendOptions = new ChartLegendOptions()
                {
                    IsEnableLegend = false
                },
            }));
            shape2[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(),
            new BarChartSetting()
            {
                ChartLegendOptions = new ChartLegendOptions()
                {
                    LegendPosition = ChartLegendOptions.LegendPositionValues.RIGHT
                }
            }));
            shape3[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new LineChartSetting()
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
            }));
            shape4[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new LineChartSetting()));
            shape5[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new AreaChartSetting()));
            shape6[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
            {
                Text = "Test"
            }));
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

        private DataCell[][] CreateDataCellPayload(int payloadSize = 5, bool IsValueAxis = false)
        {
            Random random = new();
            DataCell[][] data = new DataCell[payloadSize][];
            data[0] = new DataCell[payloadSize];
            for (int col = 1; col < payloadSize; col++)
            {
                data[0][col] = new DataCell
                {
                    CellValue = $"Series {col}",
                    DataType = CellDataType.STRING
                };
            }
            for (int row = 1; row < payloadSize; row++)
            {
                data[row] = new DataCell[payloadSize];
                data[row][0] = new DataCell
                {
                    CellValue = $"Category {row}",
                    DataType = CellDataType.STRING,
                    StyleId = 1
                };
                for (int col = IsValueAxis ? 0 : 1; col < payloadSize; col++)
                {
                    data[row][col] = new DataCell
                    {
                        CellValue = random.Next(1, 100).ToString(),
                        DataType = CellDataType.NUMBER,
                        NumberFormat = "General",
                        StyleId = 1
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