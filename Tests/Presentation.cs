// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Excel;
using OpenXMLOffice.Global;
using OpenXMLOffice.Presentation;

namespace OpenXMLOffice.Tests
{
    /// <summary>
    /// Presentation Test Class
    /// </summary>
    [TestClass]
    public class Presentation
    {
        #region Private Fields

        private static PowerPoint powerPoint = new(new MemoryStream());

        #endregion Private Fields

        #region Public Methods

        /// <summary>
        /// Save Presenation on text completion cleanup
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            powerPoint.Save();
        }

        /// <summary>
        /// Initialize Presentation For Test
        /// </summary>
        /// <param name="context">
        /// </param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            powerPoint = new(string.Format("../../test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), null);
        }

        /// <summary>
        /// Add All Chart Types to Slide
        /// </summary>
        [TestMethod]
        public void AddAllChartTypesToSlide()
        {
            //1
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting());
            //2
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting()
            {
                areaChartTypes = AreaChartTypes.STACKED,
                chartAxesOptions = new()
                {
                    horizontalFontSize = 20,
                    verticalFontSize = 25
                }
            });
            //3
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new AreaChartSetting()
            {
                title = "",
                areaChartTypes = AreaChartTypes.PERCENT_STACKED,
                chartDataSetting = new()
                {
                    chartDataColumnEnd = 1
                }
            });
            //4
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                barChartDataLabel = new BarChartDataLabel()
                {
                    dataLabelPosition = BarChartDataLabel.DataLabelPositionValues.INSIDE_END,
                    showValue = true,
                },
                barChartSeriesSettings = new(){
                    new(),
                    new(){
                        barChartDataLabel = new BarChartDataLabel(){
                            dataLabelPosition = BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END,
                            showCategoryName= true
                        }
                    }
                }
            });
            //5
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                title = "Change Data Layout",
                barChartTypes = BarChartTypes.STACKED
            });
            //6
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new BarChartSetting()
            {
                barChartTypes = BarChartTypes.PERCENT_STACKED
            });
            //7
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                title = "Color Change Chart",
                chartLegendOptions = new ChartLegendOptions()
                {
                    legendPosition = ChartLegendOptions.LegendPositionValues.TOP,
                    fontSize = 5
                },
                columnChartSeriesSettings = new(){
                    new(){
                        fillColor= "AABBCC"
                    },
                    new(){
                        fillColor= "CCBBAA"
                    }
                }
            });
            //8
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                columnChartTypes = ColumnChartTypes.STACKED
            });
            //9
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ColumnChartSetting()
            {
                columnChartTypes = ColumnChartTypes.PERCENT_STACKED
            });
            //10
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting());
            //11
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                lineChartTypes = LineChartTypes.STACKED
            });
            //12
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                lineChartTypes = LineChartTypes.PERCENT_STACKED
            });
            //13
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                lineChartTypes = LineChartTypes.CLUSTERED_MARKER
            });
            //14
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                lineChartTypes = LineChartTypes.STACKED_MARKER
            });
            //15
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                lineChartTypes = LineChartTypes.PERCENT_STACKED_MARKER
            });
            //16
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new PieChartSetting());
            //17
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new PieChartSetting()
            {
                pieChartTypes = PieChartTypes.DOUGHNUT,
                pieChartDataLabel = new()
                {
                    dataLabelPosition = PieChartDataLabel.DataLabelPositionValues.SHOW,
                    showCategoryName = true,
                    showValue = true,
                    separator = ". "
                }
            });
            //18
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting());
            //19
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH
            });
            //20
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH_MARKER
            });
            //21
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT
            });
            //22
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT_MARKER
            });
            //23
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(3, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.BUBBLE
            });
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Add Blank Slide to the PPT
        /// </summary>
        [TestMethod]
        public void AddBlankSlide()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Add Single Chart to the Slide
        /// </summary>
        [TestMethod]
        public void AddDevChart()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new LineChartSetting()
            {
                lineChartDataLabel = new LineChartDataLabel()
                {
                    dataLabelPosition = LineChartDataLabel.DataLabelPositionValues.LEFT,
                    showCategoryName = true,
                    showValue = true,
                    separator = ". "
                },
                chartDataSetting = new ChartDataSetting()
                {
                    chartDataColumnEnd = 2,
                    valueFromColumn = new Dictionary<uint, uint>(){
                        {2,4}
                    }
                },
                lineChartSeriesSettings = new(){
                    null,
                    new(){
                        lineChartDataLabel = new LineChartDataLabel(){
                            dataLabelPosition = LineChartDataLabel.DataLabelPositionValues.RIGHT
                        }
                    }
                }
            });
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Add Picture to the slide
        /// </summary>
        [TestMethod]
        public void AddPicture()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new PictureSetting());
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new PictureSetting());
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Add All type of sctter charts
        /// </summary>
        [TestMethod]
        public void AddScatterPlot()
        {
            //1
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                title = "Default"
            });
            //2
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH,
                title = "Scatter Smooth"
            });
            //3
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH_MARKER,
                title = "Scatter Smooth Market"
            });
            //4
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT,
                title = "Scatter Stright"
            });
            //5
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT_MARKER,
                title = "Scatter Straight Marker"
            });
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(3, true), new ScatterChartSetting()
            {
                scatterChartTypes = ScatterChartTypes.BUBBLE,
                title = "Scatter  Bubble"
            });
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Add Table To the Slide
        /// </summary>
        [TestMethod]
        public void AddTable()
        {
            Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
            slide.AddTable(CreateTableRowPayload(5), new TableSetting()
            {
                name = "New Table",
                widthType = TableSetting.WidthOptionValues.PERCENTAGE,
                tableColumnWidth = new() { 80, 20 }
            });
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Open and Edit Existing Presentation
        /// </summary>
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
                text = "Dravia Vemal",
                fontFamily = "Bernard MT Condensed"
            }));
            shapes2[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
            {
                text = "Vemal Dravia",
                textBackground = "777777"
            }));
            shapes3[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
            {
                text = "This is text box",
                fontSize = 22,
                isBold = true,
                textColor = "AAAAAA"
            }));
            powerPoint1.MoveSlideByIndex(4, 0);
            powerPoint1.SaveAs(string.Format("../../edit-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Open and Edit Existing Presentation with Chart
        /// </summary>
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
                title = "Title",
                chartLegendOptions = new ChartLegendOptions()
                {
                    isEnableLegend = false
                },
            }));
            shape2[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(),
            new BarChartSetting()
            {
                chartLegendOptions = new ChartLegendOptions()
                {
                    legendPosition = ChartLegendOptions.LegendPositionValues.RIGHT
                }
            }));
            shape3[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new LineChartSetting()
            {
                chartAxesOptions = new ChartAxesOptions()
                {
                    isHorizontalAxesEnabled = false
                },
                chartGridLinesOptions = new ChartGridLinesOptions()
                {
                    isMajorCategoryLinesEnabled = true,
                    isMajorValueLinesEnabled = true,
                    isMinorCategoryLinesEnabled = true,
                    isMinorValueLinesEnabled = true,
                }
            }));
            shape4[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new LineChartSetting()));
            shape5[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new AreaChartSetting()));
            shape6[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
            {
                text = "Test"
            }));
            powerPoint1.SaveAs(string.Format("../../chart-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Open and close Presentation without editing
        /// </summary>
        [TestMethod]
        public void OpenExistingPresentationNonEdit()
        {
            PowerPoint powerPoint1 = new("./TestFiles/basic_test.pptx", false);
            powerPoint1.Save();
            Assert.IsTrue(true);
        }

        /// <summary>
        /// Check PPT File creation
        /// </summary>
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

        private static DataCell[][] CreateDataCellPayload(int payloadSize = 5, bool IsValueAxis = false)
        {
            Random random = new();
            DataCell[][] data = new DataCell[payloadSize][];
            data[0] = new DataCell[payloadSize];
            for (int col = 1; col < payloadSize; col++)
            {
                data[0][col] = new DataCell
                {
                    cellValue = $"Series {col}",
                    dataType = CellDataType.STRING
                };
            }
            for (int row = 1; row < payloadSize; row++)
            {
                data[row] = new DataCell[payloadSize];
                data[row][0] = new DataCell
                {
                    cellValue = $"Category {row}",
                    dataType = CellDataType.STRING
                };
                for (int col = IsValueAxis ? 0 : 1; col < payloadSize; col++)
                {
                    data[row][col] = new DataCell
                    {
                        cellValue = random.Next(1, 100).ToString(),
                        dataType = CellDataType.NUMBER,
                        styleSetting = new()
                        {
                            numberFormat = "0.00",
                            fontSize = 20
                        }
                    };
                }
            }
            return data;
        }

        private static TableRow[] CreateTableRowPayload(int rowCount)
        {
            TableRow[] data = new TableRow[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                TableRow row = new()
                {
                    height = 370840,
                    tableCells = new List<TableCell>
                {
                    new() {
                        value = $"Row {i + 1}, Cell 1",
                        textColor = "FF0000",
                        bottomBorder = true,
                        fontSize=22,
                        leftBorder = true,
                        topBorder = true,
                        rightBorder =true,
                    },
                    new() {
                        value = $"Row {i + 1}, Cell 2",
                        textColor = "00FF00",
                        isBold = true,
                        bottomBorder = true,
                        leftBorder = true,
                        topBorder = true,
                        rightBorder =true,
                        isItalic = true
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