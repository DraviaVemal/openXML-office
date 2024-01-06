// Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License. See License in
// the project root for license information.
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
                BarChartSeriesSettings = new List<BarChartSeriesSetting>(){
                    new(),
                    new(){
                        BarChartDataLabel = new BarChartDataLabel(){
                            DataLabelPosition = BarChartDataLabel.eDataLabelPosition.OUTSIDE_END
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
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting());
            //19
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH
            });
            //20
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH_MARKER
            });
            //21
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT
            });
            //22
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT_MARKER
            });
            //23
            // powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting()
            // {
            //     ScatterChartTypes = ScatterChartTypes.BUBBLE
            // });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddScatterPlot()
        {
            DataCell[][] data = new DataCell[4][];
            data[0] = new DataCell[4];
            data[1] = new DataCell[4];
            data[2] = new DataCell[4];
            data[3] = new DataCell[4];
            data[0][1] = new DataCell
            {
                CellValue = "Series 1",
                DataType = CellDataType.STRING
            };
            data[0][2] = new DataCell
            {
                CellValue = "Series 2",
                DataType = CellDataType.STRING
            };
            data[0][3] = new DataCell
            {
                CellValue = "Series 3",
                DataType = CellDataType.STRING
            };

            data[1][0] = new DataCell
            {
                CellValue = "1",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[1][1] = new DataCell
            {
                CellValue = "10",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[1][2] = new DataCell
            {
                CellValue = "12",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[1][3] = new DataCell
            {
                CellValue = "12.5",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };

            data[2][0] = new DataCell
            {
                CellValue = "2",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[2][1] = new DataCell
            {
                CellValue = "20",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[2][2] = new DataCell
            {
                CellValue = "22",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[2][3] = new DataCell
            {
                CellValue = "13",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };

            data[3][0] = new DataCell
            {
                CellValue = "3",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[3][1] = new DataCell
            {
                CellValue = "3",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[3][2] = new DataCell
            {
                CellValue = "7",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };
            data[3][3] = new DataCell
            {
                CellValue = "4.5",
                DataType = CellDataType.NUMBER,
                numberFormatting = "General",
                styleId = 1
            };

            //1
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new ScatterChartSetting()
            {
                Title = "Default"
            });
            //2
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH,
                Title = "Scatter Smooth"
            });
            //3
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH_MARKER,
                Title = "Scatter Smooth Market"
            });
            //4
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT,
                Title = "Scatter Stright"
            });
            //5
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new ScatterChartSetting()
            {
                ScatterChartTypes = ScatterChartTypes.SCATTER_STRIGHT_MARKER,
                Title = "Scatter Straight Marker"
            });
            // powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new
            // ScatterChartSetting() { ScatterChartTypes = ScatterChartTypes.BUBBLE, Title =
            // "Scatter Straight Bubble" });
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void AddDevChart()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new ScatterChartSetting()
            {
                Title = "Color Change Chart",
                ScatterChartTypes = ScatterChartTypes.SCATTER_SMOOTH,
                ChartDataSetting = new ChartDataSetting()
                {
                    ChartDataRowStart = 1,
                    ChartDataColumnStart = 2
                },
                ChartLegendOptions = new ChartLegendOptions()
                {
                    legendPosition = ChartLegendOptions.eLegendPosition.TOP
                }
            }
            );
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
        public void AddPicture()
        {
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new PictureSetting());
            powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new PictureSetting());
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
            shapes1[0].ReplaceTextBox(new TextBox(new TextBoxSetting()
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
                    legendPosition = ChartLegendOptions.eLegendPosition.RIGHT
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
            shape6[0].ReplaceChart(new Chart(Slide, CreateDataCellPayload(), new PieChartSetting()));
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