// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using X = OpenXMLOffice.Spreadsheet_2007;
using G = OpenXMLOffice.Global_2007;
using OpenXMLOffice.Presentation_2007;
using OpenXMLOffice.Global_2016;

namespace OpenXMLOffice.Tests
{
	/// <summary>
	/// Presentation Test Class
	/// </summary>
	[TestClass]
	public class Presentation
	{
		private static PowerPoint powerPoint = new(new MemoryStream());

		private static readonly string resultPath = "../../testResult";
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
			if (!Directory.Exists(resultPath))
			{
				Directory.CreateDirectory(resultPath);
			}
			powerPoint = new(string.Format("{1}/test-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath), null);
		}

		/// <summary>
		/// Add All Chart Types to Slide
		/// </summary>
		[TestMethod]
		[TestCategory("Chart")]
		public void Add2007Charts()
		{
			//1
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.AreaChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new()
			});
			//2
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.AreaChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				areaChartType = G.AreaChartTypes.STACKED,
				chartAxesOptions = new()
				{
					horizontalFontSize = 20,
					verticalFontSize = 25
				}
			});
			//3
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.AreaChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				titleOptions = new()
				{
					title = "test"
				},
				areaChartType = G.AreaChartTypes.PERCENT_STACKED,
				chartDataSetting = new()
				{
					chartDataColumnEnd = 2
				}
			});
			//4
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.BarChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				chartAxesOptions = new()
				{
					isHorizontalAxesEnabled = false,
				},
				barChartDataLabel = new G.BarChartDataLabel()
				{
					dataLabelPosition = G.BarChartDataLabel.DataLabelPositionValues.INSIDE_END,
					showValue = true,
				},
				barChartSeriesSettings = new(){
					new(),
					new(){
						barChartDataLabel = new G.BarChartDataLabel(){
							dataLabelPosition = G.BarChartDataLabel.DataLabelPositionValues.OUTSIDE_END,
							showCategoryName= true
						}
					}
				}
			});
			//5
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.BarChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				titleOptions = new()
				{
					isItalic = true,
					title = "Bar Chart"
				},
				barChartType = G.BarChartTypes.STACKED
			});
			//6
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.BarChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				barChartType = G.BarChartTypes.PERCENT_STACKED
			});
			//7
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.ColumnChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				titleOptions = new()
				{
					title = "Column Chart"
				},
				chartLegendOptions = new G.ChartLegendOptions()
				{
					legendPosition = G.ChartLegendOptions.LegendPositionValues.TOP,
					fontSize = 5
				},
				columnChartSeriesSettings = new(){
					null,
					new(){
						columnChartDataPointSettings = new(){
						null,
						new(){
							fillColor = "FF0000"
						},
						new(){
							fillColor = "00FF00"
						},
					},
						fillColor= "AABBCC"
					},
					new(){
						fillColor= "CCBBAA"
					}
				}
			});
			//8
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.ColumnChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				columnChartType = G.ColumnChartTypes.STACKED
			});
			//9
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.ColumnChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				columnChartType = G.ColumnChartTypes.PERCENT_STACKED
			});
			//10
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				lineChartSeriesSettings = new(){
					new(){
						lineChartLineFormat = new(){
							dashType = G.DrawingPresetLineDashValues.DASH_DOT,
							lineColor = "FF0000",
							beginArrowValues= G.DrawingBeginArrowValues.ARROW,
							endArrowValues= G.DrawingEndArrowValues.TRIANGLE,
							lineStartWidth = G.LineWidthValues.MEDIUM,
							lineEndWidth = G.LineWidthValues.LARGE,
							outlineCapTypeValues = G.OutlineCapTypeValues.ROUND,
							outlineLineTypeValues = G.OutlineLineTypeValues.DOUBLE,
							width = 5
						}
					}
				}
			});
			//11
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				plotAreaOptions = new()
				{
					manualLayout = new()
					{
						x = 0.2F,
						y = 0.1F,
						width = 0.5F,
						height = 0.5F
					}
				},
				lineChartType = G.LineChartTypes.STACKED
			});
			//12
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				lineChartType = G.LineChartTypes.PERCENT_STACKED,
				chartLegendOptions = new()
				{
					manualLayout = new()
					{
						x = 0.5F,
						y = 0.9F,
						width = 0.5F,
						height = 0.1F
					}
				}
			});
			//13
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				lineChartType = G.LineChartTypes.CLUSTERED_MARKER
			});
			//14
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				lineChartType = G.LineChartTypes.STACKED_MARKER
			});
			//15
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				lineChartType = G.LineChartTypes.PERCENT_STACKED_MARKER
			});
			//16
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.PieChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
			//17
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.PieChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				pieChartType = G.PieChartTypes.DOUGHNUT,
				pieChartDataLabel = new()
				{
					dataLabelPosition = G.PieChartDataLabel.DataLabelPositionValues.SHOW,
					showCategoryName = true,
					showValue = true,
					separator = ". "
				}
			});
			//18
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
			//19
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_SMOOTH
			});
			//20
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_SMOOTH_MARKER
			});
			//21
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_STRIGHT
			});
			//22
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_STRIGHT_MARKER
			});
			//23
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(3, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.BUBBLE
			});
			Assert.IsTrue(true);
		}

		/// <summary>
		///
		/// </summary>
		[TestMethod]
		[TestCategory("Chart")]
		public void Add2016Charts()
		{
			X.DataCell[][] data = new X.DataCell[9][];
			data[0] = new X.DataCell[2];
			data[0][1] = new()
			{
				cellValue = "Series 1",
				dataType = X.CellDataType.STRING
			};
			for (int i = 1; i < 9; i++)
			{
				data[i] = new X.DataCell[2];
				data[i][0] = new X.DataCell()
				{
					cellValue = $"Category {i}",
					dataType = X.CellDataType.STRING
				};
				int val = (i % 2) == 0 ? -i : i;
				data[i][1] = new X.DataCell()
				{
					cellValue = $"{val}",
					dataType = X.CellDataType.NUMBER
				};
			}
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(data, new WaterfallChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
		}

		/// <summary>
		/// Add Blank Slide to the PPT
		/// </summary>
		[TestMethod]
		[TestCategory("Slide")]
		public void AddBlankSlide()
		{
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Add Single Chart to the Slide
		/// </summary>
		[TestMethod]
		[TestCategory("Chart")]
		public void AddDevChart()
		{
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(), new G.ColumnChartSetting<G.PresentationSetting>()
			{
				columnChartType = G.ColumnChartTypes.STACKED_3D,
				applicationSpecificSetting = new(),
				titleOptions = new()
				{
					title = "Dev Chart"
				}
			});
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Add Combo Chart to the Slide
		/// </summary>
		[TestMethod]
		[TestCategory("Chart")]
		public void AddComboChart()
		{
			G.ComboChartSetting<G.PresentationSetting> comboChartSetting = new()
			{
				applicationSpecificSetting = new(),
				titleOptions = new()
				{
					title = "Combo Chart"
				},
			};
			comboChartSetting.AddComboChartsSetting(new G.AreaChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
			comboChartSetting.AddComboChartsSetting(new G.BarChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
			comboChartSetting.AddComboChartsSetting(new G.ColumnChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
			comboChartSetting.AddComboChartsSetting(new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				isSecondaryAxis = true
			});
			comboChartSetting.AddComboChartsSetting(new G.PieChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
			});
			// comboChartSetting.AddComboChartsSetting(new G.ScatterChartSetting());
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(10), comboChartSetting);
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Add Picture to the slide
		/// </summary>
		[TestMethod]
		[TestCategory("Picture")]
		public void AddPicture()
		{
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new G.PictureSetting());
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddPicture("./TestFiles/tom_and_jerry.jpg", new G.PictureSetting());
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Add All type of sctter charts
		/// </summary>
		[TestMethod]
		[TestCategory("Chart")]
		public void AddScatterPlot()
		{
			//1
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				titleOptions = new()
				{
					title = "Scatter Plot"
				}
			});
			//2
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_SMOOTH,
				titleOptions = new()
				{
					title = "Scatter Smooth"
				}
			});
			//3
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_SMOOTH_MARKER,
				titleOptions = new()
				{
					title = "Scatter Smooth Marker"
				}
			});
			//4
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_STRIGHT,
				titleOptions = new()
				{
					title = "Scatter Stright",
					fontSize = 20
				}
			});
			//5
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(6, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.SCATTER_STRIGHT_MARKER,
				titleOptions = new()
				{
					fontColor = "FF0000",
					title = "Scatter Stright Marker"
				}
			});
			powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK).AddChart(CreateDataCellPayload(3, true), new G.ScatterChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				scatterChartType = G.ScatterChartTypes.BUBBLE,
				titleOptions = new()
				{
					isBold = true,
					title = "Bubble Chart"
				}
			});
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Add Table To the Slide
		/// </summary>
		[TestMethod]
		[TestCategory("Table")]
		public void AddTable()
		{
			Slide slide = powerPoint.AddSlide(PresentationConstants.SlideLayoutType.BLANK);
			slide.AddTable(CreateTableRowPayload(10), new TableSetting()
			{
				name = "New Table",
				widthType = TableSetting.WidthOptionValues.PERCENTAGE,
				tableColumnWidth = new() { 80, 20 },
				x = (uint)G.ConverterUtils.PixelsToEmu(10),
				y = (uint)G.ConverterUtils.PixelsToEmu(10),
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
			shapes1[0].ReplaceTextBox(slide.AddTextBox(new G.TextBoxSetting()
			{
				text = "This is text box Font Family",
				fontFamily = "Bernard MT Condensed"
			}));
			shapes2[0].ReplaceTextBox(new TextBox(new G.TextBoxSetting()
			{
				text = "This is text box background",
				horizontalAlignment = G.HorizontalAlignmentValues.CENTER,
				textBackground = "777777"
			}));
			shapes3[0].ReplaceTextBox(new TextBox(new G.TextBoxSetting()
			{
				text = "This is text box",
				fontSize = 22,
				isBold = true,
				textColor = "AAAAAA"
			}));
			powerPoint1.MoveSlideByIndex(4, 0);
			powerPoint1.SaveAs(string.Format("{1}/edit-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
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
			shape1[0].ReplaceChart(new Chart<G.PresentationSetting>(Slide, CreateDataCellPayload(),
			new G.ColumnChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				chartLegendOptions = new G.ChartLegendOptions()
				{
					isEnableLegend = false
				},
			}));
			shape2[0].ReplaceChart(new Chart<G.PresentationSetting>(Slide, CreateDataCellPayload(),
			new G.BarChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				chartLegendOptions = new G.ChartLegendOptions()
				{
					legendPosition = G.ChartLegendOptions.LegendPositionValues.RIGHT
				}
			}));
			shape3[0].ReplaceChart(new Chart<G.PresentationSetting>(Slide, CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new(),
				chartAxesOptions = new G.ChartAxesOptions()
				{
					isHorizontalAxesEnabled = false
				},
				chartGridLinesOptions = new G.ChartGridLinesOptions()
				{
					isMajorCategoryLinesEnabled = true,
					isMajorValueLinesEnabled = true,
					isMinorCategoryLinesEnabled = true,
					isMinorValueLinesEnabled = true,
				}
			}));
			shape4[0].ReplaceChart(new Chart<G.PresentationSetting>(Slide, CreateDataCellPayload(), new G.LineChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new()
			}));
			shape5[0].ReplaceChart(new Chart<G.PresentationSetting>(Slide, CreateDataCellPayload(), new G.AreaChartSetting<G.PresentationSetting>()
			{
				applicationSpecificSetting = new()
			}));
			shape6[0].ReplaceTextBox(new TextBox(new G.TextBoxSetting()
			{
				text = "Test"
			}));
			powerPoint1.SaveAs(string.Format("{1}/chart-{0}.pptx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
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

		private static X.DataCell[][] CreateDataCellPayload(int payloadSize = 5, bool IsValueAxis = false)
		{
			Random random = new();
			X.DataCell[][] data = new X.DataCell[payloadSize][];
			data[0] = new X.DataCell[payloadSize];
			for (int col = 1; col < payloadSize; col++)
			{
				data[0][col] = new X.DataCell
				{
					cellValue = $"Series {col}",
					dataType = X.CellDataType.STRING
				};
			}
			for (int row = 1; row < payloadSize; row++)
			{
				data[row] = new X.DataCell[payloadSize];
				data[row][0] = new X.DataCell
				{
					cellValue = $"Category {row}",
					dataType = X.CellDataType.STRING
				};
				for (int col = IsValueAxis ? 0 : 1; col < payloadSize; col++)
				{
					data[row][col] = new X.DataCell
					{
						cellValue = random.Next(1, 100).ToString(),
						dataType = X.CellDataType.NUMBER,
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
						fontSize=22,
						borderSettings = new(){
							leftBorder = new(){
								showBorder = false
							},
							topBorder = new(){
								showBorder = true,
								borderColor = "FF0000",
								width = 2
							},
							rightBorder = new(){
								showBorder = false
							},
							bottomBorder = new(){
								showBorder = true
							}
						},
						horizontalAlignment = G.HorizontalAlignmentValues.LEFT+ (i % 4)
					},
					new() {
						value = $"Row {i + 1}, Cell 2",
						textColor = "00FF00",
						isBold = true,
						borderSettings = new(){
							leftBorder = new(){
								showBorder = true
							},
							topBorder = new(){
								showBorder = false
							},
							rightBorder = new(){
								showBorder = true,
								borderColor= "0000FF",
								width = 1
							},
							bottomBorder = new(){
								showBorder = true
							}
						},
						isItalic = true,
						verticalAlignment = G.VerticalAlignmentValues.TOP +(i % 3)
					},
				}
				};
				data[i] = row;
			}
			return data;
		}


	}
}
