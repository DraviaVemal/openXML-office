// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Spreadsheet_2007;

namespace OpenXMLOffice.Tests
{
	/// <summary>
	/// Excel Test
	/// </summary>
	[TestClass]
	public class Spreadsheet
	{
		private static readonly Excel excel = new();

		private static readonly string resultPath = "../../testResult";

		/// <summary>
		/// Initialize excel Test
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
			excel.AddSheet();
		}

		/// <summary>
		/// Save the Test File After execution
		/// </summary>
		[ClassCleanup]
		public static void ClassCleanup()
		{
			excel.SaveAs(string.Format("{1}/test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
		}

		/// <summary>
		/// Add Sheet Test
		/// </summary>
		[TestMethod]
		public void AddSheet()
		{
			Worksheet worksheet = excel.AddSheet("TestSheet1");
			Assert.IsNotNull(worksheet);
			Assert.AreEqual("TestSheet1", worksheet.GetSheetName());
		}

		/// <summary>
		/// Rename Sheet Based on Index Test
		/// </summary>
		[TestMethod]
		public void RenameSheet()
		{
			Worksheet worksheet = excel.AddSheet("Sheet11");
			Assert.IsNotNull(worksheet);
			Assert.IsTrue(excel.RenameSheet("Sheet11", "Data1"));
		}

		/// <summary>
		/// Set Cell Test
		/// </summary>
		[TestMethod]
		public void SetColumn()
		{
			Worksheet worksheet = excel.AddSheet("Data3");
			Assert.IsNotNull(worksheet);
			worksheet.SetColumn("A1", new ColumnProperties()
			{
				width = 30
			});
			worksheet.SetColumn("C4", new ColumnProperties()
			{
				width = 30,
				bestFit = true
			});
			worksheet.SetColumn("G7", new ColumnProperties()
			{
				hidden = true
			});
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Set Row Test
		/// </summary>
		[TestMethod]
		public void SetRow()
		{
			Worksheet worksheet = excel.AddSheet("Data2");
			Assert.IsNotNull(worksheet);
			worksheet.SetRow("A1", new DataCell[6]{
				new(){
					cellValue = "test1",
					dataType = CellDataType.STRING
				},
				 new(){
					cellValue = "test2",
					dataType = CellDataType.STRING
				},
				 new(){
					cellValue = "test3",
					dataType = CellDataType.STRING
				},
				 new(){
					cellValue = "test4",
					dataType = CellDataType.STRING,
					styleSetting = new(){
						fontSize = 20
					}
				},
				 new(){
					cellValue = "2.51",
					dataType = CellDataType.NUMBER,
					styleSetting = new(){
						numberFormat = "00.000",
					}
				},new(){
					cellValue = "5.51",
					dataType = CellDataType.NUMBER,
					styleSetting = new(){
						numberFormat = "₹ #,##0.00;₹ -#,##0.00",
					}
				}
			}, new RowProperties()
			{
				height = 20
			});
			worksheet.SetRow("C1", new DataCell[1]{
				new(){
					cellValue = "Re Update",
					dataType = CellDataType.STRING
				}
			}, new RowProperties()
			{
				height = 30
			});
			Assert.IsTrue(true);
		}

		/// <summary>
		///
		/// </summary>
		[TestMethod]
		public void AddPicture()
		{
			Worksheet worksheet = excel.AddSheet("Data4");
			Assert.IsNotNull(worksheet);
			worksheet.SetRow("D3", new DataCell[1]{
				new(){
					cellValue = "Re Update",
					dataType = CellDataType.STRING
				}
			}, new RowProperties()
			{
				height = 30
			});
			worksheet.AddPicture("./TestFiles/tom_and_jerry.jpg", new()
			{
				imageType = ImageType.JPEG,
				fromCol = 6,
				fromRow = 6,
				toCol = 8,
				toRow = 8
			});
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Test All Chart Implementation
		/// </summary>
		[TestMethod]
		public void AddAllCharts()
		{
			Worksheet worksheet = excel.AddSheet("Area Chart");
			excel.RemoveSheet("Sheet1");
			int row = 0;
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new AreaChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new AreaChartSetting<ExcelSetting>()
			{
				areaChartType = AreaChartTypes.STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 5
					},
					to = new()
					{
						row = 41,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new AreaChartSetting<ExcelSetting>()
			{
				areaChartType = AreaChartTypes.PERCENT_STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 42,
						column = 5
					},
					to = new()
					{
						row = 62,
						column = 20
					}
				}
			});
			row = 0;
			worksheet = excel.AddSheet("Bar Chart");
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new BarChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new BarChartSetting<ExcelSetting>()
			{
				barChartType = BarChartTypes.STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 5
					},
					to = new()
					{
						row = 41,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new BarChartSetting<ExcelSetting>()
			{
				barChartType = BarChartTypes.PERCENT_STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 42,
						column = 5
					},
					to = new()
					{
						row = 62,
						column = 20
					}
				}
			});
			row = 0;
			worksheet = excel.AddSheet("Column Chart");
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new ColumnChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new ColumnChartSetting<ExcelSetting>()
			{
				columnChartType = ColumnChartTypes.STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 5
					},
					to = new()
					{
						row = 41,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new ColumnChartSetting<ExcelSetting>()
			{
				columnChartType = ColumnChartTypes.PERCENT_STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 42,
						column = 5
					},
					to = new()
					{
						row = 62,
						column = 20
					}
				}
			});
			row = 0;
			worksheet = excel.AddSheet("Line Chart");
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				lineChartType = LineChartTypes.STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 5
					},
					to = new()
					{
						row = 41,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				lineChartType = LineChartTypes.PERCENT_STACKED,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 42,
						column = 5
					},
					to = new()
					{
						row = 62,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				lineChartType = LineChartTypes.CLUSTERED_MARKER,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 21
					},
					to = new()
					{
						row = 20,
						column = 36
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				lineChartType = LineChartTypes.STACKED_MARKER,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 21
					},
					to = new()
					{
						row = 41,
						column = 36
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				lineChartType = LineChartTypes.PERCENT_STACKED_MARKER,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 42,
						column = 21
					},
					to = new()
					{
						row = 62,
						column = 36
					}
				}
			});
			row = 0;
			worksheet = excel.AddSheet("Pie Chart");
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new PieChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new PieChartSetting<ExcelSetting>()
			{
				pieChartType = PieChartTypes.DOUGHNUT,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 5
					},
					to = new()
					{
						row = 41,
						column = 20
					}
				}
			});
			row = 0;
			worksheet = excel.AddSheet("Combo Chart");
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			ComboChartSetting<ExcelSetting> comboChartSetting = new()
			{
				secondaryAxisPosition = AxisPosition.TOP,
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 21,
						column = 5
					},
					to = new()
					{
						row = 41,
						column = 20
					}
				}
			};
			comboChartSetting.AddComboChartsSetting(new LineChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
			});
			comboChartSetting.AddComboChartsSetting(new BarChartSetting<ExcelSetting>()
			{
				isSecondaryAxis = true,
				applicationSpecificSetting = new()
			});
			comboChartSetting.AddComboChartsSetting(new ColumnChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, comboChartSetting);
		}

		/// <summary>
		/// Open and close Excel without editing
		/// </summary>
		[TestMethod]
		public void OpenExistingexcelNonEdit()
		{
			Excel excel1 = new("./TestFiles/basic_test.xlsx", false);
			excel1.SaveAs(string.Format("{1}/ReadEdit-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
			Assert.IsTrue(true);
		}

		/// <summary>
		/// Test existing file
		/// </summary>
		[TestMethod]
		public void OpenExistingExcel()
		{
			Excel excel1 = new("./TestFiles/basic_test.xlsx", true);
			Worksheet worksheet = excel1.AddSheet("AreaChart");
			int row = 0;
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new AreaChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			worksheet = excel1.AddSheet("LineChart");
			row = 0;
			CreateDataCellPayload().ToList().ForEach(rowData =>
			{
				worksheet.SetRow(ConverterUtils.ConvertToExcelCellReference(++row, 1), rowData, new());
			});
			worksheet.AddChart(new()
			{
				cellIdStart = "A1",
				cellIdEnd = "D4"
			}, new LineChartSetting<ExcelSetting>()
			{
				applicationSpecificSetting = new()
				{
					from = new()
					{
						row = 5,
						column = 5
					},
					to = new()
					{
						row = 20,
						column = 20
					}
				}
			});
			excel1.SaveAs(string.Format("{1}/Edit-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
			Assert.IsTrue(true);
		}

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
						}
					};
				}
			}
			return data;
		}

	}
}
