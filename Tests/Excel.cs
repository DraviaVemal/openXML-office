// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Excel;

namespace OpenXMLOffice.Tests
{
	/// <summary>
	/// Excel Test
	/// </summary>
	[TestClass]
	public class Excel
	{
		private static Spreadsheet spreadsheet = new(new MemoryStream());


		/// <summary>
		/// Save the Test File After execution
		/// </summary>
		[ClassCleanup]
		public static void ClassCleanup()
		{
			spreadsheet.Save();
		}

		/// <summary>
		/// Initialize Spreadsheet Test
		/// </summary>
		/// <param name="context">
		/// </param>
		[ClassInitialize]
		public static void ClassInitialize(TestContext context)
		{
			spreadsheet = new(string.Format("../../test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")));
		}

		/// <summary>
		/// Add Sheet Test
		/// </summary>
		[TestMethod]
		public void AddSheet()
		{
			Worksheet worksheet = spreadsheet.AddSheet();
			Assert.IsNotNull(worksheet);
			Assert.AreEqual("Sheet1", worksheet.GetSheetName());
		}

		/// <summary>
		/// Rename Sheet Based on Index Test
		/// </summary>
		[TestMethod]
		public void RenameSheet()
		{
			Worksheet worksheet = spreadsheet.AddSheet("Sheet11");
			Assert.IsNotNull(worksheet);
			Assert.IsTrue(spreadsheet.RenameSheet("Sheet11", "Data1"));
		}

		/// <summary>
		/// Set Cell Test
		/// </summary>
		[TestMethod]
		public void SetColumn()
		{
			Worksheet worksheet = spreadsheet.AddSheet("Data3");
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
			Worksheet worksheet = spreadsheet.AddSheet("Data2");
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
		/// Create Xslx File Based on File Test
		/// </summary>
		[TestMethod]
		public void SheetConstructorFile()
		{
			Spreadsheet spreadsheet1 = new("../try.xlsx");
			Assert.IsNotNull(spreadsheet1);
			spreadsheet1.Save();
			File.Delete("../try.xlsx");
		}

		/// <summary>
		/// Create Xslx File Based on Stream Test
		/// </summary>
		[TestMethod]
		public void SheetConstructorStream()
		{
			MemoryStream memoryStream = new();
			Spreadsheet spreadsheet1 = new(memoryStream);
			Assert.IsNotNull(spreadsheet1);
		}


	}
}
