/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using OpenXMLOffice.Excel;

namespace OpenXMLOffice.Tests
{
    /// <summary>
    /// Excel Test
    /// </summary>
    [TestClass]
    public class Excel
    {
        #region Private Fields

        private static Spreadsheet spreadsheet = new(new MemoryStream());

        #endregion Private Fields

        #region Public Methods

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
                Width = 30
            });
            worksheet.SetColumn("C4", new ColumnProperties()
            {
                Width = 30,
                BestFit = true
            });
            worksheet.SetColumn("G7", new ColumnProperties()
            {
                Hidden = true
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
            worksheet.SetRow("A1", new DataCell[5]{
                new(){
                    CellValue = "test1",
                    DataType = CellDataType.STRING
                },
                 new(){
                    CellValue = "test2",
                    DataType = CellDataType.STRING
                },
                 new(){
                    CellValue = "test3",
                    DataType = CellDataType.STRING
                },
                 new(){
                    CellValue = "test4",
                    DataType = CellDataType.STRING
                },
                 new(){
                    CellValue = "test5",
                    DataType = CellDataType.STRING
                }
            }, new RowProperties()
            {
                Height = 20
            });
            worksheet.SetRow("C1", new DataCell[1]{
                new(){
                    CellValue = "Re Update",
                    DataType = CellDataType.STRING
                }
            }, new RowProperties()
            {
                Height = 30
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

        #endregion Public Methods
    }
}