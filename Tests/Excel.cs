using OpenXMLOffice.Excel;

namespace OpenXMLOffice.Tests;

[TestClass]
public class Excel
{
    private static Spreadsheet spreadsheet = new(new MemoryStream());

    [ClassInitialize]
    public static void ClassInitialize(TestContext context)
    {
        spreadsheet = new(string.Format("../../test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")), DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
    }

    [ClassCleanup]
    public static void ClassCleanup()
    {
        spreadsheet.Save();
    }

    [TestMethod]
    public void SheetConstructorFile()
    {
        Spreadsheet spreadsheet1 = new("../try.xlsx", DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        Assert.IsNotNull(spreadsheet1);
        spreadsheet1.Save();
        File.Delete("../try.xlsx");
    }

    [TestMethod]
    public void SheetConstructorStream()
    {
        MemoryStream memoryStream = new();
        Spreadsheet spreadsheet1 = new(memoryStream);
        Assert.IsNotNull(spreadsheet1);
    }


    [TestMethod]
    public void AddSheet()
    {
        Worksheet worksheet = spreadsheet.AddSheet();
        Assert.IsNotNull(worksheet);
        Assert.AreEqual("Sheet1", worksheet.GetSheetName());
    }


    [TestMethod]
    public void RenameSheet()
    {
        Worksheet worksheet = spreadsheet.AddSheet("Sheet11");
        Assert.IsNotNull(worksheet);
        Assert.IsTrue(spreadsheet.RenameSheet("Sheet11", "Data1"));
    }


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
            height = 20
        });
        worksheet.SetRow("C1", new DataCell[1]{
            new(){
                CellValue = "Re Update",
                DataType = CellDataType.STRING
            }
        }, new RowProperties()
        {
            height = 30
        });
        Assert.IsTrue(true);
    }

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

}