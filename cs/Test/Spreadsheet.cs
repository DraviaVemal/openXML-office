// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using draviavemal.openxml_office.global_2007;
using draviavemal.openxml_office.spreadsheet_2007;

namespace openxmloffice.Tests
{
    /// <summary>
    /// Excel Test
    /// </summary>
    [TestClass]
    public class Spreadsheet
    {
        private static readonly string resultPath = "../../TestOutputFiles";
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
            PrivacyProperties.ShareComponentRelatedDetails = false;
            PrivacyProperties.ShareIpGeoLocation = false;
            PrivacyProperties.ShareOsDetails = false;
            PrivacyProperties.SharePackageRelatedDetails = false;
            PrivacyProperties.ShareUsageCounterDetails = false;
        }
        /// <summary>
        /// Save the Test File After execution
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            // excel.SaveAs(string.Format("{1}/test-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
        }
        /// <summary>
        /// 
        /// </summary>
        // [TestMethod]
        // public void BlankFile()
        // {
        //     Excel excel2 = new(new()
        //     {
        //         IsInMemory = false
        //     });
        //     excel2.SaveAs(string.Format("{1}/Blank-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
        // }

        /// <summary>
        /// Test existing file
        /// </summary>
        [TestMethod]
        public void OpenExistingExcelStyleString()
        {
            Excel excel1 = new("/home/draviavemal/repo/OpenXML-Office/cs/Test/TestFiles/basic_test.xlsx", new()
            {
                IsInMemory = false
            });
            excel1.SaveAs(string.Format("{1}/EditStyle-{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
            Assert.IsTrue(true);
        }
    }
}
