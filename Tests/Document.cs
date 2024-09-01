// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using OpenXMLOffice.Global_2007;
using OpenXMLOffice.Document_2007;
namespace OpenXMLOffice.Tests
{
	/// <summary>
	/// Excel Test
	/// </summary>
	[TestClass]
	public class Document
	{
		private static readonly Word word = new(new WordProperties
		{
			coreProperties = new()
			{
				title = "Test File",
				creator = "OpenXML-Office",
				subject = "Test Subject",
				tags = "Test",
				category = "Test Category",
				description = "Describe the test file"
			}
		});
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
			word.SaveAs(string.Format("{1}/test-{0}.docx", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), resultPath));
		}

		/// <summary>
		/// 
		/// </summary>
		[TestMethod]
		public void emptyPage()
		{
			Assert.IsTrue(true);
		}
	}
}
