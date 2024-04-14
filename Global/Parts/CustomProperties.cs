// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.Reflection;
using System.Xml;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	///
	/// </summary>
	public class CustomProperties
	{
		/// <summary>
		///
		/// </summary>
		public static void AddOrUpdateOpenXMLCustomProperties(Stream stream)
		{
			try
			{
				using (XmlTextWriter writer = new XmlTextWriter(stream, System.Text.Encoding.UTF8))
				{
					writer.Formatting = Formatting.Indented;
					writer.Indentation = 2;
					writer.WriteStartDocument(true);
					writer.WriteStartElement("OpenXML-Office");
					writer.WriteElementString("Created", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
					writer.WriteElementString("Version", Assembly.GetExecutingAssembly().GetName().Version.ToString());
					writer.WriteEndElement();
					writer.Flush();
					writer.Dispose();
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Additional Property Error" + ex.Message);
			}
		}
	}
}
