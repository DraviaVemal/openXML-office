// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.Reflection;
using System.Xml;

namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	///
	/// </summary>
	public class XMLHelper
	{
		/// <summary>
		///
		/// </summary>
		public static void AddOrUpdateCoreProperties(Stream stream)
		{
			try
			{
				string timeStamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
				using XmlTextWriter writer = new(stream, System.Text.Encoding.UTF8);
				writer.Formatting = Formatting.Indented;
				writer.Indentation = 2;
				writer.WriteStartDocument(true);
				writer.WriteStartElement("cp", "coreProperties", "https://schemas.openxmlformats.org/package/2006/metadata/core-properties");
				writer.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
				writer.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
				writer.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
				writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
				writer.WriteElementString("dc:creator", "OpenXML-Office");
				writer.WriteElementString("cp:lastModifiedBy", "OpenXML-Office");
				writer.WriteStartElement("dcterms:created");
				writer.WriteAttributeString("xsi:type", "dcterms:W3CDTF");
				writer.WriteString(timeStamp);
				writer.WriteEndElement();
				writer.WriteStartElement("dcterms:modified");
				writer.WriteAttributeString("xsi:type", "dcterms:W3CDTF");
				writer.WriteString(timeStamp);
				writer.WriteEndElement();
				writer.WriteEndElement();
				// <dcterms:created xsi:type=\"dcterms:W3CDTF\">{1}</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">{1}</dcterms:modified></cp:coreProperties>", "OpenXML-Office", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")));
				writer.Flush();
				writer.Dispose();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Core Property Error" + ex.Message);
			}
		}

		/// <summary>
		///
		/// </summary>
		public static void AddOrUpdateOpenXMLProperties(Stream stream)
		{
			try
			{
				using XmlTextWriter writer = new(stream, System.Text.Encoding.UTF8);
				writer.Formatting = Formatting.Indented;
				writer.Indentation = 2;
				writer.WriteStartDocument(true);
				writer.WriteStartElement("OpenXML-Office");
				writer.WriteElementString("Created", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
				writer.WriteElementString("Version", Assembly.GetExecutingAssembly().GetName().Version!.ToString());
				writer.WriteEndElement();
				writer.Flush();
				writer.Dispose();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Additional Property Error" + ex.Message);
			}
		}
	}
}
