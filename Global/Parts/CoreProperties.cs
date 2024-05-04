// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
namespace OpenXMLOffice.Global_2007
{
	/// <summary>
	///
	/// </summary>
	public class CoreProperties
	{
		/// <summary>
		///
		/// </summary>
		public static void AddCoreProperties(Stream stream, CorePropertiesModel corePropertiesModel)
		{
			try
			{
				string timeStamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
				using (XmlTextWriter writer = new XmlTextWriter(stream, System.Text.Encoding.UTF8))
				{
					writer.Formatting = Formatting.Indented;
					writer.Indentation = 2;
					writer.WriteStartDocument(true);
					writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
					writer.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
					writer.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
					writer.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
					writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
					if (corePropertiesModel.title != null)
					{
						writer.WriteElementString("dc:title", corePropertiesModel.title);
					}
					if (corePropertiesModel.subject != null)
					{
						writer.WriteElementString("dc:subject", corePropertiesModel.subject);
					}
					if (corePropertiesModel.description != null)
					{
						writer.WriteElementString("dc:description", corePropertiesModel.description);
					}
					if (corePropertiesModel.tags != null)
					{
						writer.WriteElementString("cp:keywords", corePropertiesModel.tags);
					}
					if (corePropertiesModel.category != null)
					{
						writer.WriteElementString("cp:category", corePropertiesModel.category);
					}
					writer.WriteElementString("dc:creator", corePropertiesModel.creator);
					writer.WriteElementString("cp:lastModifiedBy", corePropertiesModel.creator);
					writer.WriteStartElement("dcterms:created");
					writer.WriteAttributeString("xsi:type", "dcterms:W3CDTF");
					writer.WriteString(timeStamp);
					writer.WriteEndElement();
					writer.WriteStartElement("dcterms:modified");
					writer.WriteAttributeString("xsi:type", "dcterms:W3CDTF");
					writer.WriteString(timeStamp);
					writer.WriteEndElement();
					writer.WriteEndElement();
					writer.Flush();
					writer.Dispose();
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Core Property Error" + ex.Message);
			}
		}
		/// <summary>
		///
		/// </summary>
		public static void UpdateModifiedDetails(Stream stream, CorePropertiesModel corePropertiesModel = null)
		{
			XDocument doc = XDocument.Load(stream);
			stream.Position = 0;
			stream.SetLength(0);
			XElement lastModifiedByElement = doc.Descendants().First(e => e.Name.LocalName == "lastModifiedBy");
			XElement modifiedElement = doc.Descendants().First(e => e.Name.LocalName == "modified");
			lastModifiedByElement.SetValue(corePropertiesModel.creator);
			modifiedElement.SetValue(DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"));
			doc.Save(stream);
		}
	}
}
