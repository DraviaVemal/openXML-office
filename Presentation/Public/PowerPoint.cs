// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System.IO;
using OpenXMLOffice.Global_2007;
using System.Reflection;

namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// PowerPoint class to work with PowerPoint files
	/// Update PrivacyProperties to set your usage statics data sharing
	/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
	/// </summary>
	public class PowerPoint : PrivacyProperties
	{
		private readonly Presentation presentation;
		/// <summary>
		/// Create New file in the system
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public PowerPoint(PowerPointProperties powerPointProperties = null)
		{
			presentation = new Presentation(powerPointProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source file will be cloned and released. hence can be replace by saveAs method if you want to update the same file.
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public PowerPoint(string filePath, bool isEditable, PowerPointProperties powerPointProperties = null)
		{
			isFileEdited = true;
			presentation = new Presentation(filePath, isEditable, powerPointProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source stream is copied and closed.
		/// Note : Make Clone in your source application if you want to retain the stream handle
		/// Read Privacy Details document at https://openxml-office.draviavemal.com/privacy-policy
		/// </summary>
		public PowerPoint(Stream Stream, bool IsEditable, PowerPointProperties powerPointProperties = null)
		{
			isFileEdited = true;
			presentation = new Presentation(Stream, IsEditable, powerPointProperties);
		}
		/// <summary>
		/// Add new slide to the presentation
		/// </summary>
		public Slide AddSlide(PresentationConstants.SlideLayoutType slideLayoutType)
		{
			return presentation.AddSlide(slideLayoutType);
		}
		/// <summary>
		/// Get Slide by index
		/// </summary>
		public Slide GetSlideByIndex(int SlideIndex)
		{
			return presentation.GetSlideByIndex(SlideIndex);
		}
		/// <summary>
		/// Get Slide count
		/// </summary>
		public int GetSlideCount()
		{
			return presentation.GetSlideCount();
		}
		/// <summary>
		/// Move slide by index
		/// </summary>
		public void MoveSlideByIndex(int SourceIndex, int TargetIndex)
		{
			presentation.MoveSlideByIndex(SourceIndex, TargetIndex);
		}
		/// <summary>
		/// Remove slide by index
		/// </summary>
		public void RemoveSlideByIndex(int SlideIndex)
		{
			presentation.RemoveSlideByIndex(SlideIndex);
		}
		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targeting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(string filePath)
		{
			SendAnonymousSaveStates(Assembly.GetExecutingAssembly().GetName());
			presentation.SaveAs(filePath);
		}

		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targeting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(Stream stream)
		{
			SendAnonymousSaveStates(Assembly.GetExecutingAssembly().GetName());
			presentation.SaveAs(stream);
		}
	}
}
