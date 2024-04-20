// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System.IO;
namespace OpenXMLOffice.Presentation_2007
{
	/// <summary>
	/// PowerPoint class to work with PowerPoint files
	/// </summary>
	public class PowerPoint
	{
		private readonly Presentation presentation;
		/// <summary>
		/// Create New file in the system
		/// </summary>
		public PowerPoint(PowerPointProperties powerPointProperties = null)
		{
			presentation = new Presentation(powerPointProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source file will be cloned and released. hence can be replace by saveAs method if you want to update the same file.
		/// </summary>
		public PowerPoint(string filePath, bool isEditable, PowerPointProperties powerPointProperties = null)
		{
			presentation = new Presentation(filePath, isEditable, powerPointProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point.
		/// Source stream is copied and closed.
		/// Note : Make Clone in your source application if you want to retain the stream handle
		/// </summary>
		public PowerPoint(Stream Stream, bool IsEditable, PowerPointProperties powerPointProperties = null)
		{
			presentation = new Presentation(Stream, IsEditable, powerPointProperties);
		}
		/// <summary>
		/// Add new slide to the presentation
		/// </summary>
		/// <param name="slideLayoutType">
		/// </param>
		/// <returns>
		/// </returns>
		public Slide AddSlide(PresentationConstants.SlideLayoutType slideLayoutType)
		{
			return presentation.AddSlide(slideLayoutType);
		}
		/// <summary>
		/// Get Slide by index
		/// </summary>
		/// <param name="SlideIndex">
		/// </param>
		/// <returns>
		/// </returns>
		public Slide GetSlideByIndex(int SlideIndex)
		{
			return presentation.GetSlideByIndex(SlideIndex);
		}
		/// <summary>
		/// Get Slide count
		/// </summary>
		/// <returns>
		/// </returns>
		public int GetSlideCount()
		{
			return presentation.GetSlideCount();
		}
		/// <summary>
		/// Move slide by index
		/// </summary>
		/// <param name="SourceIndex">
		/// </param>
		/// <param name="TargetIndex">
		/// </param>
		public void MoveSlideByIndex(int SourceIndex, int TargetIndex)
		{
			presentation.MoveSlideByIndex(SourceIndex, TargetIndex);
		}
		/// <summary>
		/// Remove slide by index
		/// </summary>
		/// <param name="SlideIndex">
		/// </param>
		public void RemoveSlideByIndex(int SlideIndex)
		{
			presentation.RemoveSlideByIndex(SlideIndex);
		}
		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targetting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(string filePath)
		{
			presentation.SaveAs(filePath);
		}
		/// <summary>
		/// Even on edit file OpenXML-Office Will clone the source and work on top of it to protect the integrity of source file.
		/// You can save the document at the end of lifecycle targetting the edit file to update or new file.
		/// This is supported for both file path and data stream
		/// </summary>
		public void SaveAs(Stream stream)
		{
			presentation.SaveAs(stream);
		}
	}
}
