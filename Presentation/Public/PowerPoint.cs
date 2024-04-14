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
		public PowerPoint(PresentationProperties powerPointProperties = null)
		{
			presentation = new Presentation(powerPointProperties);
		}
		/// <summary>
		/// Open and work with existing file
		/// </summary>
		/// <param name="filePath">
		/// </param>
		/// <param name="isEditable">
		/// </param>
		/// <param name="powerPointProperties">
		/// </param>
		public PowerPoint(string filePath, bool isEditable, PresentationProperties powerPointProperties = null)
		{
			presentation = new Presentation(filePath, isEditable, powerPointProperties);
		}
		/// <summary>
		/// Works with in memory object can be saved to file at later point
		/// </summary>
		public PowerPoint(Stream Stream, bool IsEditable, PresentationProperties powerPointProperties = null)
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
		/// Save the file as new file
		/// </summary>
		/// <param name="filePath">
		/// </param>
		public void SaveAs(string filePath)
		{
			presentation.SaveAs(filePath);
		}
	}
}
