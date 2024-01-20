// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

namespace OpenXMLOffice.Presentation {
    /// <summary>
    /// PowerPoint class to work with PowerPoint files
    /// </summary>
    public class PowerPoint {
        #region Private Fields

        private readonly Presentation presentation;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Create New file in the system
        /// </summary>
        /// <param name="filePath">
        /// </param>
        /// <param name="powerPointProperties">
        /// </param>
        public PowerPoint(string filePath,PresentationProperties? powerPointProperties = null) {
            presentation = new(filePath,powerPointProperties);
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
        public PowerPoint(string filePath,bool isEditable,PresentationProperties? powerPointProperties = null) {
            presentation = new(filePath,isEditable,powerPointProperties);
        }

        /// <summary>
        /// Works with in memory object can be saved to file at later point
        /// </summary>
        public PowerPoint(Stream Stream,PresentationProperties? powerPointProperties = null) {
            presentation = new(Stream,powerPointProperties);
        }

        /// <summary>
        /// Works with in memory object can be saved to file at later point
        /// </summary>
        public PowerPoint(Stream Stream,bool IsEditable,PresentationProperties? powerPointProperties = null) {
            presentation = new(Stream,IsEditable,powerPointProperties);
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Add new slide to the presentation
        /// </summary>
        /// <param name="slideLayoutType">
        /// </param>
        /// <returns>
        /// </returns>
        public Slide AddSlide(PresentationConstants.SlideLayoutType slideLayoutType) {
            return presentation.AddSlide(slideLayoutType);
        }

        /// <summary>
        /// Get Slide by index
        /// </summary>
        /// <param name="SlideIndex">
        /// </param>
        /// <returns>
        /// </returns>
        public Slide GetSlideByIndex(int SlideIndex) {
            return presentation.GetSlideByIndex(SlideIndex);
        }

        /// <summary>
        /// Get Slide count
        /// </summary>
        /// <returns>
        /// </returns>
        public int GetSlideCount() {
            return presentation.GetSlideCount();
        }

        /// <summary>
        /// Move slide by index
        /// </summary>
        /// <param name="SourceIndex">
        /// </param>
        /// <param name="TargetIndex">
        /// </param>
        public void MoveSlideByIndex(int SourceIndex,int TargetIndex) {
            presentation.MoveSlideByIndex(SourceIndex,TargetIndex);
        }

        /// <summary>
        /// Remove slide by index
        /// </summary>
        /// <param name="SlideIndex">
        /// </param>
        public void RemoveSlideByIndex(int SlideIndex) {
            presentation.RemoveSlideByIndex(SlideIndex);
        }

        /// <summary>
        /// Save the file
        /// </summary>
        public void Save() {
            presentation.Save();
        }

        /// <summary>
        /// Save the file as new file
        /// </summary>
        /// <param name="filePath">
        /// </param>
        public void SaveAs(string filePath) {
            presentation.SaveAs(filePath);
        }

        #endregion Public Methods
    }
}