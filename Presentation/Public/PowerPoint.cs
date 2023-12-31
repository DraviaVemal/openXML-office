using DocumentFormat.OpenXml;

namespace OpenXMLOffice.Presentation
{
    public class PowerPoint
    {
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
        public PowerPoint(string filePath, PresentationProperties? powerPointProperties = null)
        {
            presentation = new(filePath, powerPointProperties);
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
        public PowerPoint(string filePath, bool isEditable, PresentationProperties? powerPointProperties = null)
        {
            presentation = new(filePath, isEditable, powerPointProperties);
        }

        /// <summary>
        /// Works with in memory object can be saved to file at later point
        /// </summary>
        /// <param name="filePath">
        /// </param>
        /// <param name="powerPointProperties">
        /// </param>
        public PowerPoint(Stream stream, PresentationProperties? powerPointProperties = null)
        {
            presentation = new(stream, powerPointProperties);
        }

        /// <summary>
        /// Works with in memory object can be saved to file at later point
        /// </summary>
        /// <param name="filePath">
        /// </param>
        /// <param name="powerPointProperties">
        /// </param>
        public PowerPoint(Stream stream, bool isEditable, PresentationProperties? powerPointProperties = null)
        {
            presentation = new(stream, isEditable, powerPointProperties);
        }

        #endregion Public Constructors

        #region Public Methods

        public Slide AddSlide(PresentationConstants.SlideLayoutType slideLayoutType)
        {
            return presentation.AddSlide(slideLayoutType);
        }

        public Slide GetSlideByIndex(int SlideIndex)
        {
            return presentation.GetSlideByIndex(SlideIndex);
        }

        public void MoveSlideByIndex(int SourceIndex, int TargetIndex)
        {
            presentation.MoveSlideByIndex(SourceIndex, TargetIndex);
        }

        public void RemoveSlideByIndex(int SlideIndex)
        {
            presentation.RemoveSlideByIndex(SlideIndex);
        }

        public void Save()
        {
            presentation.Save();
        }

        public void SaveAs(string filePath)
        {
            presentation.SaveAs(filePath);
        }

        #endregion Public Methods
    }
}