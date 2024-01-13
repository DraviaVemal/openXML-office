/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// Represents the data type of a cell in a worksheet.
    /// </summary>
    public enum CellDataType
    {
        /// <summary>
        /// Represents a date cell.
        /// </summary>
        DATE,
        /// <summary>
        /// Represents a number cell.
        /// </summary>
        NUMBER,
        /// <summary>
        /// Represents a string cell.
        /// </summary>
        STRING
    }

    /// <summary>
    /// Represents the properties of a column in a worksheet.
    /// </summary>
    public class ColumnProperties
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets a value indicating whether the column width should be automatically adjusted to fit the contents.
        /// </summary>
        public bool BestFit;

        /// <summary>
        /// Gets or sets a value indicating whether the column is hidden.
        /// </summary>
        public bool Hidden;

        /// <summary>
        /// Gets or sets the width of the column.
        /// </summary>
        public double? Width;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents a data cell in a worksheet.
    /// </summary>
    public class DataCell
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets the value of the cell.
        /// </summary>
        public string? CellValue;

        /// <summary>
        /// Gets or sets the data type of the cell.
        /// </summary>
        public CellDataType DataType;

        /// <summary>
        /// Gets or sets the number format of the cell.
        /// </summary>
        public string NumberFormat = "General";

        /// <summary>
        /// Gets or sets the style ID of the cell.
        /// </summary>
        public int? StyleId;

        #endregion Public Fields
    }

    /// <summary>
    /// Represents a record in a worksheet.
    /// </summary>
    public class Record
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Record"/> class with the specified value.
        /// </summary>
        /// <param name="Value">The value of the record.</param>
        public Record(string Value)
        {
            this.Value = Value;
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Gets or sets the ID of the record.
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Gets or sets the value of the record.
        /// </summary>
        public string Value { get; set; }

        #endregion Public Properties
    }

    /// <summary>
    /// Represents the properties of a row in a worksheet.
    /// </summary>
    public class RowProperties
    {
        #region Public Fields

        /// <summary>
        /// Gets or sets the height of the row.
        /// </summary>
        public double? Height;

        /// <summary>
        /// Gets or sets a value indicating whether the row is hidden.
        /// </summary>
        public bool Hidden;

        #endregion Public Fields
    }
}