/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

namespace OpenXMLOffice.Excel
{
    public enum CellDataType
    {
        DATE,
        NUMBER,
        STRING
    }

    public class ColumnProperties
    {
        #region Public Fields

        public bool BestFit;
        public bool Hidden;
        public double? Width;

        #endregion Public Fields
    }

    public class DataCell
    {
        #region Public Fields

        public string? CellValue;
        public CellDataType DataType;
        public string? NumberFormatting;
        public int? StyleId;

        #endregion Public Fields
    }

    public class RowProperties
    {
        #region Public Fields

        public double? Height;
        public bool Hidden;

        #endregion Public Fields
    }
    
    public class Record
    {
        #region Public Constructors

        public Record(string Value)
        {
            this.Value = Value;
        }

        #endregion Public Constructors

        #region Public Properties

        public int Id { get; set; }
        public string Value { get; set; }

        #endregion Public Properties
    }
}