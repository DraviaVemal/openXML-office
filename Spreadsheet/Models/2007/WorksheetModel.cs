// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
namespace OpenXMLOffice.Spreadsheet_2007
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
		/// <summary>
		/// Gets or sets a value indicating whether the column width should be automatically
		/// adjusted to fit the contents.
		/// </summary>
		public bool bestFit;
		/// <summary>
		/// Gets or sets a value indicating whether the column is hidden.
		/// </summary>
		public bool hidden;
		/// <summary>
		/// Gets or sets the width of the column.
		/// </summary>
		public double? width;
	}
	/// <summary>
	/// Represents a data cell in a worksheet.
	/// </summary>
	public class DataCell
	{
		/// <summary>
		/// Gets or sets the value of the cell.
		/// </summary>
		public string? cellValue;
		/// <summary>
		/// Gets or sets the data type of the cell.
		/// </summary>
		public CellDataType dataType;
		/// <summary>
		/// It is highgly recomended to use styleId instead of styleSetting.
		/// </summary>
		/// warning: styleSetting will be ignored if styleId is not null
		public CellStyleSetting? styleSetting = new();
		/// <summary>
		/// Use file level styleId instead of styleSetting.
		/// Can get the styleId from spreadsheet.GetCellStyleId(CellStyleSetting)
		/// </summary>
		public uint? styleId;
	}
	/// <summary>
	/// Represents a record in a worksheet.
	/// </summary>
	public class StringRecord
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="StringRecord"/> class with the specified value.
		/// </summary>
		/// <param name="Value">
		/// The value of the record.
		/// </param>
		public StringRecord(string Value)
		{
			this.Value = Value;
		}
		/// <summary>
		/// Gets or sets the ID of the record.
		/// </summary>
		public int Id { get; set; }
		/// <summary>
		/// Gets or sets the value of the record.
		/// </summary>
		public string Value { get; set; }
	}
	/// <summary>
	/// Represents the properties of a row in a worksheet.
	/// </summary>
	public class RowProperties
	{
		/// <summary>
		/// Gets or sets the height of the row.
		/// </summary>
		public double? height;
		/// <summary>
		/// Gets or sets a value indicating whether the row is hidden.
		/// </summary>
		public bool hidden;
	}
}
