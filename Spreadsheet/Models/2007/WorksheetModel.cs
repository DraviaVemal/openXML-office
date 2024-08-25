// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using DocumentFormat.OpenXml;
using OpenXMLOffice.Global_2007;

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
		STRING,
		/// <summary>
		/// 
		/// </summary>
		FORMULA
	}
	/// <summary>
	/// View setting for each worksheet
	/// </summary>
	public enum WorkSheetViewsValues
	{
		/// <summary>
		/// 
		/// </summary>
		NORMAL,
		/// <summary>
		/// 
		/// </summary>
		PAGE_BREAK_PREVIEW,
		/// <summary>
		/// 
		/// </summary>
		PAGE_LAYOUT,
	}
	/// <summary>
	/// 
	/// </summary>
	public enum WorkSheetViewFreezeDirection
	{
		/// <summary>
		/// 
		/// </summary>
		TOP,
		/// <summary>
		/// 
		/// </summary>
		LEFT
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
	/// Represents a each column cell in a row.
	/// </summary>
	public class ColumnCell
	{
		/// <summary>
		/// Gets or sets the value of the cell.
		/// </summary>
		public string cellValue;
		/// <summary>
		/// Gets or sets the data type of the cell.
		/// </summary>
		public CellDataType dataType;
		/// <summary>
		/// Cell hyperlink
		/// </summary>
		public HyperlinkProperties hyperlinkProperties;
		/// <summary>
		/// It is highly recommended to use styleId instead of styleSetting.
		/// </summary>
		/// warning: styleSetting will be ignored if styleId is not null
		public CellStyleSetting styleSetting = new CellStyleSetting();
		/// <summary>
		/// Use file level styleId instead of styleSetting.
		/// Can get the styleId from spreadsheet.GetCellStyleId(CellStyleSetting)
		/// </summary>
		public uint? styleId;
	}
	/// <summary>
	/// Represents a string record in a worksheet.
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
	/// Represents a formula record in a worksheet.
	/// </summary>
	public class CalculationRecord
	{
		/// <summary>
		/// 
		/// </summary>
		public CalculationRecord(string CellId, int SheetIndex)
		{
			this.CellId = CellId;
			this.SheetIndex = SheetIndex;
		}
		/// <summary>
		/// Gets or sets the ID of the record.
		/// </summary>
		public int Id { get; set; }
		/// <summary>
		/// Gets or sets the CellId of the record.
		/// </summary>
		public string CellId { get; set; }
		/// <summary>
		/// Gets or sets sheet index for the formula.
		/// </summary>
		public int SheetIndex { get; set; }
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
	/// <summary>
	/// Merge Cell Options
	/// </summary>
	public class MergeCellRange
	{
		/// <summary>
		/// 
		/// </summary>
		public string topLeftCell;
		/// <summary>
		/// 
		/// </summary>
		public string bottomRightCell;
	}
	/// <summary>
	/// 
	/// </summary>
	public class WorkSheetViewOption
	{
		private uint zoomScale = 100;
		/// <summary>
		/// Set screen Zoom level.
		/// Range between 10 to 400
		/// </summary>
		public uint ZoomScale
		{
			get
			{
				return zoomScale;
			}
			set
			{
				if (value < 10)
				{
					zoomScale = 10;
				}
				else if (value > 400)
				{
					zoomScale = 400;
				}
				else
				{
					zoomScale = value;
				}
			}
		}
		/// <summary>
		/// 
		/// </summary>
		public WorkSheetViewsValues workSheetViewsValue = WorkSheetViewsValues.NORMAL;
		/// <summary>
		/// 
		/// </summary>
		public bool showGridLine = true;
		/// <summary>
		/// 
		/// </summary>
		public bool showFormula = true;
		/// <summary>
		/// 
		/// </summary>
		public bool showGridLines = true;
		/// <summary>
		/// 
		/// </summary>
		public bool showRowColHeaders = true;
		/// <summary>
		/// 
		/// </summary>
		public bool showRuler = false;
		// /// <summary>
		// /// 
		// /// </summary>
		// public uint freezeCount = 0;
		// /// <summary>
		// /// 
		// /// </summary>
		// public WorkSheetViewFreezeDirection workSheetViewFreezeDirection = WorkSheetViewFreezeDirection.TOP;
	}
}
